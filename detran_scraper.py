import time
import requests
from bs4 import BeautifulSoup

BASE = "https://sistemas.detran.ce.gov.br/central"

ORGAOS_RENAINF = {
    # Federal
    "100":    "SENATRAN / DENATRAN (Federal)",
    "300":    "DNIT – Dept. Nacional de Infraestrutura de Transportes",
    # Ceará – estadual
    "106100": "DETRAN-CE / CIRETRAN",
    "106200": "DER-CE – Dept. de Estradas de Rodagem",
    "212470": "CIRETRAN – Circunscrição Regional CE",
    "214470": "BPRE – Batalhão de Polícia Rodoviária Estadual CE",
    # Fortaleza – municipal
    "213890": "AMC – Autarquia Municipal de Trânsito (Fortaleza)",
    "215850": "ETUFOR / STTU – Trânsito Municipal Fortaleza",
    # Municípios do Ceará
    "213190": "Prefeitura de Apuiares – CE",
    "215370": "Prefeitura de Reriutaba – CE",
}


def _get_form_csrf(session: requests.Session) -> str:
    """
    Carrega o formulário de login via AJAX (como o JS faz ao clicar em Taxas/Multas)
    e extrai o token CSRF do campo hidden 'authenticity_token'.
    """
    r = session.get(
        f"{BASE}/veiculos/detalhamento_servico",
        params={"codigo": "0"},
        headers={"Referer": BASE, "X-Requested-With": "XMLHttpRequest"},
        timeout=15,
    )
    if r.status_code != 200:
        raise RuntimeError(f"Erro ao carregar formulário de login: HTTP {r.status_code}")
    soup = BeautifulSoup(r.text, "html.parser")
    token_input = soup.find("input", {"name": "authenticity_token"})
    if token_input and token_input.get("value"):
        return token_input["value"]
    raise RuntimeError("Token CSRF não encontrado no formulário de login")


def consultar_multas(placa: str, renavam: str) -> dict:
    """
    Consulta multas de um veículo no portal DETRAN-CE.

    Retorna:
        {
            "placa": str,
            "status_multas": str,
            "total_multas": int,
            "erro": str | None,
            "multas": [
                {
                    "ait", "ait_originaria", "motivo",
                    "data_infracao", "data_vencimento",
                    "valor_original", "valor_a_pagar", "valor_numerico",
                    "id_orgao", "orgao", "cod_infracao_ctb",
                }
            ]
        }
    """
    session = requests.Session()
    session.headers.update({
        "User-Agent": (
            "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
            "AppleWebKit/537.36 (KHTML, like Gecko) "
            "Chrome/124.0.0.0 Safari/537.36"
        ),
    })

    # Passo 1 — estabelecer sessão e carregar form de login
    session.get(BASE, timeout=15)
    csrf = _get_form_csrf(session)

    # Passo 2 — autenticar
    r = session.post(
        f"{BASE}/veiculos/login",
        data={
            "authenticity_token": csrf,
            "veiculo[tipo_formulario]": "1",
            "veiculo[placa]": placa.upper().strip(),
            "veiculo[renavam_chassi]": renavam.strip(),
        },
        headers={"Referer": f"{BASE}/veiculos/detalhamento_servico?codigo=0"},
        allow_redirects=True,
        timeout=20,
    )

    # Login retorna JSON: {"status":"succ","errors":""} ou {"status":"error","errors":"msg"}
    try:
        json_resp = r.json()
        if json_resp.get("status") != "succ":
            erro = json_resp.get("errors") or "Placa ou RENAVAM inválidos"
            return {"placa": placa.upper(), "erro": erro, "multas": [], "total_multas": 0}
    except Exception:
        # Fallback: verifica URL se não for JSON
        if "veiculos/login" in r.url or r.status_code >= 400:
            return {"placa": placa.upper(), "erro": "Falha na autenticação", "multas": [], "total_multas": 0}

    # Passo 3 — buscar multas
    r_multas = session.get(f"{BASE}/veiculos/multas", timeout=20)
    multas = _parsear_multas(r_multas.text)

    return {
        "placa": placa.upper(),
        "erro": None,
        "status_multas": "",
        "total_multas": len(multas),
        "multas": multas,
    }


def _parsear_multas(html: str) -> list:
    soup = BeautifulSoup(html, "html.parser")
    multas = []

    for row in soup.select("table tr"):
        cells = [td.get_text(strip=True) for td in row.find_all("td")]
        checkbox = row.find("input", {"type": "checkbox", "class": "multas"})

        if not checkbox or len(cells) < 8:
            continue

        valor_raw = checkbox.get("value", "")
        # Ignora linhas placeholder (ex: value="***" = sem multas reais)
        if not valor_raw or set(valor_raw) == {"*"}:
            continue

        partes = valor_raw.split("*")
        id_orgao = partes[0] if partes else "?"

        multas.append({
            "ait":              cells[1],
            "ait_originaria":   cells[2],
            "motivo":           cells[3],
            "data_infracao":    cells[4],
            "data_vencimento":  cells[5],
            "valor_original":   cells[6],
            "valor_a_pagar":    cells[7],
            "valor_numerico":   float(checkbox.get("data-valor", 0) or 0),
            "id_orgao":         id_orgao,
            "orgao":            ORGAOS_RENAINF.get(id_orgao, f"Órgão cod. {id_orgao}"),
            "cod_infracao_ctb": partes[2] if len(partes) > 2 else "",
        })

    return multas


if __name__ == "__main__":
    resultado = consultar_multas("SBQ0F39", "8AJBA3FS3R0346149")
    if resultado.get("erro"):
        print(f"Erro: {resultado['erro']}")
    else:
        print(f"Total: {resultado['total_multas']} multas")
        for m in resultado["multas"]:
            print(
                f"  {m['ait']} | {m['orgao']:<45} | "
                f"{m['motivo'][:35]:<35} | {m['valor_a_pagar']}"
            )
