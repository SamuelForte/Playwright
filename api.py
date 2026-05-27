from typing import List, Optional, Any, Dict
from fastapi import FastAPI
from fastapi.middleware.cors import CORSMiddleware
import os
from pydantic import BaseModel
from playwright.sync_api import sync_playwright, TimeoutError

# Importa funções do script existente
from detran_manual import processar_veiculo, salvar_no_excel, formatar_valor_br

app = FastAPI(title="DETRAN-CE API", version="1.0.0")


def _build_allowed_origins() -> List[str]:
    origins = {
        "http://localhost:3000",
        "http://127.0.0.1:3000",
        "https://samuelforte.github.io",
        "https://samuelforte.github.io/Playwright",
    }
    extra = os.getenv("ALLOWED_ORIGINS", "")
    if extra:
        for o in extra.split(","):
            norm = o.strip().rstrip("/")
            if norm:
                origins.add(norm)
    return sorted(origins)


# Habilita CORS para permitir chamadas do GitHub Pages e localhost
allow_origin_regex_env = os.getenv("ALLOWED_ORIGIN_REGEX")
allow_origin_regex_default = r"https://(.*\.railway\.app|.*\.vercel\.app)"

app.add_middleware(
    CORSMiddleware,
    allow_origins=_build_allowed_origins(),
    allow_origin_regex=allow_origin_regex_env or allow_origin_regex_default,
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

class Veiculo(BaseModel):
    placa: str
    renavam: str

class ConsultaRequest(BaseModel):
    veiculo: Veiculo
    salvar_excel: bool = True

class ConsultaLoteRequest(BaseModel):
    veiculos: List[Veiculo]
    salvar_excel: bool = True

@app.get("/health")
def health() -> Dict[str, Any]:
    return {"status": "ok"}

@app.post("/consultar")
def consultar(req: ConsultaRequest) -> Dict[str, Any]:
    try:
        with sync_playwright() as p:
            # Para ambiente local, manter não-headless facilita interações
            browser = p.chromium.launch(headless=False)
            total, multas = processar_veiculo(browser, req.veiculo.dict(), 1)
            browser.close()

        if req.salvar_excel and multas:
            salvar_no_excel(multas)

        return {
            "total": total,
            "total_formatado": f"R$ {formatar_valor_br(total)}",
            "quantidade_multas": len(multas),
            "multas": multas,
        }
    except TimeoutError:
        return {"error": "timeout", "message": "Tempo excedido durante a consulta."}
    except Exception as e:
        return {"error": "internal_error", "message": str(e)}

@app.post("/consultar-lote")
def consultar_lote(req: ConsultaLoteRequest) -> Dict[str, Any]:
    todas_multas: List[Dict[str, Any]] = []
    total_geral = 0.0

    try:
        with sync_playwright() as p:
            browser = p.chromium.launch(headless=False)
            for i, v in enumerate(req.veiculos, 1):
                total, multas = processar_veiculo(browser, v.dict(), i)
                total_geral += total
                todas_multas.extend(multas)
            browser.close()

        if req.salvar_excel and todas_multas:
            salvar_no_excel(todas_multas)

        return {
            "total_geral": total_geral,
            "total_geral_formatado": f"R$ {formatar_valor_br(total_geral)}",
            "quantidade_multas": len(todas_multas),
            "multas": todas_multas,
        }
    except TimeoutError:
        return {"error": "timeout", "message": "Tempo excedido durante a consulta."}
    except Exception as e:
        return {"error": "internal_error", "message": str(e)}

if __name__ == "__main__":
    import uvicorn
    uvicorn.run("api:app", host="0.0.0.0", port=8000, reload=True)
