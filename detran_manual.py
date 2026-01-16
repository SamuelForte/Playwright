import time  # Importa a biblioteca para pausas (sleep)
import re  # Importa a biblioteca de expressões regulares para buscas flexíveis
import csv  # Importa a biblioteca para manipulação de arquivos CSV
from datetime import datetime  # Importa para registrar a data e hora da consulta
from playwright.sync_api import sync_playwright, TimeoutError  # Importa as ferramentas de automação do navegador

# ================= CONFIGURAÇÕES =================

URL = "https://sistemas.detran.ce.gov.br/central"  # Define o endereço do site do DETRAN-CE
CSV_ARQUIVO = "_temp_detran.csv"  # Nome do arquivo temporário onde os dados serão salvos

VEICULOS = [  # Lista de dicionários contendo os dados dos carros
    {"placa": "SBA7F09", "renavam": "01365705622"},  # Dados do veículo 1
    {"placa": "TIF1J98", "renavam": "01450499292"},  # Dados do veículo 2
]

# Definição de padrões de busca para botões (ignora maiúsculas/minúsculas)
REGEX_BOTAO_CONSULTAR = re.compile("consultar|confirmar|pesquisar", re.I)  # Padrão para botões de busca
REGEX_BOTAO_FECHAR = re.compile("fechar", re.I)  # Padrão para botões de fechar popups
REGEX_BOTAO_EMITIR = re.compile("emitir", re.I)  # Padrão para botões de emissão de boletos
REGEX_CLIQUE_AQUI = re.compile("clique aqui", re.I)  # Padrão para links de detalhes

# ================= UTILIDADES =================

def log(msg: str):  # Função simples para exibir mensagens no terminal
    print(msg)  # Imprime a mensagem enviada como argumento


def detectar_pendencias(texto: str) -> dict:  # Função que analisa o texto da página
    texto = texto.lower()  # Converte todo o texto para minúsculo para facilitar a busca
    resultado = {  # Dicionário inicial com valores padrão (nada encontrado)
        "multas": 0,  # Contador de multas
        "ipva": False,  # Status do IPVA
        "licenciamento": False,  # Status do Licenciamento
        "motivos_multas": []  # Lista para armazenar os motivos das multas
    }
    match = re.search(r"possui\s+(\d+)\s+multa", texto)  # Procura o padrão "possui X multas"
    if match:  # Se encontrar o padrão acima
        resultado["multas"] = int(match.group(1))  # Extrai o número e salva no dicionário
    if "emita aqui seu ipva" in texto or "débito de ipva" in texto:  # Verifica termos de IPVA
        resultado["ipva"] = True  # Marca como pendente se achar o texto
    if "imprimir seu licenciamento" in texto:  # Verifica termos de licenciamento
        resultado["licenciamento"] = True  # Marca como pendente se achar o texto
    return resultado  # Retorna o dicionário preenchido


def salvar_csv(dados: dict):  # Função para gravar os dados em planilha
    arquivo_existe = False  # Variável de controle para saber se o arquivo já existe
    try:  # Tenta abrir o arquivo para leitura
        with open(CSV_ARQUIVO, "r", encoding="utf-8"):  # Abre o arquivo CSV
            arquivo_existe = True  # Se abriu, o arquivo já existe
    except FileNotFoundError:  # Se der erro de arquivo não encontrado
        pass  # Não faz nada, a variável continua como False
    with open(CSV_ARQUIVO, "a", newline="", encoding="utf-8") as f:  # Abre o arquivo no modo 'anexar' (append)
        writer = csv.writer(f)  # Cria o objeto que escreve no CSV
        if not arquivo_existe:  # Se for um arquivo novo
            writer.writerow(["data_hora", "placa", "renavam", "quantidade_multas", "ipva", "licenciamento", "motivos_multas"])  # Escreve o cabeçalho
        
        # Formata os motivos das multas em uma string separada por | 
        motivos_str = " | ".join(dados.get("motivos_multas", [])) if dados.get("motivos_multas") else "Nenhuma"
        
        writer.writerow([  # Escreve a linha de dados do veículo atual
            datetime.now().strftime("%Y-%m-%d %H:%M:%S"),  # Data e hora atual
            dados["placa"],  # Placa do carro
            dados["renavam"],  # Renavam do carro
            dados["multas"],  # Quantidade de multas achadas
            "SIM" if dados["ipva"] else "NÃO",  # Converte Booleano para SIM/NÃO
            "SIM" if dados["licenciamento"] else "NÃO",  # Converte Booleano para SIM/NÃO
            motivos_str  # Motivos das multas separados por |
        ])

# ================= AÇÕES NA TELA =================

def fechar_popup(page):  # Função para tentar fechar anúncios ou avisos
    try:  # Tenta realizar a ação
        page.get_by_role("button", name=REGEX_BOTAO_FECHAR).click(timeout=3000)  # Clica no botão 'Fechar' se ele aparecer em 3s
    except:  # Se não encontrar o botão
        pass  # Segue a vida sem erro


def acessar_taxas_multas(page):  # Função para navegar no menu lateral
    page.get_by_text("Taxas / Multas", exact=False).click(timeout=15000)  # Clica na opção de multas no menu


def preencher_dados(page, placa, renavam):  # Função para inserir dados no formulário
    campo_placa = page.locator('input[placeholder*="Placa"], input[id*="placa"]').first  # Localiza o campo de placa
    campo_renavam = page.locator('input[placeholder*="Renavam"], input[id*="renavam"]').first  # Localiza o campo de renavam
    campo_placa.wait_for(state="visible", timeout=20000)  # Espera o campo da placa ficar visível
    campo_renavam.wait_for(state="visible", timeout=20000)  # Espera o campo do renavam ficar visível
    campo_placa.type(placa, delay=80)  # Digita a placa letra por letra com atraso de 80ms
    campo_renavam.type(renavam, delay=80)  # Digita o renavam com atraso de 80ms


def clicar_consultar(page):  # Função para enviar o formulário
    page.get_by_role("button", name=REGEX_BOTAO_CONSULTAR).click(timeout=10000)  # Clica no botão de busca


def abrir_detalhe_multas(page):  # Função para ver a lista de multas
    try:  # Tenta realizar a ação
        page.get_by_text(REGEX_CLIQUE_AQUI).first.click(timeout=8000)  # Clica no link 'Clique aqui' para ver detalhes
        log("🔍 Tela de emissão de multas aberta")  # Informa sucesso no log
        return True  # Retorna sucesso
    except:  # Se der erro (ex: link não existe)
        log("⚠️ Não foi possível abrir o detalhe das multas")  # Informa falha no log
        return False  # Retorna falha


def marcar_checkboxes_multas(page) -> bool:  # Função para selecionar as multas na tabela
    try:  # Tenta realizar a ação
        checkboxes = page.locator('table input[type="checkbox"]')  # Localiza todos os checkboxes da tabela
        total = checkboxes.count()  # Conta quantos foram encontrados
        if total == 0:  # Se não houver nenhum
            log("⚠️ Nenhum checkbox encontrado")  # Avisa no log
            return False  # Retorna falha
        for i in range(total):  # Percorre cada checkbox encontrado
            cb = checkboxes.nth(i)  # Pega o checkbox na posição 'i'
            if cb.is_visible() and not cb.is_checked():  # Se estiver visível e ainda não marcado
                cb.check(force=True)  # Marca o checkbox forçando o clique
                log(f"☑️ Multa {i + 1} marcada")  # Avisa qual foi marcada
        return True  # Retorna sucesso após marcar todos
    except Exception as e:  # Se der erro no processo
        log(f"❌ Erro ao marcar multas: {e}")  # Mostra o erro no log
        return False  # Retorna falha


def clicar_emitir(page):  # Função para gerar o boleto
    try:  # Tenta realizar a ação
        botao = page.get_by_role("button", name=REGEX_BOTAO_EMITIR)  # Localiza o botão 'Emitir'
        botao.wait_for(state="visible", timeout=15000)  # Espera ele aparecer por até 15s
        botao.click(force=True)  # Clica no botão
        log("🧾 Botão EMITIR clicado")  # Informa no log
        time.sleep(5)  # Aguarda 5 segundos para o site gerar o PDF/boleto
    except Exception as e:  # Caso o botão não seja clicável
        log(f"❌ Erro ao clicar em Emitir: {e}")  # Mostra o erro no log


# ================= FLUXO PRINCIPAL =================

def processar_veiculo(browser, veiculo: dict, indice: int):  # Função que coordena a consulta de um carro
    log("\n" + "=" * 50)  # Linha divisória no terminal
    log(f"🚗 CONSULTA {indice}")  # Mostra o número da consulta atual
    log(f"Placa: {veiculo['placa']}")  # Mostra a placa sendo processada
    log(f"Renavam: {veiculo['renavam']}")  # Mostra o renavam sendo processado

    context = browser.new_context()  # Cria um novo contexto (limpa cookies e cache)
    page = context.new_page()  # Abre uma nova aba no navegador

    try:  # Inicia o bloco de navegação segura
        log("🌐 Acessando DETRAN...")  # Informa o início do acesso
        page.goto(URL, wait_until="domcontentloaded", timeout=30000)  # Navega até a URL do DETRAN

        fechar_popup(page)  # Tenta fechar avisos iniciais
        acessar_taxas_multas(page)  # Clica na seção de taxas
        preencher_dados(page, veiculo["placa"], veiculo["renavam"])  # Digita placa e renavam
        clicar_consultar(page)  # Clica no botão de busca

        time.sleep(4)  # Espera 4 segundos para a página carregar os resultados

        texto = page.locator("body").inner_text()  # Captura todo o texto visível da página
        resultado = detectar_pendencias(texto)  # Analisa o texto para ver o que o carro deve

        log("\n📄 RESULTADO")  # Cabeçalho de resultado no log
        if resultado["multas"] == 0 and not resultado["ipva"] and not resultado["licenciamento"]:  # Se tudo estiver zerado
            log("✅ NÃO POSSUI PENDÊNCIAS")  # Informa que está limpo
        else:  # Se houver algo pendente
            log("⚠️ POSSUI PENDÊNCIAS")  # Avisa que tem dívidas
            if resultado["multas"] > 0:  # Se o problema for multa
                log(f" - Multas: {resultado['multas']}")  # Mostra a quantidade
                if abrir_detalhe_multas(page):  # Tenta abrir a tela de emissão
                    time.sleep(4)  # Espera a tela carregar
                    if marcar_checkboxes_multas(page):  # Tenta marcar as multas
                        clicar_emitir(page)  # Tenta clicar no botão de pagar/emitir
            if resultado["ipva"]:  # Se o IPVA estiver atrasado
                log(" - IPVA em débito")  # Informa no log
            if resultado["licenciamento"]:  # Se o licenciamento estiver atrasado
                log(" - Licenciamento pendente")  # Informa no log

        salvar_csv({  # Salva as informações coletadas no arquivo
            "placa": veiculo["placa"],  # Placa consultada
            "renavam": veiculo["renavam"],  # Renavam consultado
            **resultado  # Adiciona os resultados (multas, ipva, lic.)
        })

    except TimeoutError:  # Caso o site demore demais para responder
        log("❌ Timeout — site não respondeu")  # Informa erro de tempo
    except Exception as e:  # Qualquer outro erro inesperado
        log(f"❌ Erro geral: {e}")  # Informa o erro ocorrido
    finally:  # Sempre executa ao final, com erro ou não
        page.close()  # Fecha a aba atual
        context.close()  # Fecha o contexto de navegação


def main():  # Função de entrada do programa
    with sync_playwright() as p:  # Inicia o Playwright
        browser = p.chromium.launch(  # Lança o navegador Chromium
            headless=False,  # Abre o navegador visualmente (False) para você ver o processo
            args=[  # Argumentos adicionais
                "--disable-blink-features=AutomationControlled",  # Tenta evitar detecção como robô
                "--start-maximized"  # Inicia o navegador com janela maximizada
            ]
        )

        log(f"📋 {len(VEICULOS)} veículos configurados")  # Informa quantos carros serão olhados

        for i, veiculo in enumerate(VEICULOS, start=1):  # Loop para cada veículo na lista
            processar_veiculo(browser, veiculo, i)  # Executa a função de processamento definida acima
            espera = 20 if i == 1 else 35  # Define um tempo de espera (maior após o primeiro para evitar bloqueio)
            log(f"\n⏳ Aguardando {espera}s para próxima consulta...")  # Avisa sobre a pausa
            time.sleep(espera)  # Faz a pausa obrigatória

        log("\n🏁 TODAS AS CONSULTAS FINALIZADAS")  # Finaliza o log
        log(f"📁 CSV gerado: {CSV_ARQUIVO}")  # Informa o local do arquivo gerado
        input("Pressione ENTER para fechar...")  # Mantém o navegador aberto até você dar Enter
        browser.close()  # Fecha o navegador por completo


if __name__ == "__main__":  # Verifica se o script está sendo rodado diretamente
    main()  # Chama a função principal para começar tudo