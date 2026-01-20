import time
import re

import os
from datetime import datetime
from playwright.sync_api import sync_playwright, TimeoutError
import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

try:
    import pdfplumber
except ImportError:
    pdfplumber = None

# ================= CONFIGURAÇÕES =================

URL = "https://sistemas.detran.ce.gov.br/central"
EXCEL_ARQUIVO = "resultado_detran_organizado.xlsx"
INTERVALO_ENTRE_CONSULTAS = 2  # segundos - reduzido de 5

VEICULOS = [
    {"placa": "SBA7F09", "renavam": "01365705622"},
    {"placa": "TIF1J98", "renavam": "01450499292"},
    {"placa": "TIF1J99", "renavam": "01450499293"},
    {"placa": "TIF1J93", "renavam": "01450499295"},
    {"placa": "TIF1J93", "renavam": "01450499295"},
    {"placa": "TIF1J93", "renavam": "01450499295"},
    {"placa": "TIF1J93", "renavam": "01450499295"},
]

TIMEOUT_PADRAO = 20000
TIMEOUT_MULTAS = 20000
TIMEOUT_TABELA = 20000

DELAY_SCROLL = 0.2  # reduzido de 0.4
DELAY_CHECKBOX = 0.2  # reduzido de 0.4
DELAY_EMITIR = 2  # reduzido de 4
DELAY_DIGITACAO = 0.1  # reduzido de 0.3

REGEX_BOTAO_CONSULTAR = re.compile("consultar|confirmar|pesquisar", re.I)
REGEX_BOTAO_FECHAR = re.compile("fechar", re.I)
REGEX_BOTAO_EMITIR = re.compile("emitir", re.I)
REGEX_CLIQUE_AQUI = re.compile("clique aqui", re.I)
REGEX_VALOR = re.compile(r"R\$[\s]*([\d.,]+)")
REGEX_MULTAS = re.compile(r"possui\s+(\d+)\s+multa", re.I)

# ================= UTIL =================

def log(msg):
    print(msg)

def formatar_valor_br(valor):
    return f"{valor:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")

# ================= FORM =================

def preencher_dados(page, placa, renavam):
    """Preenche placa e renavam com delay entre caracteres"""
    campo_placa = page.locator('input[placeholder*="Placa" i]')
    campo_renavam = page.locator('input[placeholder*="Renavam" i]')
    
    # Limpa e preenche placa com delay
    campo_placa.click(force=True)
    page.keyboard.press("Control+A")
    page.keyboard.press("Backspace")
    for char in placa:
        page.keyboard.press(char)
        time.sleep(DELAY_DIGITACAO)
    
    # Limpa e preenche renavam com delay
    campo_renavam.click(force=True)
    page.keyboard.press("Control+A")
    page.keyboard.press("Backspace")
    for char in renavam:
        page.keyboard.press(char)
        time.sleep(DELAY_DIGITACAO)

# ================= AÇÕES =================

def fechar_popup(page):
    try:
        page.get_by_role("button", name=REGEX_BOTAO_FECHAR).click(timeout=3000)
    except:
        pass

def acessar_taxas_multas(page):
    page.get_by_text("Taxas / Multas", exact=False).click()

def clicar_consultar(page):
    with page.expect_navigation(wait_until="networkidle"):
        page.get_by_role("button", name=REGEX_BOTAO_CONSULTAR).click()

# ================= MULTAS =================

def abrir_detalhe_multas(page):
    page.get_by_text(REGEX_CLIQUE_AQUI).first.wait_for(timeout=TIMEOUT_MULTAS)
    page.get_by_text(REGEX_CLIQUE_AQUI).first.click()
    page.wait_for_load_state("networkidle")
    log("🔍 Tela de multas aberta")

def extrair_valor(texto):
    valores = REGEX_VALOR.findall(texto)
    if valores:
        return float(valores[-1].replace(".", "").replace(",", "."))
    return 0.0

def processar_multas(page):
    tabela = page.locator("table")
    tabela.wait_for(timeout=TIMEOUT_TABELA)

    linhas = tabela.locator("tbody tr")
    qtd = linhas.count()

    indices_validos = []
    total = 0.0
    motivos = []

    for i in range(qtd):
        linha = linhas.nth(i)
        texto = linha.inner_text().replace("\n", " ")
        valor = extrair_valor(texto)

        if valor > 0:
            indices_validos.append(i)
            total += valor
            motivos.append(texto)
            log(f"📝 Multa válida linha {i} → R$ {valor:.2f}")

    log(f"💰 Total calculado: R$ {formatar_valor_br(total)}")
    return motivos, total, indices_validos

# ================= SELEÇÃO CORRETA DAS MULTAS =================

def marcar_checkboxes_multas(page, indices):
    tabela = page.locator("table")
    linhas = tabela.locator("tbody tr")

    marcadas = 0

    for i in indices:
        linha = linhas.nth(i)
        linha.scroll_into_view_if_needed()
        time.sleep(DELAY_SCROLL)

        try:
            # 🔥 CLICA NO ELEMENTO REAL DO CHECKBOX (Material UI)
            checkbox = linha.locator(
                'mat-checkbox label, mat-checkbox span, input[type="checkbox"]'
            ).first

            checkbox.click(force=True)
            time.sleep(DELAY_CHECKBOX)
            marcadas += 1
            log(f"☑️ Multa {marcadas} selecionada (linha {i})")

        except Exception as e:
            log(f"⚠️ Falha ao marcar linha {i}: {e}")

    log(f"✅ {marcadas} multas selecionadas com sucesso")

def extrair_codigo_pix(page):
    """Extrai o código de pagamento PIX da página antes de emitir."""
    try:
        # Procura pelo botão com onclick="copiarParaClipboard('pix-multas')"
        # ou similar e extrai o valor associado
        
        # Tenta encontrar o elemento com o atributo onclick
        elementos = page.locator('[onclick*="pix"]').all() if page.locator('[onclick*="pix"]').count() > 0 else []
        
        if elementos:
            for elem in elementos:
                texto = elem.inner_text() if elem else ""
                log(f"🔍 Elemento PIX encontrado: {texto}")
        
        # Tenta extrair código de pagamento do texto visível
        texto_pagina = page.inner_text("body")
        
        # Procura por padrão de código de pagamento: números separados por espaço
        # Formato típico: 856300000010 041300062027 601302026898 06128693005
        padrao_codigo = r"(\d{12}\s+\d{12}\s+\d{12}\s+\d{11})"
        match = re.search(padrao_codigo, texto_pagina)
        
        if match:
            codigo = match.group(1).strip()
            log(f"💳 Código PIX extraído: {codigo}")
            return codigo
        
        log("⚠️ Código PIX não encontrado na página")
        return "-"
    except Exception as e:
        log(f"⚠️ Erro ao extrair código PIX: {e}")
        return "-"

    log(f"✅ {marcadas} multas selecionadas com sucesso")

def clicar_emitir(page, context, pasta_boletos):
    """Clica em Emitir, espera aparecer o botão Baixar Extrato e baixa o PDF."""
    botao_emitir = page.get_by_role("button", name=REGEX_BOTAO_EMITIR)
    botao_emitir.wait_for(timeout=TIMEOUT_TABELA)

    def salvar_download(download):
        nome_arquivo = download.suggested_filename or f"extrato_{int(time.time())}.pdf"
        caminho_destino = os.path.join(pasta_boletos, nome_arquivo)
        download.save_as(caminho_destino)
        log(f"💾 Boleto salvo via download: {caminho_destino}")
        return caminho_destino

    # 1) Clica em Emitir para revelar o botão "Baixar Extrato"
    botao_emitir.click()
    log("🧾 Emitir clicado")
    page.wait_for_timeout(800)

    # Localiza o botão Baixar Extrato (ou variações) mostrado na imagem
    seletor_baixar = (
        'button:has-text("Baixar Extrato"), a:has-text("Baixar Extrato"), '
        'button:has-text("Baixar"), a:has-text("Baixar"), '
        'button:has-text("Extrato"), a:has-text("Extrato")'
    )
    botao_baixar = page.locator(seletor_baixar).first

    try:
        botao_baixar.wait_for(timeout=20000)
    except Exception:
        log("⚠️ Botão Baixar Extrato não apareceu.")
        return None

    # 2) Clica em Baixar Extrato - isso abre o PDF em nova aba
    botao_baixar.click(force=True)
    log("⬇️ Baixar Extrato clicado")
    page.wait_for_timeout(2000)

    # 3) Captura a nova página/aba que abriu com o PDF
    paginas = context.pages
    pagina_pdf = None
    
    for p in reversed(paginas):
        if "gerar_boleto" in p.url or "pdf" in p.url.lower():
            pagina_pdf = p
            break
    
    if not pagina_pdf:
        log("⚠️ Nenhuma aba PDF encontrada")
        return None
    
    log(f"📄 PDF aberto em nova aba")
    pagina_pdf.wait_for_load_state("load", timeout=15000)
    page.wait_for_timeout(3000)
    
    # 4) Clica no ícone de download no viewer do PDF
    try:
        # Procura especificamente pelo botão "Baixar Extrato" dentro da página
        seletores_download = [
            'button#btn-exibir-extrato',  # ID específico do botão
            'button.btn.btn-success#btn-exibir-extrato',  # Combinação de classe e ID
            'button[id="btn-exibir-extrato"]',  # Seletor alternativo
            'button[aria-label="Fazer download"]',
            'button[aria-label="Download"]',
            '#download',
            'button#download',
            'cr-icon-button#download',
            'button[aria-label*="download" i]',
            'button[title*="download" i]',
            'button[title*="Download" i]',
            '[role="button"][aria-label*="download" i]',
        ]
        
        log("🔍 Procurando botão de download...")
        botao_download_encontrado = False
        
        for seletor in seletores_download:
            try:
                botao = pagina_pdf.locator(seletor).first
                if botao.is_visible(timeout=1000):
                    log(f"✅ Encontrou botão com seletor: {seletor}")
                    botao.click(force=True)
                    log("✅ Clicou no botão de download")
                    botao_download_encontrado = True
                    page.wait_for_timeout(1000)
                    break
            except Exception as e:
                pass
        
        if not botao_download_encontrado:
            log("⚠️ Botão visual não encontrado, tentando Ctrl+S...")
            pagina_pdf.keyboard.press("Control+S")
            page.wait_for_timeout(1500)
        
        # Aguarda o download
        try:
            with pagina_pdf.expect_download(timeout=25000) as download_info:
                page.wait_for_timeout(2000)
            
            download = download_info.value
            nome_arquivo = download.suggested_filename or f"extrato_{int(time.time())}.pdf"
            caminho_destino = os.path.join(pasta_boletos, nome_arquivo)
            download.save_as(caminho_destino)
            log(f"💾 PDF salvo: {caminho_destino}")
            
            pagina_pdf.close()
            return caminho_destino
        except TimeoutError:
            log("⚠️ Timeout esperando download")
            pagina_pdf.close()
            return None
        
    except Exception as e:
        log(f"⚠️ Erro ao tentar baixar PDF: {e}")
        try:
            pagina_pdf.close()
        except:
            pass
        return None

def extrair_dados_do_pdf(caminho_pdf):
    """Extrai código de pagamento, órgão autuador e descrição do PDF."""
    try:
        if not pdfplumber:
            log("⚠️ pdfplumber não está instalado")
            return "-", "-"
        
        # Valida se o arquivo existe e é PDF
        if not os.path.exists(caminho_pdf):
            log(f"⚠️ Arquivo não encontrado: {caminho_pdf}")
            return "-", "-"
        
        with open(caminho_pdf, 'rb') as f:
            header = f.read(10)
            if not header.startswith(b'%PDF'):
                log(f"⚠️ Arquivo {caminho_pdf} não é um PDF válido")
                return "-", "-"
        
        with pdfplumber.open(caminho_pdf) as pdf:
            texto = ""
            linhas = []
            for page in pdf.pages[:2]:  # Lê primeiras 2 páginas (cabeçalho e descrição)
                conteudo = page.extract_text() or ""
                texto += conteudo
                linhas.extend(conteudo.splitlines())

            log("🔎 Prévia do PDF (linhas iniciais):")
            for l in linhas[:8]:
                log(f"   {l}")

            codigo_pagamento = "-"
            descricao_pdf = "-"
            orgao = "-"

            # 1) Extrai código de pagamento - procura por padrão numérico específico
            # Geralmente tem 47 dígitos com barras ou está próximo a "Código de Pagamento"
            for i, linha in enumerate(linhas):
                linha_limpa = linha.strip()
                # Procura por código com muitos dígitos (padrão de boleto: 47 dígitos)
                if re.match(r"^\d{4}\s*\d{4}\s*\d{4}\s*\d{4}", linha_limpa) or \
                   re.match(r"^\d{11}\s*\d{10}\s*\d{10}\s*\d{16}", linha_limpa) or \
                   (len(re.sub(r"\D", "", linha_limpa)) >= 40 and "código" in linhas[i-1].lower() if i > 0 else False):
                    codigo_pagamento = linha_limpa
                    log(f"💳 Código de Pagamento encontrado: {codigo_pagamento}")
                    break
            
            # 2) Extrai órgão autuador - procura especificamente por "Órgão Autuador" ou "Autuador"
            for i, linha in enumerate(linhas):
                linha_low = linha.lower()
                if "órgão" in linha_low and "autua" in linha_low:
                    # A próxima linha com conteúdo deve ser o nome do órgão
                    for proxima in linhas[i+1:]:
                        proxima_limpa = proxima.strip()
                        if proxima_limpa and len(proxima_limpa) > 2:
                            orgao = proxima_limpa
                            log(f"🏢 Órgão Autuador encontrado: {orgao}")
                            break
                    if orgao != "-":
                        break
            
            # Se não encontrou com "Órgão Autuador", tenta procurar por padrões conhecidos
            if orgao == "-":
                # Procura por padrões de órgãos específicos
                padrao_orgaos = [
                    (r"DEMUTRAN\s+([A-Z\s]+?)(?=\n|$)", "DEMUTRAN"),
                    (r"SEMOB", "SEMOB"),
                    (r"POL[IÍ]CIA\s+MILITAR", "PM"),
                    (r"POL[IÍ]CIA\s+FEDERAL", "PF"),
                    (r"POL[IÍ]CIA\s+RODOVI[ÁA]RIA", "PRF"),
                    (r"EMPRESA\s+DE\s+TRANSPORTE", "Transporte"),
                    (r"DEPARTAMENTO\s+ESTADUAL", "DETRAN"),
                    (r"AG[ÊE]NCIA\s+DE\s+TR[ÂA]NSITO", "Trânsito"),
                ]
                
                for pattern, fallback in padrao_orgaos:
                    match = re.search(pattern, texto, re.IGNORECASE)
                    if match:
                        if "DEMUTRAN" in fallback:
                            # Extrai o nome completo do DEMUTRAN
                            orgao = match.group(0).strip()
                        else:
                            orgao = fallback
                        log(f"🏢 Órgão Autuador encontrado (padrão): {orgao}")
                        break
            
            # 3) Extrai descrição: pega a linha logo após "Descrição (Taxa / Multa)"
            for i, linha in enumerate(linhas):
                linha_low = linha.lower()
                if "descri" in linha_low and "taxa" in linha_low:
                    for proxima in linhas[i+1:]:
                        proxima_limpa = proxima.strip()
                        if proxima_limpa:
                            descricao_pdf = proxima_limpa
                            break
                    break

            # 4) Combina código de pagamento + descrição na variável final
            resultado_pdf = descricao_pdf
            if codigo_pagamento != "-":
                if descricao_pdf != "-":
                    resultado_pdf = f"{codigo_pagamento} | {descricao_pdf}"
                else:
                    resultado_pdf = codigo_pagamento

        return orgao, resultado_pdf
    except Exception as e:
        log(f"⚠️ Erro ao ler PDF: {e}")
        return "-", "-"

# ================= PROCESSAMENTO =================

def extrair_pendencias(texto):
    match = REGEX_MULTAS.search(texto)
    return int(match.group(1)) if match else 0

def salvar_no_excel(multas_lista):
    """Salva multas no Excel com formatação"""
    if not multas_lista:
        log("⚠️ Nenhuma multa para salvar")
        return
    
    df_novo = pd.DataFrame(multas_lista)
    
    try:
        # Tenta fechar arquivo se estiver aberto
        import os
        if os.path.exists(EXCEL_ARQUIVO):
            try:
                import gc
                gc.collect()
            except:
                pass
        
        # Salva o novo DataFrame
        df_novo.to_excel(EXCEL_ARQUIVO, index=False, sheet_name="Resultado DETRAN", engine='openpyxl')
    except PermissionError:
        log(f"⚠️ Arquivo {EXCEL_ARQUIVO} está aberto. Feche e tente novamente!")
        return
    except Exception as e:
        log(f"⚠️ Erro ao salvar Excel: {e}")
        return
    
    # Formatar Excel
    try:
        wb = openpyxl.load_workbook(EXCEL_ARQUIVO)
        ws = wb.active
        
        header_fill = PatternFill("solid", fgColor="1F4E78")
        header_font = Font(bold=True, color="FFFFFF")
        center = Alignment(horizontal="center", vertical="center", wrap_text=True)
        left = Alignment(horizontal="left", vertical="top", wrap_text=True)
        border = Border(
            left=Side(style="thin"), right=Side(style="thin"),
            top=Side(style="thin"), bottom=Side(style="thin")
        )
        
        for cell in ws[1]:
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = center
            cell.border = border
        
        for row in ws.iter_rows(min_row=2):
            for cell in row:
                cell.border = border
                cell.alignment = left if cell.column in (5, 11) else center
        
        larguras = {"A": 12, "B": 5, "C": 15, "D": 18, "E": 55, "F": 14, "G": 14, "H": 16, "I": 16, "J": 18, "K": 55}
        for col, w in larguras.items():
            ws.column_dimensions[col].width = w
        
        ws.freeze_panes = "A2"
        wb.save(EXCEL_ARQUIVO)
        log(f"✅ Dados salvos em: {EXCEL_ARQUIVO}")
    except Exception as e:
        log(f"⚠️ Erro ao formatar Excel: {e}")

def processar_veiculo(browser, veiculo, indice):
    log("\n" + "=" * 50)
    log(f"🚗 CONSULTA {indice} | {veiculo['placa']}")

    # Cria pasta de download com data de hoje
    pasta_base = "boletos"
    data_hoje = datetime.now().strftime("%d-%m-%Y")
    pasta_boletos = os.path.join(pasta_base, data_hoje)
    
    if not os.path.exists(pasta_boletos):
        os.makedirs(pasta_boletos)
        log(f"📁 Pasta '{pasta_boletos}' criada")

    context = browser.new_context(
        accept_downloads=True
    )
    page = context.new_page()
    multas_lista = []
    numero_sequencial = 0

    try:
        page.goto(URL)
        fechar_popup(page)
        acessar_taxas_multas(page)
        preencher_dados(page, veiculo["placa"], veiculo["renavam"])
        clicar_consultar(page)

        texto = page.inner_text("body").lower()
        qtd_multas = extrair_pendencias(texto)

        log(f"📄 Multas encontradas: {qtd_multas}")

        total = 0.0
        motivos = []

        if qtd_multas > 0:
            abrir_detalhe_multas(page)
            motivos, total, indices = processar_multas(page)
            
            # Processa cada multa para salvar no Excel
            for motivo in motivos:
                numero_sequencial += 1
                
                # DEBUG: Mostra o texto bruto
                log(f"\n🔍 TEXTO BRUTO MULTA {numero_sequencial}:")
                log(f"  {motivo[:200]}...")
                
                # Extrai AIT
                ait = "-"
                match_ait = re.search(r"([A-Z]{1,3}\d{6,})\s*--", motivo)
                if match_ait:
                    ait = match_ait.group(1)
                
                # Extrai datas
                datas = re.findall(r"\d{2}/\d{2}/\d{4}", motivo)
                data_infracao = datas[0] if len(datas) > 0 else "-"
                vencimento = datas[1] if len(datas) > 1 else "-"
                
                # Extrai valores
                valores = re.findall(r"R\$\s*([\d.,]+)", motivo)
                valor = "-"
                valor_a_pagar = "-"
                if len(valores) == 1:
                    valor = f"R$ {valores[0]}"
                    valor_a_pagar = f"R$ {valores[0]}"
                elif len(valores) >= 2:
                    valor = f"R$ {valores[-2]}"
                    valor_a_pagar = f"R$ {valores[-1]}"
                
                # Extrai descrição - versão SIMPLIFICADA
                # Remove checkbox, AIT, datas e valores
                descricao = motivo
                # Remove checkbox vazio no início
                descricao = re.sub(r"^\s*\□?\s*", "", descricao)
                # Remove AIT
                descricao = re.sub(r"[A-Z]{1,3}\d{6,}\s*--\s*", "", descricao)
                # Remove datas
                descricao = re.sub(r"\d{2}/\d{2}/\d{4}", "", descricao)
                # Remove valores
                descricao = re.sub(r"R\$\s*[\d.,]+", "", descricao)
                # Remove espaços extras
                descricao = re.sub(r"\s+", " ", descricao).strip()
                
                if not descricao:
                    descricao = "-"
                
                # Exibe informações da multa
                log(f"\n✏️ MULTA {numero_sequencial}")
                log(f"  AIT: {ait}")
                log(f"  📋 Descrição: {descricao}")
                log(f"  📅 Data: {data_infracao} | Vencimento: {vencimento}")
                log(f"  💰 Valor: {valor} → A Pagar: {valor_a_pagar}")
                
                multas_lista.append({
                    "Placa": veiculo["placa"],
                    "#": numero_sequencial,
                    "AIT": ait,
                    "AIT Originária": "-",
                    "Motivo": descricao,
                    "Data Infração": data_infracao,
                    "Data Vencimento": vencimento,
                    "Valor": valor,
                    "Valor a Pagar": valor_a_pagar,
                    "Órgão Autuador": "-",
                    "Código de pagamento em barra": "-"
                })
            
            marcar_checkboxes_multas(page, indices)
            
            # Extrai o código PIX ANTES de emitir
            codigo_pix = extrair_codigo_pix(page)
            
            # Emite, baixa o PDF e extrai dados
            orgao_autuador = "-"
            descricao_pdf = "-"
            caminho_pdf = clicar_emitir(page, context, pasta_boletos)
            if caminho_pdf:
                orgao_autuador, descricao_pdf = extrair_dados_do_pdf(caminho_pdf)
                log(f"🏢 Órgão Autuador: {orgao_autuador}")
                log(f"📄 Descrição PDF: {descricao_pdf}")

            # Adiciona código PIX na descrição se encontrou
            if codigo_pix != "-":
                if descricao_pdf != "-":
                    descricao_pdf = f"{codigo_pix} | {descricao_pdf}"
                else:
                    descricao_pdf = codigo_pix

            # Adiciona dados a todas as multas processadas
            for multa in multas_lista:
                multa["Órgão Autuador"] = orgao_autuador
                multa["Código de pagamento em barra"] = descricao_pdf
        
        return total, multas_lista

    except TimeoutError:
        log("❌ Timeout")
        return 0.0, []
    finally:
        page.close()
        context.close()

# ================= MAIN =================

def main():
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=False)
        total_geral = 0.0
        todas_multas = []

        for i, v in enumerate(VEICULOS, 1):
            total, multas = processar_veiculo(browser, v, i)
            total_geral += total
            todas_multas.extend(multas)
            if i < len(VEICULOS):
                time.sleep(INTERVALO_ENTRE_CONSULTAS)

        # Salva todas as multas no Excel
        if todas_multas:
            salvar_no_excel(todas_multas)

        log(f"\n💵 TOTAL GERAL: R$ {formatar_valor_br(total_geral)}")
        input("ENTER para sair...")
        browser.close()

if __name__ == "__main__":
    main()
