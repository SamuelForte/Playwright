"""
Script de Diagnóstico Rápido do Sistema DETRAN
Verifica todos os componentes críticos
"""

import sys
import os

print("=" * 80)
print("DIAGNÓSTICO DO SISTEMA DETRAN")
print("=" * 80)

# 1. Verificar Python
print(f"\n1. PYTHON")
print(f"   Versão: {sys.version}")
print(f"   Executável: {sys.executable}")
venv_esperado = r"C:\Users\Samuel\Downloads\detran-main\.venv\Scripts\python.exe"
if sys.executable.lower() == venv_esperado.lower():
    print(f"   ✅ Executando no venv correto")
else:
    print(f"   ⚠️  Executando FORA do venv!")
    print(f"   Esperado: {venv_esperado}")

# 2. Verificar imports críticos
print(f"\n2. DEPENDÊNCIAS")

dependencias = [
    "pdfplumber",
    "pandas",
    "openpyxl",
    "playwright",
    "fastapi",
    "uvicorn",
    "supabase",
    "python-dotenv"
]

for dep in dependencias:
    try:
        mod = __import__(dep.replace("-", "_"))
        version = getattr(mod, "__version__", "N/A")
        print(f"   ✅ {dep}: {version}")
    except ImportError as e:
        print(f"   ❌ {dep}: NÃO INSTALADO ({e})")

# 3. Verificar arquivos
print(f"\n3. ARQUIVOS")

arquivos = [
    "detran_manual.py",
    "api_server.py",
    "organizar_excel.py",
    "testar_excel.py",
    "resultado_detran_organizado.xlsx"
]

for arq in arquivos:
    if os.path.exists(arq):
        tamanho = os.path.getsize(arq)
        print(f"   ✅ {arq} ({tamanho:,} bytes)")
    else:
        print(f"   ❌ {arq}: NÃO ENCONTRADO")

# 4. Verificar PDFs
print(f"\n4. PDFs BAIXADOS")

pasta_boletos = "boletos"
if os.path.exists(pasta_boletos):
    pastas = [p for p in os.listdir(pasta_boletos) if os.path.isdir(os.path.join(pasta_boletos, p))]
    total_pdfs = 0
    for pasta in sorted(pastas):
        caminho_pasta = os.path.join(pasta_boletos, pasta)
        pdfs = [f for f in os.listdir(caminho_pasta) if f.endswith('.pdf')]
        total_pdfs += len(pdfs)
        print(f"   📁 {pasta}: {len(pdfs)} PDFs")
    print(f"   Total: {total_pdfs} PDFs")
else:
    print(f"   ⚠️  Pasta boletos não encontrada")

# 5. Testar extração de PDF
print(f"\n5. TESTE DE EXTRAÇÃO DE PDF")

try:
    from detran_manual import extrair_dados_do_pdf
    
    # Procura um PDF de teste
    pdf_teste = None
    if os.path.exists("boletos/28-01-2026"):
        pdfs = [f for f in os.listdir("boletos/28-01-2026") if f.endswith('.pdf')]
        if pdfs:
            pdf_teste = os.path.join("boletos/28-01-2026", pdfs[0])
    
    if pdf_teste:
        print(f"   Testando com: {pdf_teste}")
        orgao, desc, data_inf, data_venc = extrair_dados_do_pdf(pdf_teste)
        print(f"   Órgão: {orgao}")
        print(f"   Código: {desc[:50] if len(desc) > 50 else desc}")
        print(f"   Data Infração: {data_inf}")
        print(f"   Data Vencimento: {data_venc}")
        
        if orgao != "-" or data_inf != "-":
            print(f"   ✅ Extração FUNCIONANDO!")
        else:
            print(f"   ❌ Extração NÃO está funcionando!")
    else:
        print(f"   ⚠️  Nenhum PDF de teste encontrado")
        
except Exception as e:
    print(f"   ❌ Erro ao testar extração: {e}")
    import traceback
    traceback.print_exc()

# 6. Verificar Excel
print(f"\n6. EXCEL ATUAL")

try:
    import pandas as pd
    df = pd.read_excel("resultado_detran_organizado.xlsx", sheet_name="Resultado DETRAN")
    print(f"   Total de multas: {len(df)}")
    print(f"   Colunas: {len(df.columns)}")
    
    # Verifica preenchimento
    orgaos = df["Órgão Autuador"].ne("-").sum()
    codigos = df["Código de pagamento em barra"].ne("-").sum()
    
    print(f"   Órgãos preenchidos: {orgaos}/{len(df)}")
    print(f"   Códigos preenchidos: {codigos}/{len(df)}")
    
    if orgaos > 0 and codigos > 0:
        print(f"   ✅ Dados do PDF estão sendo salvos!")
    else:
        print(f"   ❌ Dados do PDF NÃO estão sendo salvos!")
except Exception as e:
    print(f"   ❌ Erro ao ler Excel: {e}")

print("\n" + "=" * 80)
print("FIM DO DIAGNÓSTICO")
print("=" * 80)
