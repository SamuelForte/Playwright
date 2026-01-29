"""
Script de teste para validar se resultado_detran_organizado.xlsx 
está sendo preenchido corretamente com os dados extraídos dos PDFs.
"""

import pandas as pd
import os

EXCEL_FILE = "resultado_detran_organizado.xlsx"

if not os.path.exists(EXCEL_FILE):
    print(f"❌ Arquivo {EXCEL_FILE} nao encontrado!")
    print("    Execute detran_manual.py primeiro para gerar o arquivo.")
    exit()

try:
    df = pd.read_excel(EXCEL_FILE, sheet_name="Resultado DETRAN")
    print("=" * 80)
    print(f"✅ Arquivo lido com sucesso: {EXCEL_FILE}")
    print("=" * 80)
    
    # Exibe informações sobre o Excel
    print(f"\n📊 INFORMACOES GERAIS")
    print(f"   Total de multas: {len(df)}")
    print(f"   Colunas: {len(df.columns)}")
    
    # Exibe as colunas
    print(f"\n📋 COLUNAS ENCONTRADAS:")
    for i, col in enumerate(df.columns, 1):
        print(f"   {i}. {col}")
    
    # Verifica se as colunas esperadas existem
    print(f"\n✓ VALIDACAO DAS COLUNAS:")
    colunas_esperadas = [
        "Placa", "#", "AIT", "AIT Originária", "Motivo",
        "Data Infração", "Data Vencimento", "Valor", "Valor a Pagar",
        "Órgão Autuador", "Código de pagamento em barra"
    ]
    
    for col_esperada in colunas_esperadas:
        if col_esperada in df.columns:
            print(f"   ✅ {col_esperada}")
        else:
            print(f"   ❌ {col_esperada} (NAO ENCONTRADA)")
    
    # Exibe primeiras linhas
    if len(df) > 0:
        print(f"\n📄 PRIMEIRAS LINHAS:")
        print(df.head(3).to_string())
        
        # Verifica se colunas de dados do PDF estão preenchidas
        print(f"\n🔍 VERIFICACAO DE PREENCHIMENTO:")
        
        orgao_preenchidos = df["Órgão Autuador"].ne("-").sum()
        codigo_preenchidos = df["Código de pagamento em barra"].ne("-").sum()
        datas_infra_preenchidas = df["Data Infração"].ne("-").sum()
        datas_venc_preenchidas = df["Data Vencimento"].ne("-").sum()
        
        print(f"   Órgão Autuador: {orgao_preenchidos}/{len(df)} preenchidos")
        print(f"   Código de pagamento: {codigo_preenchidos}/{len(df)} preenchidos")
        print(f"   Data Infração: {datas_infra_preenchidas}/{len(df)} preenchidos")
        print(f"   Data Vencimento: {datas_venc_preenchidas}/{len(df)} preenchidos")
        
        if orgao_preenchidos > 0 or codigo_preenchidos > 0:
            print(f"\n✅ Dados do PDF estao sendo salvos corretamente!")
        else:
            print(f"\n⚠️  Nenhum dado de PDF foi preenchido.")
            print(f"    Certifique-se de que detran_manual.py extraiu os PDFs.")
    else:
        print("   ⚠️  Nenhuma multa no arquivo.")
        
except Exception as e:
    print(f"❌ Erro ao ler {EXCEL_FILE}: {e}")
    import traceback
    traceback.print_exc()
