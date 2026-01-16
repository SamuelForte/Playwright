import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
import os
import re

# Ler o CSV usando pandas diretamente
csv_file = "_temp_detran.csv"
df = pd.read_csv(csv_file, encoding='utf-8')

print(f"DataFrame shape: {df.shape}")
print(f"Colunas: {list(df.columns)}")
print(f"Primeiras 3 linhas:")
print(df.head(3))

# Remover colunas vazias se houver
df = df.dropna(axis=1, how='all')

# Se há uma coluna de motivos_multas, vamos separá-la em linhas individuais
excel_file = "resultado_detran_organizado.xlsx"

# Criar um novo DataFrame organizado
linhas_novos = []

for idx, row in df.iterrows():
    data_hora = row['data_hora']
    placa = row['placa']
    renavam = row['renavam']
    quantidade_multas = row['quantidade_multas']
    ipva = row['ipva']
    licenciamento = row['licenciamento']
    valor_total = row.get('valor_total_multas', '')
    
    # Se existir coluna de motivos
    if 'motivos_multas' in df.columns:
        motivos = row.get('motivos_multas', '')
        if pd.notna(motivos) and motivos and str(motivos).strip():
            # Separar os motivos por |
            motivos_lista = str(motivos).split(' | ')
            for motivo in motivos_lista:
                motivo = motivo.strip()
                
                # Pular linha TOTAL
                if 'TOTAL R$' in motivo or not motivo:
                    continue
                
                # Extrair informações do motivo usando regex
                # Padrão: AIT -- DESCRIÇÃO DATA_INFRACAO VENCIMENTO VALOR_ORIGINAL VALOR_DESCONTO
                # Exemplo: V607910965 -- TRANSITAR EM... 06/11/2025 30/01/2026 R$ 130,16 R$ 104,13
                
                ait = ''
                descricao = ''
                data_infracao = ''
                vencimento = ''
                valor_original = ''
                valor_desconto = ''
                
                # Extrair AIT (letras + números no início ou após espaço antes de --)
                match_ait = re.search(r'([A-Z]\d+)\s*--', motivo)
                if match_ait:
                    ait = match_ait.group(1)
                
                # Extrair datas (formato dd/mm/yyyy)
                datas = re.findall(r'(\d{2}/\d{2}/\d{4})', motivo)
                if len(datas) >= 2:
                    data_infracao = datas[0]
                    vencimento = datas[1]
                elif len(datas) == 1:
                    data_infracao = datas[0]
                
                # Extrair valores (R$ xxx,xx)
                valores = re.findall(r'R\$\s*([\d.,]+)', motivo)
                if len(valores) >= 2:
                    valor_original = f"R$ {valores[-2]}"
                    valor_desconto = f"R$ {valores[-1]}"
                elif len(valores) == 1:
                    valor_desconto = f"R$ {valores[0]}"
                
                # Extrair descrição (entre -- e a primeira data, ou até encontrar padrão de data)
                match_desc = re.search(r'--\s*(.+?)\s+\d{2}/\d{2}/\d{4}', motivo)
                if match_desc:
                    descricao = match_desc.group(1).strip()
                else:
                    # Se não encontrar o padrão com data, pegar tudo após --
                    match_desc2 = re.search(r'--\s*(.+?)(?:\s+R\$|\s+\d{2}/\d{2}/\d{4}|$)', motivo)
                    if match_desc2:
                        descricao = match_desc2.group(1).strip()
                    else:
                        # Último recurso: pegar tudo após --
                        match_desc3 = re.search(r'--\s*(.+)', motivo)
                        if match_desc3:
                            descricao = match_desc3.group(1).strip()
                        else:
                            descricao = motivo
                
                linhas_novos.append({
                    'Data/Hora Consulta': data_hora,
                    'Placa': placa,
                    'Renavam': renavam,
                    'AIT': ait,
                    'Descrição da Infração': descricao,
                    'Data Infração': data_infracao,
                    'Vencimento': vencimento,
                    'Valor Original': valor_original,
                    'Valor com Desconto': valor_desconto,
                    'IPVA Pendente': ipva,
                    'Licenciamento Pendente': licenciamento
                })
        else:
            linhas_novos.append({
                'Data/Hora Consulta': data_hora,
                'Placa': placa,
                'Renavam': renavam,
                'AIT': '-',
                'Descrição da Infração': 'Sem multas',
                'Data Infração': '-',
                'Vencimento': '-',
                'Valor Original': '-',
                'Valor com Desconto': '-',
                'IPVA Pendente': ipva,
                'Licenciamento Pendente': licenciamento
            })
    else:
        linhas_novos.append({
            'Data/Hora Consulta': data_hora,
            'Placa': placa,
            'Renavam': renavam,
            'AIT': '-',
            'Descrição da Infração': 'Sem informações',
            'Data Infração': '-',
            'Vencimento': '-',
            'Valor Original': '-',
            'Valor com Desconto': '-',
            'IPVA Pendente': ipva,
            'Licenciamento Pendente': licenciamento
        })

df_novo = pd.DataFrame(linhas_novos)

# Salvar em Excel
df_novo.to_excel(excel_file, sheet_name='Resultado DETRAN', index=False)

# Formatação no Excel
wb = openpyxl.load_workbook(excel_file)
ws = wb.active

# Definir estilos
header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
header_font = Font(bold=True, color="FFFFFF", size=11)
center_alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
left_alignment = Alignment(horizontal="left", vertical="top", wrap_text=True)
border = Border(
    left=Side(style='thin'),
    right=Side(style='thin'),
    top=Side(style='thin'),
    bottom=Side(style='thin')
)

# Aplicar formatação ao cabeçalho
for cell in ws[1]:
    cell.fill = header_fill
    cell.font = header_font
    cell.alignment = center_alignment
    cell.border = border

# Aplicar formatação às linhas
for row in ws.iter_rows(min_row=2, max_row=ws.max_row, min_col=1, max_col=ws.max_column):
    for cell in row:
        cell.border = border
        # Descrição da Infração alinhada à esquerda
        if cell.column == 5:  # Coluna E - Descrição
            cell.alignment = left_alignment
        else:
            cell.alignment = center_alignment

# Ajustar largura das colunas
column_widths = {
    'A': 18,  # Data/Hora Consulta
    'B': 12,  # Placa
    'C': 15,  # Renavam
    'D': 14,  # AIT
    'E': 50,  # Descrição da Infração
    'F': 14,  # Data Infração
    'G': 14,  # Vencimento
    'H': 16,  # Valor Original
    'I': 16,  # Valor com Desconto
    'J': 16,  # IPVA Pendente
    'K': 20   # Licenciamento Pendente
}

for col, width in column_widths.items():
    ws.column_dimensions[col].width = width

# Congelar a primeira linha
ws.freeze_panes = "A2"

# Salvar
wb.save(excel_file)

print(f"\n✅ Arquivo Excel organizado: {excel_file}")
print(f"📊 Total de registros: {len(df_novo)}")
print(f"🚗 Veículos únicos: {df_novo['Placa'].nunique()}")
print(f"\n📝 Colunas criadas:")
for col in df_novo.columns:
    print(f"   - {col}")

# Deletar arquivo temporário
if os.path.exists(csv_file):
    os.remove(csv_file)
    print(f"\n🗑️  Arquivo temporário removido: {csv_file}")

