# Extração de Dados dos PDFs de Multas

## Resumo das Mudanças

Foram implementadas melhorias na extração de dados dos PDFs de multas para garantir que as informações corretas sejam capturadas e preenchidas no arquivo `resultado_detran_organizado.xlsx`.

### Campos Atualizados

1. **Data Infração** - Data da ocorrência da infração de trânsito
2. **Data Vencimento** - Data de vencimento do pagamento
3. **Órgão Autuador** - Órgão responsável pela autuação (DETRAN-CE, DEMUTRAN RUSSAS, etc)
4. **Código de pagamento em barra** - Código de barras para pagamento + descrição

---

## Detalhes Técnicos

### 1. Extração de Data Infração e Vencimento

**Método Principal:**
- Procura pela linha que contém a multa no PDF (padrão: `ÓRGÃO | CÓDIGO | DESCRIÇÃO DATA1 DATA2`)
- A primeira data é a **Data da Infração** (quando ocorreu a infração)
- A segunda data é a **Data de Vencimento** (prazo para pagamento)

**Exemplo:**
```
DETRAN-CE | V607910965 | 07455 | TRANSITAR EM VELOCIDADE 06/11/2025 30/01/2026 130,16 104,13
                                                                ^^^^^^^^^^^ ^^^^^^^^^^^
                                                                Infração   Vencimento
```

**Métodos Alternativos:**
1. Se não encontrar na linha da multa, procura pelo cabeçalho "Data Infração Vencimento"
2. Se falhar, ordena cronologicamente as datas encontradas (filtrando 2020-2030)
3. Usa a primeira data como infração e a última como vencimento

### 2. Extração do Órgão Autuador

**Método Principal:**
- Extrai a palavra antes do primeiro `|` na linha da multa
- Exemplo: `DETRAN-CE | ...` → **DETRAN-CE**
- Exemplo: `DEMUTRAN RUSSAS | ...` → **DEMUTRAN RUSSAS**

**Método Alternativo:**
- Se não encontrar, procura por padrões conhecidos (DETRAN, DEMUTRAN, SEMOB, PM, PF, PRF)

### 3. Extração do Código de Barras

**Método:**
- Procura pela linha contendo exatamente **47-48 dígitos** numéricos
- Ignora linhas com texto adicional
- Preserva os espaços (formato: `12345678 12345678 12345678 12345678`)

**Exemplo:**
```
856300000010 041300062027 601302026898 061286930005
```

---

## Fluxo de Processamento

```
1. PDF é baixado (clicar_emitir)
2. extrair_dados_do_pdf() é chamado
   ├─ Extrai código de barras (47-48 dígitos)
   ├─ Extrai órgão (antes do |)
   ├─ Extrai datas (primeira = infração, segunda = vencimento)
   └─ Retorna: (orgao, descricao_pdf, data_infracao, vencimento)
3. Dados são atualizados nas multas do grupo processado
4. Excel é salvo com informações atualizadas
```

---

## Dependências Adicionadas

```
pdfplumber==0.11.0    # Extração de texto de PDFs
supabase==2.0.0       # Backend services
python-dotenv==1.0.0  # Variáveis de ambiente
```

---

## Testes Validados

### PDF 1: Extrato_6601163057.pdf
- ✅ Órgão: DETRAN-CE
- ✅ Data Infração: 06/11/2025
- ✅ Data Vencimento: 30/01/2026
- ✅ Código Barras: 856300000010 041300062027 601302026898 061286930005

### PDF 2: Extrato_6601204759.pdf
- ✅ Órgão: DEMUTRAN RUSSAS
- ✅ Data Infração: 14/09/2025
- ✅ Data Vencimento: 24/11/2025
- ✅ Código Barras: 856400000175 112500062024 601302026898 074576826007

---

## Resultado Final

Todas as multas no arquivo `resultado_detran_organizado.xlsx` serão preenchidas com:
- **Datas corretas** extraídas dos PDFs
- **Órgão correto** de autuação
- **Código de barras** para pagamento

A atualização ocorre automaticamente após cada emissão de boleto.
