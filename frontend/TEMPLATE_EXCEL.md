# Template de Excel para Importação de Veículos

Este é um modelo de planilha que pode ser usado para importar múltiplos veículos de uma vez.

## Formato Necessário

A planilha deve conter **obrigatoriamente** estas duas colunas:

- `placa` - Placa do veículo (formato AAA0000 ou AAA0A00)
- `renavam` - RENAVAM do veículo (11 dígitos)

## Exemplo de Dados

```
placa   | renavam
--------|-------------
SBA7F09 | 01365705622
TIF1J98 | 01450499292
ABC1234 | 12345678901
XYZ9876 | 98765432109
```

## Como Criar no Excel

1. Abra o Microsoft Excel ou Google Sheets
2. Na primeira linha, coloque os cabeçalhos: `placa` e `renavam`
3. Nas linhas seguintes, adicione os dados dos veículos
4. Salve como `.xlsx` ou `.xls`

## Download

Você pode baixar um template pronto aqui: [template_veiculos.xlsx](./template_veiculos.xlsx)

Ou criar manualmente seguindo o formato acima.

## Validações

O sistema irá validar automaticamente:

✅ Placa no formato correto (Mercosul ou antigo)  
✅ RENAVAM com 11 dígitos  
✅ Campos preenchidos  

❌ Linhas com dados inválidos serão ignoradas

## Importação

1. Acesse **Nova Consulta** no sistema
2. Clique na aba **"Importar Planilha"**
3. Arraste o arquivo Excel ou clique para selecionar
4. Visualize o preview dos veículos encontrados
5. Clique em **"Confirmar Importação"**
6. Os veículos serão adicionados à lista

## Limite

- Não há limite de veículos por arquivo
- Recomendado: até 100 veículos por consulta para melhor performance
- Para consultas maiores, considere dividir em lotes

---

**Dica:** Mantenha uma planilha mestre com todos os veículos da frota e faça consultas periódicas!
