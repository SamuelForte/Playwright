# 🚗 Sistema DETRAN-CE - Consulta de Multas

Sistema completo para consulta automatizada de multas do DETRAN-CE com interface web profissional.

## 📁 Estrutura do Projeto

```
detran-main/
├── 📂 backend/
│   ├── detran_manual.py       # Automação Playwright (consulta DETRAN)
│   ├── api_server.py          # API FastAPI (endpoints REST)
│   ├── requirements.txt       # Dependências Python
│   └── boletos/              # PDFs baixados (por data)
│
└── 📂 frontend/
    ├── src/
    │   ├── app/              # Páginas Next.js
    │   ├── components/       # Componentes React
    │   └── lib/             # Utilitários (API client)
    ├── package.json         # Dependências Node
    └── README.md           # Documentação completa
```

## 🎯 Funcionalidades

### ✅ Backend (Python)
- ✅ Automação com Playwright para consulta no site do DETRAN-CE
- ✅ Extração de dados de multas (AIT, valores, datas, órgão autuador)
- ✅ Download automático de PDFs de boletos
- ✅ Geração de planilha Excel organizada
- ✅ Extração de código de pagamento PIX
- ✅ Organização de arquivos por data

### ✅ Frontend (Next.js)
- ✅ Dashboard com resumo de consultas
- ✅ Cadastro manual de veículos com validação
- ✅ Importação de planilhas Excel/CSV
- ✅ Acompanhamento em tempo real do processamento
- ✅ Visualização de resultados com filtros e busca
- ✅ Download de Excel e PDFs
- ✅ Histórico de consultas
- ✅ Interface responsiva e profissional

## 🚀 Como Usar

### 1️⃣ Instalar Dependências

**Backend:**
```powershell
# Instalar Python 3.8+
pip install -r requirements.txt
playwright install chromium
```

**Frontend:**
```powershell
cd frontend
npm install
```

### 2️⃣ Rodar o Sistema

**Terminal 1 - Backend (API):**
```powershell
python api_server.py
# ou
uvicorn api_server:app --reload --port 8000
```

**Terminal 2 - Frontend:**
```powershell
cd frontend
npm run dev
```

### 3️⃣ Acessar

- **Frontend:** http://localhost:3000
- **API Docs:** http://localhost:8000/docs

## 📖 Documentação Completa

### Frontend
- [📘 README Completo](./frontend/README.md)
- [⚡ Início Rápido](./frontend/INICIO_RAPIDO.md)
- [🔗 Integração Backend](./frontend/INTEGRACAO_BACKEND.md)
- [📋 Comandos Úteis](./frontend/COMANDOS_RAPIDOS.md)
- [🎯 Visão Geral](./frontend/VISAO_COMPLETA.md)
- [📊 Template Excel](./frontend/TEMPLATE_EXCEL.md)

### Backend
- [📘 Como Usar](./COMO_USAR.txt)
- [📝 README Original](./README.md)

## 🎨 Telas do Sistema

1. **Dashboard** - Resumo com cards de estatísticas
2. **Nova Consulta** - Cadastro manual ou importação de Excel
3. **Processamento** - Status em tempo real com polling
4. **Resultados** - Tabela detalhada com filtros e downloads
5. **Histórico** - Lista de consultas anteriores

## 🔌 Endpoints da API

```
POST   /consultas                    → Iniciar nova consulta
GET    /consultas/{id}/status        → Obter status (polling)
GET    /consultas/{id}/resultado     → Buscar multas
GET    /consultas/{id}/excel         → Download Excel
GET    /consultas/{id}/pdf/{file}    → Download PDF
GET    /consultas/historico          → Listar histórico
GET    /health                       → Health check
```

## 🛠️ Tecnologias

### Backend
- **Python 3.8+** - Linguagem principal
- **Playwright** - Automação web (headless browser)
- **FastAPI** - Framework web moderno
- **Pandas** - Manipulação de dados
- **openpyxl** - Geração de Excel
- **pdfplumber** - Extração de texto de PDFs

### Frontend
- **Next.js 14** - Framework React
- **TypeScript** - Tipagem estática
- **Material-UI** - Componentes de interface
- **React Query** - Gerenciamento de estado
- **Axios** - Cliente HTTP
- **XLSX** - Leitura de planilhas

## 📊 Fluxo de Uso

```
1. Usuário adiciona veículos (manual ou Excel)
   ↓
2. Clica em "Iniciar Consulta"
   ↓
3. Backend inicia automação Playwright
   ↓
4. Para cada veículo:
   - Acessa DETRAN-CE
   - Extrai multas
   - Baixa PDFs
   - Extrai dados dos PDFs
   ↓
5. Gera Excel consolidado
   ↓
6. Frontend exibe resultados
   ↓
7. Usuário baixa Excel e/ou PDFs
```

## 🎯 Casos de Uso

- 📋 **Despachantes** - Consulta rápida para clientes
- 🚚 **Empresas de Transporte** - Monitoramento de frotas
- 🏢 **Frotas Corporativas** - Gestão de multas
- 💼 **Contabilidade Veicular** - Relatórios financeiros
- 👤 **Uso Pessoal** - Consulta de veículos próprios

## ⚙️ Configuração

### Variáveis de Ambiente

**Frontend (`.env.local`):**
```env
NEXT_PUBLIC_API_URL=http://localhost:8000
```

**Backend (em `detran_manual.py`):**
```python
URL = "https://sistemas.detran.ce.gov.br/central"
EXCEL_ARQUIVO = "resultado_detran_organizado.xlsx"
INTERVALO_ENTRE_CONSULTAS = 2  # segundos
```

## 🔐 Segurança

- ✅ Validação de entrada de dados
- ✅ CORS configurado
- ✅ Sanitização de uploads
- ✅ Tratamento de erros
- ⚠️ **Produção:** Adicionar autenticação e HTTPS

## 🚨 Solução de Problemas

### Erro de Conexão Recusada

**Problema:** `ERR_CONNECTION_REFUSED`

**Solução:**
1. Verifique se o site do DETRAN está acessível
2. Confirme a URL em `detran_manual.py`
3. Teste manualmente no navegador

### Frontend não se conecta ao Backend

**Problema:** Erro de CORS ou timeout

**Solução:**
1. Confirme que a API está rodando em `:8000`
2. Verifique CORS em `api_server.py`
3. Valide URL em `frontend/.env.local`

### Excel não é gerado

**Problema:** Erro ao salvar Excel

**Solução:**
1. Feche o arquivo Excel se estiver aberto
2. Verifique permissões da pasta
3. Confirme que `openpyxl` está instalado

## 📈 Melhorias Futuras

- [ ] Autenticação de usuários
- [ ] Banco de dados persistente (PostgreSQL/MongoDB)
- [ ] Agendamento automático de consultas
- [ ] Notificações por email
- [ ] Relatórios analíticos
- [ ] Exportação para outros formatos (PDF, CSV)
- [ ] Integração com outros DETRANs

## 📞 Suporte

Para problemas ou dúvidas:
1. Verifique a [documentação completa](./frontend/)
2. Consulte o [guia de comandos](./frontend/COMANDOS_RAPIDOS.md)
3. Revise a [integração backend](./frontend/INTEGRACAO_BACKEND.md)

## 📄 Licença

Este projeto é fornecido "como está" para fins educacionais e de automação pessoal.

---

**Desenvolvido com ❤️ usando Python, FastAPI, Next.js e Material-UI**

🚀 **Sistema 100% funcional e pronto para uso!**
