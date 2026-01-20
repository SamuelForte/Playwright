# ⚡ Comandos Rápidos

## 🚀 Setup Inicial

```powershell
# Navegar para pasta frontend
cd frontend

# Instalar dependências
npm install

# Voltar para raiz
cd ..
```

## 💻 Desenvolvimento

### Rodar Frontend

```powershell
cd frontend
npm run dev
```

Acesse: http://localhost:3000

### Rodar Backend (API)

```powershell
# Opção 1: Com uvicorn direto
uvicorn api_server:app --reload --port 8000

# Opção 2: Executar script
python api_server.py
```

Acesse: http://localhost:8000/docs

### Rodar Ambos Simultaneamente

**Terminal 1 (Backend):**
```powershell
python api_server.py
```

**Terminal 2 (Frontend):**
```powershell
cd frontend
npm run dev
```

## 🔧 Build e Deploy

### Build Frontend para Produção

```powershell
cd frontend

# Criar build otimizado
npm run build

# Rodar build
npm start
```

### Build Backend para Produção

```powershell
# Instalar gunicorn (produção)
pip install gunicorn

# Rodar com gunicorn
gunicorn api_server:app -w 4 -k uvicorn.workers.UvicornWorker --bind 0.0.0.0:8000
```

## 📦 Instalar Dependências Faltantes

### Frontend

```powershell
cd frontend

# Instalar todas
npm install

# Instalar individualmente se necessário
npm install @mui/material @emotion/react @emotion/styled
npm install @mui/icons-material
npm install @tanstack/react-query
npm install axios
npm install xlsx
npm install react-dropzone
npm install date-fns
```

### Backend

```powershell
pip install fastapi
pip install uvicorn[standard]
pip install playwright
pip install pandas
pip install openpyxl
pip install pdfplumber
```

## 🧹 Limpeza

```powershell
# Limpar cache do Next.js
cd frontend
rm -rf .next

# Limpar node_modules (se precisar reinstalar)
rm -rf node_modules
npm install

# Limpar cache Python
rm -rf __pycache__
rm -rf venv
```

## 🐛 Debug

### Ver Logs do Frontend

```powershell
cd frontend
npm run dev
# Logs aparecerão no terminal
```

### Ver Logs do Backend

```powershell
# Rodar com logs detalhados
uvicorn api_server:app --reload --log-level debug
```

### Testar API com curl

```powershell
# Health check
curl http://localhost:8000/health

# Iniciar consulta
curl -X POST http://localhost:8000/consultas `
  -H "Content-Type: application/json" `
  -d '{"veiculos":[{"placa":"ABC1234","renavam":"01365705622"}]}'

# Ver status
curl http://localhost:8000/consultas/{id}/status
```

## 📊 Monitoramento

### Ver Processos Rodando

```powershell
# Ver processos Node
Get-Process node

# Ver processos Python
Get-Process python

# Matar processo na porta 3000
netstat -ano | findstr :3000
taskkill /PID <PID> /F

# Matar processo na porta 8000
netstat -ano | findstr :8000
taskkill /PID <PID> /F
```

## 🔄 Atualizar Dependências

```powershell
# Frontend
cd frontend
npm update

# Backend
pip install --upgrade fastapi uvicorn playwright
```

## 📝 Úteis

### Gerar requirements.txt

```powershell
pip freeze > requirements.txt
```

### Instalar do requirements.txt

```powershell
pip install -r requirements.txt
```

### Ver Versão Instalada

```powershell
# Node e npm
node --version
npm --version

# Python e pip
python --version
pip --version

# Next.js
cd frontend
npm list next
```

## 🎯 Workflows Comuns

### Adicionar Novo Componente

```powershell
# Criar arquivo
cd frontend/src/components
New-Item MeuComponente.tsx

# Importar no código
# import MeuComponente from '@/components/MeuComponente'
```

### Adicionar Nova Página

```powershell
# Criar pasta e arquivo
cd frontend/src/app
mkdir minha-pagina
New-Item minha-pagina/page.tsx

# Acessar em: http://localhost:3000/minha-pagina
```

### Adicionar Novo Endpoint

Edite `api_server.py`:

```python
@app.get("/novo-endpoint")
async def novo_endpoint():
    return {"mensagem": "Olá!"}
```

### Hot Reload

Ambos frontend e backend têm **hot reload automático**:

- ✅ Frontend: Salve arquivo → atualiza automaticamente
- ✅ Backend: Salve arquivo → reinicia automaticamente (com --reload)

## 🚨 Solução de Problemas

### Erro: EADDRINUSE (porta em uso)

```powershell
# Ver quem está usando a porta
netstat -ano | findstr :3000

# Matar processo
taskkill /PID <PID> /F
```

### Erro: Module not found

```powershell
cd frontend
rm -rf node_modules package-lock.json
npm install
```

### Erro: Python import

```powershell
# Ativar venv
.\venv\Scripts\Activate.ps1

# Ou instalar globalmente
pip install <pacote>
```

---

**💡 Dica:** Mantenha este arquivo aberto enquanto desenvolve!
