# 🚀 Guia de Deploy - Vercel + Render + Supabase

## Pré-requisitos
- Conta no [Vercel](https://vercel.com) (grátis)
- Conta no [Render](https://render.com) (grátis)
- Projeto Supabase já configurado ✅

---

## 1️⃣ Deploy do Frontend (Vercel)

### Passo 1: Instalar Vercel CLI
```bash
npm install -g vercel
```

### Passo 2: Fazer deploy do frontend
```bash
cd frontend
vercel
```

Siga as instruções:
- Fazer login na Vercel
- Confirmar o projeto
- Aceitar as configurações padrão

### Passo 3: Configurar variáveis de ambiente
No painel da Vercel (dashboard):
1. Vá em **Settings** → **Environment Variables**
2. Adicione:
   - `NEXT_PUBLIC_SUPABASE_URL` = sua URL do Supabase
   - `NEXT_PUBLIC_SUPABASE_ANON_KEY` = sua chave anônima do Supabase
   - `NEXT_PUBLIC_API_URL` = (será adicionada depois do deploy do backend)

### Passo 4: Redesploy
```bash
vercel --prod
```

---

## 2️⃣ Deploy do Backend Python (Render)

### Passo 1: Criar conta no Render
1. Acesse [render.com](https://render.com)
2. Faça login com GitHub

### Passo 2: Criar novo Web Service
1. Clique em **New** → **Web Service**
2. Conecte seu repositório GitHub (ou faça upload manual)
3. Configure:
   - **Name**: `detran-api`
   - **Runtime**: `Python 3`
   - **Build Command**: 
     ```
     pip install -r requirements.txt && playwright install chromium && playwright install-deps
     ```
   - **Start Command**: 
     ```
     python app.py
     ```
   - **Plan**: `Free`

### Passo 3: Configurar variáveis de ambiente
Em **Environment**:
- `SUPABASE_URL` = sua URL do Supabase
- `SUPABASE_SERVICE_KEY` = sua chave de serviço do Supabase (não a anon key!)
- `PYTHON_VERSION` = `3.11.0`

### Passo 4: Deploy
Clique em **Create Web Service** e aguarde o deploy (5-10 min)

### Passo 5: Obter URL do backend
Copie a URL gerada (ex: `https://detran-api.onrender.com`)

---

## 3️⃣ Conectar Frontend ao Backend

### Atualizar URL da API na Vercel
1. Vá no painel da Vercel
2. **Settings** → **Environment Variables**
3. Adicione/atualize:
   - `NEXT_PUBLIC_API_URL` = `https://seu-app.onrender.com`
4. Vá em **Deployments** → **Redeploy**

---

## 4️⃣ Configurar CORS no Backend

O backend precisa permitir requisições do domínio da Vercel.

### Atualizar api_server.py:
Adicione o domínio da Vercel em `allow_origins`:
```python
app.add_middleware(
    CORSMiddleware,
    allow_origins=[
        "http://localhost:3000",
        "https://seu-app.vercel.app",  # ← Adicionar
        "https://seu-dominio-customizado.com"  # Se tiver
    ],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)
```

Depois faça push para GitHub (Render vai redesploy automaticamente).

---

## 5️⃣ Verificação Final

### Teste o backend:
```bash
curl https://seu-app.onrender.com/health
```

### Teste o frontend:
Acesse `https://seu-app.vercel.app`

---

## 📊 Limites Gratuitos

### Vercel (Frontend):
- ✅ Ilimitado
- ✅ SSL automático
- ✅ CDN global
- ✅ Sem hibernação

### Render (Backend):
- ⚠️ 750 horas/mês (suficiente para 1 serviço)
- ⚠️ Hiberna após 15 min de inatividade
- ⚠️ Primeiro request pode demorar 30s (cold start)
- ✅ SSL automático

### Supabase:
- ✅ PostgreSQL 500MB
- ✅ 50.000 usuários autenticados
- ✅ 1GB de storage
- ✅ 2GB de transferência/mês

---

## 🔄 Deploy Automático (CI/CD)

### Vercel:
- Push para `main` → deploy automático ✅

### Render:
1. Conecte o repositório GitHub
2. Ative **Auto-Deploy** nas configurações
3. Push para `main` → deploy automático ✅

---

## 🐛 Troubleshooting

### Erro de CORS
- Adicione o domínio da Vercel no `allow_origins` do backend
- Redesploy o backend

### Backend demora para responder
- Normal no plano gratuito (cold start)
- Considere manter o serviço ativo com ping periódico

### Erro de dependências Python
- Verifique se `requirements.txt` está na raiz
- Certifique-se que Playwright está instalado corretamente

### Frontend não conecta ao backend
- Verifique se `NEXT_PUBLIC_API_URL` está configurada corretamente
- Use HTTPS (não HTTP)

---

## 💰 Custos

Total: **R$ 0,00/mês** 🎉

Limitações:
- Backend hiberna após inatividade
- 750h/mês (31 dias = 744h, então praticamente ilimitado para 1 app)

Se precisar de mais, considere:
- **Railway**: $5/mês de crédito grátis
- **Render Paid**: $7/mês (sem hibernação)
