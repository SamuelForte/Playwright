# 🔐 Sistema de Autenticação OAuth 2.0 - DETRAN-CE

Sistema completo de autenticação usando Google OAuth 2.0 com JWT próprio.

## 🚀 Quick Start

```bash
# 1. Instalar dependências e configurar banco
npm run setup

# 2. Configurar variáveis de ambiente
cp .env.example .env
# Edite .env com suas credenciais do Google

# 3. Rodar em desenvolvimento
npm run dev

# 4. Servidor estará em:
# http://localhost:4000
```

## 📋 Pré-requisitos

- Node.js 18+
- npm ou yarn
- Credenciais do Google Cloud Console

## 🔑 Obter Credenciais do Google

### Passo 1: Criar Projeto no Google Cloud Console

1. Acesse: https://console.cloud.google.com/
2. Clique em **"Criar Projeto"** ou selecione um existente
3. Dê um nome ao projeto (ex: `detran-ce-auth`)
4. Clique em **"Criar"**

### Passo 2: Habilitar Google OAuth 2.0 API

1. No menu lateral, vá em **"APIs e Serviços"** → **"Biblioteca"**
2. Busque por **"Google+ API"** e clique em **"Ativar"**
3. (Opcional) Ative também **"People API"** para mais informações do usuário

### Passo 3: Configurar Tela de Consentimento

1. No menu lateral, vá em **"APIs e Serviços"** → **"Tela de consentimento OAuth"**
2. Escolha **"Externo"** (ou "Interno" se for workspace)
3. Preencha:
   - **Nome do app**: DETRAN-CE Automation
   - **E-mail de suporte**: seu-email@gmail.com
   - **Domínios autorizados**: localhost (para desenvolvimento)
   - **E-mail de contato do desenvolvedor**: seu-email@gmail.com
4. Clique em **"Salvar e continuar"**
5. Em **"Escopos"**, clique em **"Adicionar ou remover escopos"**
6. Selecione:
   - `.../auth/userinfo.email`
   - `.../auth/userinfo.profile`
7. Clique em **"Salvar e continuar"**
8. Em **"Usuários de teste"**, adicione seu e-mail do Google
9. Clique em **"Salvar e continuar"**

### Passo 4: Criar Credenciais OAuth 2.0

1. No menu lateral, vá em **"APIs e Serviços"** → **"Credenciais"**
2. Clique em **"+ Criar credenciais"** → **"ID do cliente OAuth"**
3. Escolha **"Aplicativo da Web"**
4. Preencha:
   - **Nome**: DETRAN-CE Auth Server
   - **URIs de redirecionamento autorizados**:
     ```
     http://localhost:4000/auth/google/callback
     ```
   - **Origens JavaScript autorizadas** (opcional):
     ```
     http://localhost:3000
     http://localhost:4000
     ```
5. Clique em **"Criar"**
6. **COPIE** o `Client ID` e o `Client Secret` que aparecerem
7. Cole no arquivo `.env`:
   ```env
   GOOGLE_CLIENT_ID=123456789-abcdef.apps.googleusercontent.com
   GOOGLE_CLIENT_SECRET=GOCSPX-abc123def456
   ```

### Passo 5: Gerar JWT Secret

```bash
# No terminal (Linux/Mac):
openssl rand -base64 32

# Ou use um gerador online:
# https://generate-secret.vercel.app/32

# Cole o resultado no .env:
JWT_SECRET=sua-chave-gerada-aqui
```

## 📁 Estrutura de Arquivos

```
auth/
├── server.ts              # Servidor Express com todas as rotas
├── prisma/
│   ├── schema.prisma      # Schema do banco de dados
│   └── auth.db            # Banco SQLite (gerado automaticamente)
├── package.json           # Dependências
├── tsconfig.json          # Configuração TypeScript
├── .env.example           # Template de variáveis
├── .env                   # Suas credenciais (não commitar!)
└── README.md              # Este arquivo
```

## 🔌 Endpoints da API

### Autenticação

#### `GET /auth/google`
Redireciona para tela de login do Google.

**Uso no frontend:**
```typescript
window.location.href = 'http://localhost:4000/auth/google';
```

#### `GET /auth/google/callback`
Callback do Google (automático). Redireciona para:
```
http://localhost:3000/auth/callback?token=JWT_TOKEN
```

**Frontend deve capturar o token:**
```typescript
// Em /auth/callback
const params = new URLSearchParams(window.location.search);
const token = params.get('token');
localStorage.setItem('accessToken', token);
```

#### `POST /auth/logout`
Invalida o refreshToken do usuário.

**Headers:**
```
Authorization: Bearer SEU_JWT_TOKEN
```

**Response:**
```json
{ "message": "Logout realizado com sucesso" }
```

### Rotas Protegidas

#### `GET /auth/me`
Retorna informações do usuário autenticado.

**Headers:**
```
Authorization: Bearer SEU_JWT_TOKEN
```

**Response:**
```json
{
  "user": {
    "id": "uuid-do-usuario",
    "email": "usuario@gmail.com",
    "name": "Nome do Usuário",
    "picture": "https://lh3.googleusercontent.com/...",
    "createdAt": "2026-01-20T10:00:00.000Z"
  }
}
```

#### `GET /auth/health`
Health check do servidor.

**Response:**
```json
{
  "status": "ok",
  "timestamp": "2026-01-20T10:00:00.000Z"
}
```

## 🛡️ Usando o Middleware de Autenticação

```typescript
import { authMiddleware } from './server';

// Em suas rotas protegidas:
app.get('/api/consultas', authMiddleware, async (req, res) => {
  // req.user contém { userId, email }
  const userId = req.user!.userId;
  // Sua lógica aqui...
});
```

## 🔄 Integração com Frontend (Next.js)

### 1. Criar página de login

```typescript
// frontend/src/app/login/page.tsx
export default function LoginPage() {
  const handleGoogleLogin = () => {
    window.location.href = 'http://localhost:4000/auth/google';
  };

  return (
    <button onClick={handleGoogleLogin}>
      Entrar com Google
    </button>
  );
}
```

### 2. Criar página de callback

```typescript
// frontend/src/app/auth/callback/page.tsx
'use client';

import { useEffect } from 'react';
import { useRouter } from 'next/navigation';

export default function AuthCallback() {
  const router = useRouter();

  useEffect(() => {
    const params = new URLSearchParams(window.location.search);
    const token = params.get('token');
    const error = params.get('error');

    if (error) {
      router.push('/login?error=' + error);
      return;
    }

    if (token) {
      localStorage.setItem('accessToken', token);
      router.push('/dashboard');
    }
  }, [router]);

  return <div>Autenticando...</div>;
}
```

### 3. Criar cliente API com autenticação

```typescript
// frontend/src/lib/api.ts
import axios from 'axios';

const api = axios.create({
  baseURL: 'http://localhost:8000',
});

// Interceptor para adicionar token em todas as requisições
api.interceptors.request.use((config) => {
  const token = localStorage.getItem('accessToken');
  if (token) {
    config.headers.Authorization = `Bearer ${token}`;
  }
  return config;
});

// Interceptor para tratar erro 401 (não autenticado)
api.interceptors.response.use(
  (response) => response,
  (error) => {
    if (error.response?.status === 401) {
      localStorage.removeItem('accessToken');
      window.location.href = '/login';
    }
    return Promise.reject(error);
  }
);

export default api;
```

## 🗄️ Comandos do Prisma

```bash
# Gerar cliente Prisma
npm run prisma:generate

# Criar migração
npm run prisma:migrate

# Abrir Prisma Studio (GUI do banco)
npm run prisma:studio

# Reset do banco (CUIDADO: apaga tudo)
npx prisma migrate reset
```

## 📊 Banco de Dados (SQLite)

**Tabela Users:**

| Campo        | Tipo     | Descrição                        |
|--------------|----------|----------------------------------|
| id           | String   | UUID único                       |
| email        | String   | Email do Google (único)          |
| name         | String   | Nome completo                    |
| picture      | String?  | URL da foto de perfil            |
| googleId     | String   | ID do usuário no Google (único)  |
| refreshToken | String?  | Token para renovação             |
| createdAt    | DateTime | Data de criação                  |
| updatedAt    | DateTime | Data de atualização              |

## 🔒 Segurança

- ✅ JWT com expiração de 7 dias
- ✅ RefreshToken armazenado no banco
- ✅ CORS configurado para CLIENT_URL apenas
- ✅ Validação de e-mail verificado
- ✅ Logs estruturados (sem console.log)
- ✅ Tratamento de erros completo
- ✅ TypeScript em 100% do código
- ✅ Graceful shutdown

## 🚨 Troubleshooting

### Erro: "redirect_uri_mismatch"
- Verifique se a URI no Google Console é exatamente: `http://localhost:4000/auth/google/callback`
- Aguarde 5 minutos após salvar as configurações no Google

### Erro: "Access blocked: This app's request is invalid"
- Configure a Tela de Consentimento OAuth
- Adicione seu e-mail em "Usuários de teste"

### Erro: "Email not verified"
- Use um e-mail verificado no Google
- Confirme seu e-mail no Google antes de usar

### Erro: "Missing required environment variables"
- Verifique se todas as variáveis estão no `.env`
- Reinicie o servidor após alterar `.env`

## 📝 Scripts Disponíveis

```bash
npm run dev          # Desenvolvimento com hot-reload
npm run build        # Build para produção
npm start            # Rodar versão de produção
npm run setup        # Setup inicial completo
npm run prisma:studio # Abrir GUI do banco
```

## 🔗 Links Úteis

- [Google Cloud Console](https://console.cloud.google.com/)
- [Documentação OAuth 2.0 do Google](https://developers.google.com/identity/protocols/oauth2)
- [Prisma Docs](https://www.prisma.io/docs)
- [JWT.io](https://jwt.io/) - Decodificar tokens

## 📄 Licença

MIT

---

**Desenvolvido para o sistema DETRAN-CE Automation**
