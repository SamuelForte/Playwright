# DETRAN-CE Frontend

Front-end completo para o sistema de consulta de multas do DETRAN-CE.

## 🚀 Tecnologias

- **Next.js 14** - Framework React com App Router
- **TypeScript** - Tipagem estática
- **Material-UI (MUI)** - Componentes de UI
- **React Query** - Gerenciamento de estado e cache
- **Axios** - Cliente HTTP (para API FastAPI)
- **Supabase JS** - Persistência de condutores/indicações
- **XLSX** - Leitura e manipulação de planilhas Excel

## 📋 Funcionalidades

### ✅ Telas Implementadas

1. **Dashboard** - Resumo com cards de estatísticas (veículos consultados, multas, valor total, PDFs)
2. **Nova Consulta** - Cadastro manual de veículos ou importação via Excel/CSV
3. **Processamento** - Acompanhamento em tempo real com polling
4. **Resultados** - Tabela detalhada com filtros, busca e download de Excel/PDFs
5. **Histórico** - Lista de consultas anteriores

### 🧩 Componentes Principais

- **VehicleForm** - Formulário com validação de placa (Mercosul e antiga) e RENAVAM
- **FileUpload** - Upload de Excel/CSV com preview
- **MultasTable** - DataGrid com paginação, filtros e ordenação
- **ProcessStatus** - Indicador de status com barra de progresso
- **StatusCard** - Cards informativos para dashboard
- **Layout** - Layout com navegação e rodapé

## 🛠️ Instalação

```bash
# Navegar para a pasta frontend
cd frontend

# Instalar dependências
npm install

# Configurar variáveis de ambiente
# Edite o arquivo .env.local
# NEXT_PUBLIC_API_URL=http://localhost:8000
# NEXT_PUBLIC_SUPABASE_URL=https://<seu-projeto>.supabase.co
# NEXT_PUBLIC_SUPABASE_ANON_KEY=sb_publishable_xxx

# Rodar em desenvolvimento
npm run dev

# Build para produção
npm run build

# Rodar em produção
npm start
```

## 🔌 Integração com Backend (FastAPI) e Supabase

- **FastAPI (http://localhost:8000)**: apenas para consultas Playwright
- **Supabase**: condutores e indicações são persistidos direto pelo frontend

### Endpoints FastAPI necessários

### Endpoints Necessários

POST   /consultas                     // Iniciar nova consulta
GET    /consultas/{id}/status         // Obter status da consulta
GET    /consultas/{id}/resultado      // Obter resultado completo
GET    /consultas/{id}/excel          // Baixar Excel
GET    /consultas/{id}/pdf/{filename} // Baixar PDF individual
GET    /consultas/historico           // Listar histórico

### Estrutura de Dados Esperada

**Veiculo:**
```json
{
  "placa": "ABC1234",
  "renavam": "01365705622"
}
```

**ConsultaStatus:**
```json
{
  "id": "uuid",
  "status": "processing",
  "veiculos": [
    {
      "placa": "ABC1234",
      "status": "completed",
      "multas_count": 3,
      "valor_total": 527.43,
      "mensagem": "Concluído com sucesso"
    }
  ],
  "total_multas": 3,
  "valor_total": 527.43,
  "created_at": "2026-01-19T10:00:00Z"
}
```

**Multa:**
```json
{
  "placa": "ABC1234",
  "numero": 1,
  "ait": "V020098768",
  "ait_originaria": "-",
  "motivo": "Transitar em velocidade superior...",
  "data_infracao": "15/01/2024",
  "data_vencimento": "20/02/2024",
  "valor": "R$ 130,16",
  "valor_a_pagar": "R$ 130,16",
  "orgao_autuador": "DEMUTRAN RUSSAS",
  "codigo_pagamento": "856300000010..."
}
```

### Supabase (tabelas sugeridas)

```sql
create table condutores (
  id uuid primary key default gen_random_uuid(),
  nome text not null,
  cpf text not null unique,
  cnh_categoria text,
  cnh_vencimento date,
  pontuacao int,
  created_at timestamptz default now()
);

create table indicacoes (
  id uuid primary key default gen_random_uuid(),
  ait text not null,
  placa text not null,
  condutor_id uuid references condutores(id),
  data_indicacao timestamptz default now(),
  status text default 'registrado'
);
```

- Variáveis no `.env.local`: `NEXT_PUBLIC_SUPABASE_URL`, `NEXT_PUBLIC_SUPABASE_ANON_KEY`
- Persistência de condutores/indicações via `src/lib/api.ts` com Supabase client

## 🎨 Características de UX/UI

- ✅ Design limpo e profissional
- ✅ Cores corporativas (azul #1F4E78)
- ✅ Responsivo (funciona em desktop e mobile)
- ✅ Feedback visual em todas as ações
- ✅ Loading states e tratamento de erros
- ✅ Validação de formulários em tempo real
- ✅ Máscaras para placa e RENAVAM
- ✅ Paginação e filtros nas tabelas
- ✅ Download de arquivos (Excel e PDF)

## 📁 Estrutura do Projeto

```
frontend/
├── src/
│   ├── app/                      # Pages (App Router)
│   │   ├── dashboard/
│   │   ├── nova-consulta/
│   │   ├── processamento/[id]/
│   │   ├── resultados/[id]/
│   │   ├── historico/
│   │   ├── layout.tsx
│   │   └── page.tsx
│   ├── components/               # Componentes reutilizáveis
│   │   ├── Layout.tsx
│   │   ├── VehicleForm.tsx
│   │   ├── FileUpload.tsx
│   │   ├── MultasTable.tsx
│   │   ├── ProcessStatus.tsx
│   │   ├── StatusCard.tsx
│   │   └── Providers.tsx
│   ├── lib/                      # Utilitários
│   │   └── api.ts               # Cliente API
│   └── theme.ts                 # Tema Material-UI
├── public/                       # Arquivos estáticos
├── .env.local                    # Variáveis de ambiente
├── package.json
├── tsconfig.json
└── next.config.js
```

## 🔐 Segurança

- ✅ Validação de entrada de dados
- ✅ Prevenção de múltiplas execuções simultâneas
- ✅ Sanitização de uploads de arquivo
- ✅ Tratamento adequado de erros
- ✅ CORS configurado no backend

## 📱 Fluxo de Uso

1. **Dashboard** → Visualizar resumo geral
2. **Nova Consulta** → Adicionar veículos (manual ou importar Excel)
3. **Iniciar Consulta** → Backend inicia automação Playwright
4. **Processamento** → Acompanhar status em tempo real
5. **Resultados** → Ver multas detalhadas, baixar Excel/PDFs
6. **Histórico** → Acessar consultas anteriores

## 🎯 Próximos Passos (Opcional)

- [ ] Implementar login/autenticação
- [ ] Adicionar filtros avançados por data
- [ ] Exportar relatórios PDF personalizados
- [ ] Gráficos e dashboards analíticos
- [ ] Notificações por email
- [ ] Agendamento de consultas automáticas

## 📞 Suporte

Sistema desenvolvido para despachantes, empresas de transporte, frotas e contabilidade veicular.

---

**Desenvolvido com ❤️ usando Next.js e Material-UI**
