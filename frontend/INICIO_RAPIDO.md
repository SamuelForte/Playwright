# 🚀 Guia Rápido de Início

## Passo 1: Instalar Dependências

```powershell
cd frontend
npm install
```

Isso irá instalar:
- Next.js 14
- Material-UI (componentes visuais)
- React Query (gerenciamento de estado)
- Axios (requisições HTTP)
- XLSX (leitura de planilhas)
- E todas as outras dependências

## Passo 2: Rodar o Frontend

```powershell
npm run dev
```

O frontend estará disponível em: **http://localhost:3000**

## Passo 3: Configurar o Backend (API)

Você precisará criar a API FastAPI que o frontend espera. Os endpoints necessários:

```python
# api.py
from fastapi import FastAPI, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel
from typing import List
import uuid

app = FastAPI()

# Configurar CORS para permitir requisições do frontend
app.add_middleware(
    CORSMiddleware,
    allow_origins=["http://localhost:3000"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

class Veiculo(BaseModel):
    placa: str
    renavam: str

class IniciarConsultaRequest(BaseModel):
    veiculos: List[Veiculo]

@app.post("/consultas")
async def iniciar_consulta(request: IniciarConsultaRequest):
    consulta_id = str(uuid.uuid4())
    
    # TODO: Chamar sua função processar_veiculo() aqui
    # em background para cada veículo
    
    return {"consulta_id": consulta_id}

@app.get("/consultas/{consulta_id}/status")
async def obter_status(consulta_id: str):
    # TODO: Retornar status da consulta
    return {
        "id": consulta_id,
        "status": "processing",
        "veiculos": [],
        "total_multas": 0,
        "valor_total": 0,
        "created_at": "2026-01-19T10:00:00Z"
    }

@app.get("/consultas/{consulta_id}/resultado")
async def obter_resultado(consulta_id: str):
    # TODO: Retornar resultado completo
    return {
        "id": consulta_id,
        "multas": [],
        "total_multas": 0,
        "valor_total": 0
    }
```

## Passo 4: Rodar o Backend

```powershell
cd ..  # Voltar para pasta raiz
python -m uvicorn api:app --reload --port 8000
```

## 📱 Como Usar a Interface

1. **Acesse http://localhost:3000**
2. **Dashboard** - Veja o resumo
3. **Nova Consulta** - Adicione veículos:
   - **Opção A:** Digite manualmente placa e RENAVAM
   - **Opção B:** Arraste um arquivo Excel com colunas `placa` e `renavam`
4. **Clique em "Iniciar Consulta Automática"**
5. **Acompanhe o processamento** em tempo real
6. **Veja os resultados** com tabela completa
7. **Baixe Excel** ou **PDFs individuais**

## 📊 Formato do Excel para Importação

Crie um arquivo Excel (.xlsx) com as colunas:

| placa   | renavam      |
|---------|--------------|
| SBA7F09 | 01365705622  |
| TIF1J98 | 01450499292  |

## 🔧 Configuração (Opcional)

Edite `.env.local` se sua API estiver em outra URL:

```env
NEXT_PUBLIC_API_URL=http://seu-servidor:8000
```

## ✅ Checklist de Integração

- [ ] Frontend rodando em http://localhost:3000
- [ ] Backend rodando em http://localhost:8000
- [ ] CORS configurado no backend
- [ ] Endpoints da API implementados
- [ ] Testar consulta de 1 veículo
- [ ] Testar importação de Excel
- [ ] Verificar download de resultados

## 🐛 Problemas Comuns

**Frontend não carrega:**
- Verifique se rodou `npm install`
- Certifique-se que a porta 3000 está livre

**Erro de CORS:**
- Adicione o middleware CORS no backend
- Verifique se a URL do frontend está permitida

**API não responde:**
- Confirme que o backend está rodando
- Verifique a URL em `.env.local`

**Upload de Excel não funciona:**
- Certifique-se que o arquivo tem as colunas `placa` e `renavam`
- Formato aceito: .xlsx, .xls, .csv

## 📞 Próximos Passos

1. Adaptar os endpoints da API para chamar suas funções do `detran_manual.py`
2. Implementar salvamento de consultas em banco de dados (opcional)
3. Adicionar autenticação se necessário
4. Deploy em servidor de produção

---

**Pronto para usar! 🎉**
