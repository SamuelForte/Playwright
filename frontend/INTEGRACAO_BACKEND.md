# 🔗 Integração Backend (FastAPI + Playwright)

Este guia mostra como integrar o frontend Next.js com o código Python existente (`detran_manual.py`).

## 📋 Arquitetura

```
Frontend (Next.js)  →  API (FastAPI)  →  Automação (Playwright)
    :3000                 :8000              detran_manual.py
```

## 🛠️ Criar API FastAPI

Crie um arquivo `api_server.py` na raiz do projeto:

```python
from fastapi import FastAPI, BackgroundTasks, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse
from pydantic import BaseModel
from typing import List, Dict, Optional
import uuid
import asyncio
from datetime import datetime
import os

# Importar funções do detran_manual.py
from detran_manual import processar_veiculo, salvar_no_excel
from playwright.sync_api import sync_playwright

app = FastAPI(title="DETRAN-CE API")

# CORS para permitir frontend
app.add_middleware(
    CORSMiddleware,
    allow_origins=["http://localhost:3000"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# Armazenamento em memória (usar banco de dados em produção)
consultas_db: Dict[str, Dict] = {}

# ================= MODELS =================

class Veiculo(BaseModel):
    placa: str
    renavam: str

class IniciarConsultaRequest(BaseModel):
    veiculos: List[Veiculo]

class VeiculoStatus(BaseModel):
    placa: str
    status: str
    multas_count: int
    valor_total: float
    mensagem: Optional[str] = None

class ConsultaStatus(BaseModel):
    id: str
    status: str
    veiculos: List[VeiculoStatus]
    total_multas: int
    valor_total: float
    created_at: str

# ================= FUNÇÕES AUXILIARES =================

def processar_consulta_background(consulta_id: str, veiculos: List[Veiculo]):
    """Processa veículos em background usando Playwright"""
    
    consulta = consultas_db[consulta_id]
    consulta["status"] = "processing"
    
    todas_multas = []
    total_geral = 0.0
    
    try:
        with sync_playwright() as p:
            browser = p.chromium.launch(headless=False)
            
            for i, veiculo_data in enumerate(veiculos, 1):
                # Atualiza status do veículo
                veiculo_status = next(
                    v for v in consulta["veiculos"] 
                    if v["placa"] == veiculo_data.placa
                )
                veiculo_status["status"] = "processing"
                veiculo_status["mensagem"] = "Consultando DETRAN..."
                
                try:
                    # Chama a função do detran_manual.py
                    veiculo_dict = {
                        "placa": veiculo_data.placa,
                        "renavam": veiculo_data.renavam
                    }
                    
                    total, multas = processar_veiculo(browser, veiculo_dict, i)
                    
                    # Atualiza status
                    veiculo_status["status"] = "completed"
                    veiculo_status["multas_count"] = len(multas)
                    veiculo_status["valor_total"] = total
                    veiculo_status["mensagem"] = f"{len(multas)} multa(s) encontrada(s)"
                    
                    todas_multas.extend(multas)
                    total_geral += total
                    
                except Exception as e:
                    veiculo_status["status"] = "error"
                    veiculo_status["mensagem"] = str(e)
            
            browser.close()
        
        # Salvar Excel
        if todas_multas:
            excel_path = f"resultado_detran_{consulta_id}.xlsx"
            consulta["excel_path"] = excel_path
            
            # Usar a função salvar_no_excel do detran_manual.py
            # mas com nome customizado
            import pandas as pd
            import openpyxl
            from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
            
            df = pd.DataFrame(todas_multas)
            df.to_excel(excel_path, index=False, sheet_name="Resultado DETRAN", engine='openpyxl')
            
            # Aplicar formatação
            wb = openpyxl.load_workbook(excel_path)
            ws = wb.active
            
            header_fill = PatternFill("solid", fgColor="1F4E78")
            header_font = Font(bold=True, color="FFFFFF")
            center = Alignment(horizontal="center", vertical="center", wrap_text=True)
            
            for cell in ws[1]:
                cell.fill = header_fill
                cell.font = header_font
                cell.alignment = center
            
            wb.save(excel_path)
        
        # Marca consulta como concluída
        consulta["status"] = "completed"
        consulta["multas"] = todas_multas
        consulta["total_multas"] = len(todas_multas)
        consulta["valor_total"] = total_geral
        
    except Exception as e:
        consulta["status"] = "error"
        consulta["erro"] = str(e)

# ================= ENDPOINTS =================

@app.post("/consultas")
async def iniciar_consulta(
    request: IniciarConsultaRequest, 
    background_tasks: BackgroundTasks
):
    """Inicia uma nova consulta de veículos"""
    
    consulta_id = str(uuid.uuid4())
    
    # Cria registro da consulta
    consultas_db[consulta_id] = {
        "id": consulta_id,
        "status": "pending",
        "veiculos": [
            {
                "placa": v.placa,
                "status": "pending",
                "multas_count": 0,
                "valor_total": 0.0,
                "mensagem": "Aguardando processamento"
            }
            for v in request.veiculos
        ],
        "total_multas": 0,
        "valor_total": 0.0,
        "created_at": datetime.now().isoformat(),
        "multas": [],
        "excel_path": None
    }
    
    # Processa em background
    background_tasks.add_task(
        processar_consulta_background, 
        consulta_id, 
        request.veiculos
    )
    
    return {"consulta_id": consulta_id}

@app.get("/consultas/{consulta_id}/status")
async def obter_status(consulta_id: str):
    """Retorna o status atual da consulta"""
    
    if consulta_id not in consultas_db:
        raise HTTPException(status_code=404, detail="Consulta não encontrada")
    
    consulta = consultas_db[consulta_id]
    
    return ConsultaStatus(
        id=consulta["id"],
        status=consulta["status"],
        veiculos=[VeiculoStatus(**v) for v in consulta["veiculos"]],
        total_multas=consulta["total_multas"],
        valor_total=consulta["valor_total"],
        created_at=consulta["created_at"]
    )

@app.get("/consultas/{consulta_id}/resultado")
async def obter_resultado(consulta_id: str):
    """Retorna o resultado completo da consulta"""
    
    if consulta_id not in consultas_db:
        raise HTTPException(status_code=404, detail="Consulta não encontrada")
    
    consulta = consultas_db[consulta_id]
    
    if consulta["status"] != "completed":
        raise HTTPException(
            status_code=400, 
            detail="Consulta ainda não foi concluída"
        )
    
    return {
        "id": consulta_id,
        "multas": consulta["multas"],
        "total_multas": consulta["total_multas"],
        "valor_total": consulta["valor_total"],
        "excel_path": consulta["excel_path"]
    }

@app.get("/consultas/{consulta_id}/excel")
async def baixar_excel(consulta_id: str):
    """Baixa o Excel da consulta"""
    
    if consulta_id not in consultas_db:
        raise HTTPException(status_code=404, detail="Consulta não encontrada")
    
    consulta = consultas_db[consulta_id]
    excel_path = consulta.get("excel_path")
    
    if not excel_path or not os.path.exists(excel_path):
        raise HTTPException(status_code=404, detail="Excel não encontrado")
    
    return FileResponse(
        excel_path,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        filename=f"resultado_detran_{consulta_id}.xlsx"
    )

@app.get("/consultas/{consulta_id}/pdf/{filename}")
async def baixar_pdf(consulta_id: str, filename: str):
    """Baixa um PDF específico"""
    
    # Buscar PDF na pasta boletos
    from datetime import datetime
    data_hoje = datetime.now().strftime("%d-%m-%Y")
    pdf_path = os.path.join("boletos", data_hoje, filename)
    
    if not os.path.exists(pdf_path):
        raise HTTPException(status_code=404, detail="PDF não encontrado")
    
    return FileResponse(
        pdf_path,
        media_type="application/pdf",
        filename=filename
    )

@app.get("/consultas/historico")
async def listar_historico():
    """Lista todas as consultas"""
    
    historico = []
    
    for consulta in consultas_db.values():
        historico.append({
            "id": consulta["id"],
            "status": consulta["status"],
            "veiculos": consulta["veiculos"],
            "total_multas": consulta["total_multas"],
            "valor_total": consulta["valor_total"],
            "created_at": consulta["created_at"]
        })
    
    # Ordenar por data (mais recente primeiro)
    historico.sort(key=lambda x: x["created_at"], reverse=True)
    
    return historico

@app.get("/health")
async def health_check():
    """Health check da API"""
    return {"status": "ok", "timestamp": datetime.now().isoformat()}

# ================= RODAR SERVIDOR =================

if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=8000)
```

## 🚀 Como Rodar

### 1. Instalar FastAPI

```powershell
pip install fastapi uvicorn
```

### 2. Rodar o Servidor

```powershell
python api_server.py
```

Ou com uvicorn diretamente:

```powershell
uvicorn api_server:app --reload --port 8000
```

### 3. Testar API

Acesse: http://localhost:8000/docs

Você verá a documentação interativa do Swagger.

## 🔄 Fluxo Completo

```
1. Frontend envia POST /consultas
   └─> API cria registro e inicia background task

2. Background task roda processar_veiculo()
   ├─> Abre Playwright
   ├─> Consulta cada veículo
   ├─> Extrai multas
   └─> Salva Excel

3. Frontend faz polling GET /status
   └─> API retorna status atualizado

4. Quando completo, frontend busca GET /resultado
   └─> API retorna todas as multas

5. Frontend baixa GET /excel
   └─> API retorna arquivo .xlsx
```

## 🗄️ Persistência (Opcional)

Para produção, substitua `consultas_db` por um banco de dados:

```python
# Exemplo com SQLite
import sqlite3

# Ou MongoDB
from pymongo import MongoClient

# Ou PostgreSQL
import psycopg2
```

## 🔒 Melhorias Recomendadas

1. **Autenticação**
```python
from fastapi.security import OAuth2PasswordBearer
oauth2_scheme = OAuth2PasswordBearer(tokenUrl="token")
```

2. **Rate Limiting**
```python
from slowapi import Limiter
limiter = Limiter(key_func=get_remote_address)
```

3. **Logs Estruturados**
```python
import logging
logging.basicConfig(level=logging.INFO)
```

4. **Validação Avançada**
```python
from pydantic import validator

class Veiculo(BaseModel):
    placa: str
    renavam: str
    
    @validator('placa')
    def validate_placa(cls, v):
        if not re.match(r'^[A-Z]{3}\d{4}$|^[A-Z]{3}\d[A-Z]\d{2}$', v):
            raise ValueError('Placa inválida')
        return v
```

## ✅ Checklist de Integração

- [ ] `api_server.py` criado
- [ ] FastAPI instalado
- [ ] CORS configurado
- [ ] Servidor rodando em :8000
- [ ] Frontend rodando em :3000
- [ ] Testar endpoint /health
- [ ] Testar consulta de 1 veículo
- [ ] Verificar salvamento de Excel
- [ ] Testar download de PDF

---

**Pronto! Agora você tem uma API completa integrada com o Playwright! 🎉**
