from fastapi import FastAPI, BackgroundTasks, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse
from pydantic import BaseModel
from typing import List, Dict, Optional
import uuid
import asyncio
from datetime import datetime
import os
import sys

# Importar funções do detran_manual.py
import detran_manual
from detran_manual import processar_veiculo, salvar_no_excel
from playwright.sync_api import sync_playwright

app = FastAPI(title="DETRAN-CE API", version="1.0.0")

# CORS para permitir frontend
app.add_middleware(
    CORSMiddleware,
    allow_origins=["http://localhost:3000"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# Armazenamento em memória (trocar por banco de dados em produção)
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
    """Processa veículos em background usando Playwright (detran_manual.py)"""
    
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
                veiculo_status["mensagem"] = "Consultando DETRAN-CE..."
                
                try:
                    # Chama a função processar_veiculo do detran_manual.py
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
                    veiculo_status["mensagem"] = f"Erro: {str(e)}"
                    print(f"❌ Erro ao processar {veiculo_data.placa}: {e}")
            
            browser.close()
        
        # Salvar Excel usando a função do detran_manual.py
        if todas_multas:
            excel_filename = f"resultado_detran_{consulta_id}.xlsx"
            
            # Temporariamente alterar o nome do arquivo
            import detran_manual
            original_excel = detran_manual.EXCEL_ARQUIVO
            detran_manual.EXCEL_ARQUIVO = excel_filename
            
            try:
                salvar_no_excel(todas_multas)
                consulta["excel_path"] = excel_filename
            finally:
                detran_manual.EXCEL_ARQUIVO = original_excel
        
        # Marca consulta como concluída
        consulta["status"] = "completed"
        consulta["multas"] = todas_multas
        consulta["total_multas"] = len(todas_multas)
        consulta["valor_total"] = total_geral
        
        print(f"✅ Consulta {consulta_id} concluída com sucesso!")
        print(f"📊 Total: {len(todas_multas)} multas | R$ {total_geral:.2f}")
        
    except Exception as e:
        consulta["status"] = "error"
        consulta["erro"] = str(e)
        print(f"❌ Erro na consulta {consulta_id}: {e}")

# ================= ENDPOINTS =================

@app.post("/consultas")
async def iniciar_consulta(
    request: IniciarConsultaRequest, 
    background_tasks: BackgroundTasks
):
    """Inicia uma nova consulta de veículos"""
    
    consulta_id = str(uuid.uuid4())
    
    print(f"\n{'='*50}")
    print(f"🚀 NOVA CONSULTA INICIADA: {consulta_id}")
    print(f"📋 Veículos: {len(request.veiculos)}")
    for v in request.veiculos:
        print(f"   🚗 {v.placa} | RENAVAM: {v.renavam}")
    print(f"{'='*50}\n")
    
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


@app.get("/config/veiculos")
async def listar_veiculos_configurados():
    """Retorna os veículos configurados em detran_manual.py (VEICULOS)."""
    return detran_manual.VEICULOS

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
            detail=f"Consulta ainda não foi concluída. Status atual: {consulta['status']}"
        )
    
    # Pegar data de hoje para PDFs
    data_hoje = datetime.now().strftime("%d-%m-%Y")
    pasta_pdfs = os.path.join("boletos", data_hoje)
    pdf_paths = []
    
    if os.path.exists(pasta_pdfs):
        pdf_paths = [f for f in os.listdir(pasta_pdfs) if f.endswith('.pdf')]
    
    return {
        "id": consulta_id,
        "multas": consulta["multas"],
        "total_multas": consulta["total_multas"],
        "valor_total": consulta["valor_total"],
        "excel_path": consulta["excel_path"],
        "pdf_paths": pdf_paths
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
    
    # Buscar PDF na pasta boletos/DD-MM-YYYY
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
    return {
        "status": "ok",
        "timestamp": datetime.now().isoformat(),
        "consultas_ativas": len(consultas_db)
    }

@app.get("/")
async def root():
    """Endpoint raiz"""
    return {
        "message": "DETRAN-CE API",
        "version": "1.0.0",
        "docs": "/docs",
        "health": "/health"
    }

# ================= STARTUP =================

@app.on_event("startup")
async def startup_event():
    print("\n" + "="*60)
    print("🚀 DETRAN-CE API INICIADA")
    print("="*60)
    print(f"📡 Servidor: http://localhost:8000")
    print(f"📚 Documentação: http://localhost:8000/docs")
    print(f"🏥 Health Check: http://localhost:8000/health")
    print(f"🌐 Frontend: http://localhost:3000")
    print("="*60 + "\n")

# ================= RODAR SERVIDOR =================

if __name__ == "__main__":
    import uvicorn
    uvicorn.run(
        app, 
        host="0.0.0.0", 
        port=8000,
        log_level="info"
    )
