from fastapi import FastAPI, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse
from pydantic import BaseModel
from typing import List, Dict, Optional, Any
import uuid
import threading
from datetime import datetime
import os
import sys
import traceback
from dotenv import load_dotenv

# Carregar variáveis de ambiente
load_dotenv()

from supabase import create_client, Client

# Importar funções do detran_manual.py
import detran_manual
from detran_manual import processar_veiculo, salvar_no_excel
from playwright.sync_api import sync_playwright

app = FastAPI(title="DETRAN-CE API", version="1.0.0")

# ===== Supabase Client =====
SUPABASE_URL = os.getenv("SUPABASE_URL") or os.getenv("NEXT_PUBLIC_SUPABASE_URL")
SUPABASE_KEY = (
    os.getenv("SUPABASE_SERVICE_KEY")
    or os.getenv("SUPABASE_SERVICE_ROLE_KEY")
    or os.getenv("SUPABASE_KEY")
    or os.getenv("NEXT_PUBLIC_SUPABASE_ANON_KEY")
    or os.getenv("NEXT_PUBLIC_SUPABASE_PUBLISHABLE_DEFAULT_KEY")
)

if not SUPABASE_URL or not SUPABASE_KEY:
    raise RuntimeError("Supabase credentials missing. Defina SUPABASE_URL e SUPABASE_SERVICE_KEY/KEY/ANON/PUBLISHABLE.")

supabase: Client = create_client(SUPABASE_URL, SUPABASE_KEY)

# CORS para permitir frontend
app.add_middleware(
    CORSMiddleware,
    allow_origins=["http://localhost:3000", "http://localhost:3001"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# Armazenamento agora no Supabase (consultas, veiculos, multas, condutores, indicacoes)

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

class Condutor(BaseModel):
    id: str
    nome: str
    cpf: str
    cnh_categoria: Optional[str] = None
    cnh_vencimento: Optional[str] = None
    pontuacao: Optional[int] = None

class CondutorCreate(BaseModel):
    nome: str
    cpf: str
    cnh_categoria: Optional[str] = None
    cnh_vencimento: Optional[str] = None
    pontuacao: Optional[int] = None

class IndicacaoRequest(BaseModel):
    ait: str
    placa: str
    condutorId: str
    data_indicacao: Optional[str] = None

# ================= FUNÇÕES AUXILIARES =================

def _supabase_or_http_error():
    if not supabase:
        raise HTTPException(status_code=500, detail="Supabase client não inicializado")


def db_insert_consulta(consulta_id: str, veiculos: List[Veiculo]):
    _supabase_or_http_error()
    now_iso = datetime.now().isoformat()

    # Cria consulta
    try:
        resp = supabase.table("consultas").insert({
            "id": consulta_id,
            "status": "pending",
            "total_multas": 0,
            "valor_total": 0.0,
            "created_at": now_iso,
            "excel_path": None,
        }).execute()
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Erro ao criar consulta: {str(e)}")

    # Cria status dos veículos
    veiculos_rows = [{
        "consulta_id": consulta_id,
        "placa": v.placa,
        "status": "pending",
        "multas_count": 0,
        "valor_total": 0.0,
        "mensagem": "Aguardando processamento",
    } for v in veiculos]

    try:
        resp_veic = supabase.table("veiculos_consulta").insert(veiculos_rows).execute()
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Erro ao salvar veículos da consulta: {str(e)}")


def db_update_consulta_status(consulta_id: str, status: str, excel_path: Optional[str] = None,
                              total_multas: Optional[int] = None, valor_total: Optional[float] = None):
    _supabase_or_http_error()
    payload: Dict[str, Any] = {"status": status}
    if excel_path is not None:
        payload["excel_path"] = excel_path
    if total_multas is not None:
        payload["total_multas"] = total_multas
    if valor_total is not None:
        payload["valor_total"] = valor_total

    try:
        resp = supabase.table("consultas").update(payload).eq("id", consulta_id).execute()
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Erro ao atualizar consulta: {str(e)}")


def db_update_veiculo_status(consulta_id: str, placa: str, data: Dict[str, Any]):
    _supabase_or_http_error()
    try:
        resp = supabase.table("veiculos_consulta").update(data).eq("consulta_id", consulta_id).eq("placa", placa).execute()
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Erro ao atualizar veículo {placa}: {str(e)}")


def db_insert_multas(consulta_id: str, multas: List[Dict]):
    _supabase_or_http_error()
    if not multas:
        return
    rows = []
    for m in multas:
        rows.append({
            "consulta_id": consulta_id,
            "placa": m.get("Placa") or m.get("placa") or "-",
            "numero": m.get("#") or m.get("numero") or 0,
            "ait": m.get("AIT") or m.get("ait") or "-",
            "ait_originaria": m.get("AIT Originária") or m.get("ait_originaria") or "-",
            "motivo": m.get("Motivo") or m.get("motivo") or "-",
            "data_infracao": m.get("Data Infração") or m.get("data_infracao") or "-",
            "data_vencimento": m.get("Data Vencimento") or m.get("data_vencimento") or "-",
            "valor": m.get("Valor") or m.get("valor") or "-",
            "valor_a_pagar": m.get("Valor a Pagar") or m.get("valor_a_pagar") or "-",
            "orgao_autuador": m.get("Órgão Autuador") or m.get("orgao_autuador") or "-",
            "codigo_pagamento": m.get("Código de pagamento em barra") or m.get("codigo_pagamento") or "-",
        })

    try:
        resp = supabase.table("multas").insert(rows).execute()
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Erro ao salvar multas: {str(e)}")


def db_get_consulta_com_status(consulta_id: str) -> Dict[str, Any]:
    _supabase_or_http_error()
    try:
        consulta = supabase.table("consultas").select("*").eq("id", consulta_id).single().execute()
        veiculos = supabase.table("veiculos_consulta").select("*").eq("consulta_id", consulta_id).execute()
    except Exception as e:
        raise HTTPException(status_code=404, detail=f"Consulta não encontrada: {str(e)}")
    return {
        "consulta": consulta.data,
        "veiculos": veiculos.data or [],
    }


def db_get_multas(consulta_id: str) -> List[Dict[str, Any]]:
    _supabase_or_http_error()
    try:
        resp = supabase.table("multas").select("*").eq("consulta_id", consulta_id).execute()
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Erro ao buscar multas: {str(e)}")
    return resp.data or []


def db_get_historico() -> List[Dict[str, Any]]:
    _supabase_or_http_error()
    try:
        resp = supabase.table("consultas").select("id,status,total_multas,valor_total,created_at").order("created_at", desc=True).execute()
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Erro ao buscar histórico: {str(e)}")
    return resp.data or []

def converter_multa_para_frontend(multa_dict: Dict) -> Dict:
    """Converte dados da planilha (detran_manual.py) para formato do frontend"""
    return {
        "placa": multa_dict.get("Placa", "-"),
        "numero": multa_dict.get("#", 0),
        "ait": multa_dict.get("AIT", "-"),
        "ait_originaria": multa_dict.get("AIT Originária", "-"),
        "motivo": multa_dict.get("Motivo", "-"),
        "data_infracao": multa_dict.get("Data Infração", "-"),
        "data_vencimento": multa_dict.get("Data Vencimento", "-"),
        "valor": multa_dict.get("Valor", "-"),
        "valor_a_pagar": multa_dict.get("Valor a Pagar", "-"),
        "orgao_autuador": multa_dict.get("Órgão Autuador", "-"),
        "codigo_pagamento": multa_dict.get("Código de pagamento em barra", "-")
    }

def processar_consulta_background(consulta_id: str, veiculos: List[Veiculo]):
    """Processa veículos em background usando Playwright (detran_manual.py)"""
    total_geral = 0.0
    todas_multas_original = []  # Para Excel

    try:
        # Marca consulta como processing
        db_update_consulta_status(consulta_id, "processing")

        with sync_playwright() as p:
            browser = p.chromium.launch(headless=False)
            
            for i, veiculo_data in enumerate(veiculos, 1):
                # Atualiza status do veículo para processing
                db_update_veiculo_status(consulta_id, veiculo_data.placa, {
                    "status": "processing",
                    "mensagem": "Consultando DETRAN-CE...",
                })

                try:
                    # Chama a função processar_veiculo do detran_manual.py
                    veiculo_dict = {
                        "placa": veiculo_data.placa,
                        "renavam": veiculo_data.renavam
                    }
                    
                    total, multas = processar_veiculo(browser, veiculo_dict, i)
                    multas_formatadas = [converter_multa_para_frontend(m) for m in multas]

                    # Atualiza status do veículo
                    db_update_veiculo_status(consulta_id, veiculo_data.placa, {
                        "status": "completed",
                        "multas_count": len(multas),
                        "valor_total": total,
                        "mensagem": f"{len(multas)} multa(s) encontrada(s)",
                    })

                    # Persiste multas
                    db_insert_multas(consulta_id, multas)

                    todas_multas_original.extend(multas)
                    total_geral += total
                    
                except Exception as e:
                    db_update_veiculo_status(consulta_id, veiculo_data.placa, {
                        "status": "error",
                        "mensagem": f"Erro: {str(e)}",
                    })
                    print(f"❌ Erro ao processar {veiculo_data.placa}:")
                    print(traceback.format_exc())
            
            browser.close()
        
        # Salvar Excel usando a função do detran_manual.py
        if todas_multas_original:
            salvar_no_excel(todas_multas_original)
            excel_path = detran_manual.EXCEL_ARQUIVO
        else:
            excel_path = None

        # Marca consulta como concluída com totais
        db_update_consulta_status(
            consulta_id,
            status="completed",
            excel_path=excel_path,
            total_multas=len(todas_multas_original),
            valor_total=total_geral,
        )
        
        print(f"✅ Consulta {consulta_id} concluída com sucesso!")
        print(f"📊 Total: {len(todas_multas_original)} multas | R$ {total_geral:.2f}")
        
    except Exception:
        db_update_consulta_status(consulta_id, "error")
        print(f"❌ Erro na consulta {consulta_id}:")
        print(traceback.format_exc())

# ================= ENDPOINTS =================

@app.get("/condutores")
async def listar_condutores():
    """Lista condutores cadastrados"""
    try:
        resp = supabase.table("condutores").select("*").order("created_at", desc=True).execute()
        return resp.data or []
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

@app.post("/condutores")
async def criar_condutor(condutor: CondutorCreate):
    """Cria um novo condutor para indicação"""
    payload = {
        "nome": condutor.nome,
        "cpf": condutor.cpf,
        "cnh_categoria": condutor.cnh_categoria,
        "cnh_vencimento": condutor.cnh_vencimento,
        "pontuacao": condutor.pontuacao,
    }
    try:
        resp = supabase.table("condutores").insert(payload).select("*").single().execute()
        return resp.data
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

@app.delete("/condutores/{condutor_id}")
async def remover_condutor(condutor_id: str):
    """Remove um condutor cadastrado"""
    try:
        resp = supabase.table("condutores").delete().eq("id", condutor_id).execute()
        if not resp.data:
            raise HTTPException(status_code=404, detail="Condutor não encontrado")
        return {"status": "ok"}
    except HTTPException:
        raise
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

@app.get("/indicacoes")
async def listar_indicacoes():
    """Lista indicações registradas"""
    try:
        resp = supabase.table("indicacoes").select("*").order("data_indicacao", desc=True).execute()
        return resp.data or []
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

@app.post("/indicacoes")
async def registrar_indicacao(payload: IndicacaoRequest):
    """Registra indicação de condutor para uma multa (AIT)"""
    data_indicacao = payload.data_indicacao or datetime.now().isoformat()

    # Verifica condutor
    try:
        condutor = supabase.table("condutores").select("id").eq("id", payload.condutorId).single().execute()
    except Exception:
        raise HTTPException(status_code=404, detail="Condutor não encontrado")

    try:
        resp = supabase.table("indicacoes").insert({
            "ait": payload.ait,
            "placa": payload.placa,
            "condutor_id": payload.condutorId,
            "data_indicacao": data_indicacao,
            "status": "registrado",
        }).select("*").single().execute()
        return resp.data
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

@app.post("/consultas")
async def iniciar_consulta(request: IniciarConsultaRequest):
    """Inicia uma nova consulta de veículos"""
    
    consulta_id = str(uuid.uuid4())
    
    print(f"\n{'='*50}")
    print(f"🚀 NOVA CONSULTA INICIADA: {consulta_id}")
    print(f"📋 Veículos: {len(request.veiculos)}")
    for v in request.veiculos:
        print(f"   🚗 {v.placa} | RENAVAM: {v.renavam}")
    print(f"{'='*50}\n")
    
    # Cria registro da consulta e veículos no Supabase
    db_insert_consulta(consulta_id, request.veiculos)
    
    # Processa em thread separada (necessário para Playwright no Windows)
    thread = threading.Thread(
        target=processar_consulta_background,
        args=(consulta_id, request.veiculos),
        daemon=True
    )
    thread.start()
    
    return {"consulta_id": consulta_id}


@app.get("/config/veiculos")
async def listar_veiculos_configurados():
    """Retorna os veículos configurados em detran_manual.py (VEICULOS)."""
    return detran_manual.VEICULOS

@app.get("/consultas/{consulta_id}/status")
async def obter_status(consulta_id: str):
    """Retorna o status atual da consulta"""
    data = db_get_consulta_com_status(consulta_id)
    consulta = data["consulta"]
    veiculos = data["veiculos"]

    return ConsultaStatus(
        id=consulta["id"],
        status=consulta.get("status", "pending"),
        veiculos=[VeiculoStatus(**v) for v in veiculos],
        total_multas=consulta.get("total_multas", 0),
        valor_total=consulta.get("valor_total", 0.0),
        created_at=consulta.get("created_at", "")
    )

@app.get("/consultas/{consulta_id}/resultado")
async def obter_resultado(consulta_id: str):
    """Retorna o resultado completo da consulta"""
    data = db_get_consulta_com_status(consulta_id)
    consulta = data["consulta"]

    if consulta.get("status") != "completed":
        raise HTTPException(
            status_code=400,
            detail=f"Consulta ainda não foi concluída. Status atual: {consulta.get('status')}"
        )

    multas = db_get_multas(consulta_id)

    # Pegar data de hoje para PDFs
    data_hoje = datetime.now().strftime("%d-%m-%Y")
    pasta_pdfs = os.path.join("boletos", data_hoje)
    pdf_paths: List[str] = []

    if os.path.exists(pasta_pdfs):
        pdf_paths = [f for f in os.listdir(pasta_pdfs) if f.endswith('.pdf')]

    return {
        "id": consulta_id,
        "multas": multas,
        "total_multas": consulta.get("total_multas", 0),
        "valor_total": consulta.get("valor_total", 0.0),
        "excel_path": consulta.get("excel_path"),
        "pdf_paths": pdf_paths,
    }

@app.get("/consultas/{consulta_id}/excel")
async def baixar_excel(consulta_id: str):
    """Baixa o Excel da consulta"""
    data = db_get_consulta_com_status(consulta_id)
    consulta = data["consulta"]
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
    try:
        return db_get_historico()
    except HTTPException:
        raise
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

@app.get("/health")
async def health_check():
    """Health check da API"""
    try:
        historico = db_get_historico()
        total = len(historico)
    except Exception:
        total = None
    return {
        "status": "ok",
        "timestamp": datetime.now().isoformat(),
        "consultas_ativas": total,
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
