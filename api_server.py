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

# Importar módulos de scraping e Excel
import detran_manual
from detran_manual import salvar_no_excel
import detran_scraper

app = FastAPI(title="DETRAN-CE API", version="1.0.0")


def _build_allowed_origins() -> List[str]:
    origins = {
        "http://localhost:3000",
        "http://localhost:3001",
        "http://localhost:3002",
        "https://samuelforte.github.io",
        "https://samuelforte.github.io/Playwright",
    }

    extra_origins = os.getenv("ALLOWED_ORIGINS", "")
    if extra_origins:
        for origin in extra_origins.split(","):
            normalized = origin.strip().rstrip("/")
            if normalized:
                origins.add(normalized)

    return sorted(origins)

# ===== Supabase Client =====
SUPABASE_URL = os.getenv("SUPABASE_URL") or os.getenv("NEXT_PUBLIC_SUPABASE_URL")
SUPABASE_KEY = (
    os.getenv("SUPABASE_SERVICE_KEY")
    or os.getenv("SUPABASE_SERVICE_ROLE_KEY")
    or os.getenv("SUPABASE_KEY")
    or os.getenv("NEXT_PUBLIC_SUPABASE_ANON_KEY")
    or os.getenv("NEXT_PUBLIC_SUPABASE_PUBLISHABLE_DEFAULT_KEY")
)

supabase: Optional[Client] = None
SUPABASE_ENABLED = False

if SUPABASE_URL and SUPABASE_KEY:
    try:
        supabase = create_client(SUPABASE_URL, SUPABASE_KEY)
        SUPABASE_ENABLED = True
    except Exception as e:
        print(f"⚠️ Supabase indisponível na inicialização. Usando memória local. Motivo: {e}")
else:
    print("⚠️ Supabase não configurado. Usando memória local para consultas.")

CONSULTAS_MEM: Dict[str, Dict[str, Any]] = {}
VEICULOS_CONSULTA_MEM: Dict[str, List[Dict[str, Any]]] = {}
MULTAS_MEM: Dict[str, List[Dict[str, Any]]] = {}

# CORS para permitir frontend
app.add_middleware(
    CORSMiddleware,
    allow_origins=_build_allowed_origins(),
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

def _using_memory_db() -> bool:
    return not SUPABASE_ENABLED or supabase is None


def db_insert_consulta(consulta_id: str, veiculos: List[Veiculo]):
    global SUPABASE_ENABLED
    now_iso = datetime.now().isoformat()
    consulta_payload = {
        "id": consulta_id,
        "status": "pending",
        "total_multas": 0,
        "valor_total": 0.0,
        "created_at": now_iso,
        "excel_path": None,
    }
    veiculos_rows = [{
        "consulta_id": consulta_id,
        "placa": v.placa,
        "status": "pending",
        "multas_count": 0,
        "valor_total": 0.0,
        "mensagem": "Aguardando processamento",
    } for v in veiculos]

    if _using_memory_db():
        CONSULTAS_MEM[consulta_id] = consulta_payload
        VEICULOS_CONSULTA_MEM[consulta_id] = veiculos_rows
        MULTAS_MEM[consulta_id] = []
        return

    # Cria consulta
    try:
        supabase.table("consultas").insert(consulta_payload).execute()
    except Exception as e:
        print(f"⚠️ Falha no Supabase ao criar consulta. Fallback memória: {e}")
        SUPABASE_ENABLED = False
        CONSULTAS_MEM[consulta_id] = consulta_payload
        VEICULOS_CONSULTA_MEM[consulta_id] = veiculos_rows
        MULTAS_MEM[consulta_id] = []
        return

    try:
        supabase.table("veiculos_consulta").insert(veiculos_rows).execute()
    except Exception as e:
        print(f"⚠️ Falha no Supabase ao salvar veículos. Fallback memória: {e}")
        SUPABASE_ENABLED = False
        CONSULTAS_MEM[consulta_id] = consulta_payload
        VEICULOS_CONSULTA_MEM[consulta_id] = veiculos_rows
        MULTAS_MEM[consulta_id] = []


def db_update_consulta_status(consulta_id: str, status: str, excel_path: Optional[str] = None,
                              total_multas: Optional[int] = None, valor_total: Optional[float] = None):
    global SUPABASE_ENABLED
    payload: Dict[str, Any] = {"status": status}
    if excel_path is not None:
        payload["excel_path"] = excel_path
    if total_multas is not None:
        payload["total_multas"] = total_multas
    if valor_total is not None:
        payload["valor_total"] = valor_total

    if _using_memory_db():
        consulta = CONSULTAS_MEM.get(consulta_id)
        if not consulta:
            raise HTTPException(status_code=404, detail="Consulta não encontrada")
        consulta.update(payload)
        return

    try:
        supabase.table("consultas").update(payload).eq("id", consulta_id).execute()
    except Exception as e:
        print(f"⚠️ Falha no Supabase ao atualizar consulta. Fallback memória: {e}")
        SUPABASE_ENABLED = False
        consulta = CONSULTAS_MEM.get(consulta_id)
        if not consulta:
            raise HTTPException(status_code=404, detail="Consulta não encontrada")
        consulta.update(payload)


def db_update_veiculo_status(consulta_id: str, placa: str, data: Dict[str, Any]):
    global SUPABASE_ENABLED
    if _using_memory_db():
        veiculos = VEICULOS_CONSULTA_MEM.get(consulta_id)
        if not veiculos:
            raise HTTPException(status_code=404, detail="Consulta não encontrada")
        for veiculo in veiculos:
            if veiculo.get("placa") == placa:
                veiculo.update(data)
                return
        raise HTTPException(status_code=404, detail=f"Veículo {placa} não encontrado")

    try:
        supabase.table("veiculos_consulta").update(data).eq("consulta_id", consulta_id).eq("placa", placa).execute()
    except Exception as e:
        print(f"⚠️ Falha no Supabase ao atualizar veículo. Fallback memória: {e}")
        SUPABASE_ENABLED = False
        veiculos = VEICULOS_CONSULTA_MEM.get(consulta_id)
        if not veiculos:
            raise HTTPException(status_code=404, detail="Consulta não encontrada")
        for veiculo in veiculos:
            if veiculo.get("placa") == placa:
                veiculo.update(data)
                return
        raise HTTPException(status_code=404, detail=f"Veículo {placa} não encontrado")


def db_insert_multas(consulta_id: str, multas: List[Dict]):
    global SUPABASE_ENABLED
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
            "valor": m.get("Valor") or m.get("valor") or m.get("valor_original") or "-",
            "valor_a_pagar": m.get("Valor a Pagar") or m.get("valor_a_pagar") or "-",
            "orgao_autuador": m.get("Órgão Autuador") or m.get("orgao_autuador") or m.get("orgao") or "-",
            "codigo_pagamento": m.get("Código de pagamento em barra") or m.get("codigo_pagamento") or "-",
        })

    if _using_memory_db():
        MULTAS_MEM.setdefault(consulta_id, []).extend(rows)
        return

    try:
        supabase.table("multas").insert(rows).execute()
    except Exception as e:
        print(f"⚠️ Falha no Supabase ao inserir multas. Fallback memória: {e}")
        SUPABASE_ENABLED = False
        MULTAS_MEM.setdefault(consulta_id, []).extend(rows)


def db_get_consulta_com_status(consulta_id: str) -> Dict[str, Any]:
    global SUPABASE_ENABLED
    if _using_memory_db():
        consulta = CONSULTAS_MEM.get(consulta_id)
        if not consulta:
            raise HTTPException(status_code=404, detail="Consulta não encontrada")
        return {
            "consulta": consulta,
            "veiculos": VEICULOS_CONSULTA_MEM.get(consulta_id, []),
        }

    try:
        consulta = supabase.table("consultas").select("*").eq("id", consulta_id).single().execute()
        veiculos = supabase.table("veiculos_consulta").select("*").eq("consulta_id", consulta_id).execute()
    except Exception as e:
        print(f"⚠️ Falha no Supabase ao obter consulta. Fallback memória: {e}")
        SUPABASE_ENABLED = False
        consulta_mem = CONSULTAS_MEM.get(consulta_id)
        if not consulta_mem:
            raise HTTPException(status_code=404, detail="Consulta não encontrada")
        return {
            "consulta": consulta_mem,
            "veiculos": VEICULOS_CONSULTA_MEM.get(consulta_id, []),
        }

    return {
        "consulta": consulta.data,
        "veiculos": veiculos.data or [],
    }


def db_get_multas(consulta_id: str) -> List[Dict[str, Any]]:
    global SUPABASE_ENABLED
    if _using_memory_db():
        return MULTAS_MEM.get(consulta_id, [])

    try:
        resp = supabase.table("multas").select("*").eq("consulta_id", consulta_id).execute()
    except Exception as e:
        print(f"⚠️ Falha no Supabase ao buscar multas. Fallback memória: {e}")
        SUPABASE_ENABLED = False
        return MULTAS_MEM.get(consulta_id, [])

    return resp.data or []


def db_get_historico() -> List[Dict[str, Any]]:
    global SUPABASE_ENABLED
    if _using_memory_db():
        historico = [
            {
                "id": consulta["id"],
                "status": consulta.get("status", "pending"),
                "total_multas": consulta.get("total_multas", 0),
                "valor_total": consulta.get("valor_total", 0.0),
                "created_at": consulta.get("created_at", ""),
            }
            for consulta in CONSULTAS_MEM.values()
        ]
        return sorted(historico, key=lambda item: item.get("created_at", ""), reverse=True)

    try:
        resp = supabase.table("consultas").select("id,status,total_multas,valor_total,created_at").order("created_at", desc=True).execute()
    except Exception as e:
        print(f"⚠️ Falha no Supabase ao buscar histórico. Fallback memória: {e}")
        SUPABASE_ENABLED = False
        historico = [
            {
                "id": consulta["id"],
                "status": consulta.get("status", "pending"),
                "total_multas": consulta.get("total_multas", 0),
                "valor_total": consulta.get("valor_total", 0.0),
                "created_at": consulta.get("created_at", ""),
            }
            for consulta in CONSULTAS_MEM.values()
        ]
        return sorted(historico, key=lambda item: item.get("created_at", ""), reverse=True)

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

def _multa_para_excel(multa: Dict, placa: str, numero: int) -> Dict:
    """Converte formato do scraper para o formato esperado por salvar_no_excel."""
    return {
        "Placa": placa,
        "#": numero,
        "AIT": multa.get("ait", "-"),
        "AIT Originária": multa.get("ait_originaria", "-"),
        "Motivo": multa.get("motivo", "-"),
        "Data Infração": multa.get("data_infracao", "-"),
        "Data Vencimento": multa.get("data_vencimento", "-"),
        "Valor": multa.get("valor_original", "-"),
        "Valor a Pagar": multa.get("valor_a_pagar", "-"),
        "Órgão Autuador": multa.get("orgao", "-"),
        "Código de pagamento em barra": "-",
    }


def processar_consulta_background(consulta_id: str, veiculos: List[Veiculo]):
    """Processa veículos em background usando requests + BeautifulSoup."""
    total_geral = 0.0
    todas_multas_excel = []  # formato capital para salvar_no_excel

    try:
        db_update_consulta_status(consulta_id, "processing")

        for veiculo_data in veiculos:
            db_update_veiculo_status(consulta_id, veiculo_data.placa, {
                "status": "processing",
                "mensagem": "Consultando DETRAN-CE...",
            })

            try:
                resultado = detran_scraper.consultar_multas(
                    veiculo_data.placa, veiculo_data.renavam
                )

                if resultado.get("erro"):
                    db_update_veiculo_status(consulta_id, veiculo_data.placa, {
                        "status": "error",
                        "mensagem": resultado["erro"],
                    })
                    continue

                multas = resultado["multas"]
                valor_veiculo = sum(m.get("valor_numerico", 0.0) for m in multas)

                # Adiciona placa a cada multa para o banco
                multas_com_placa = [{**m, "placa": veiculo_data.placa} for m in multas]
                db_insert_multas(consulta_id, multas_com_placa)

                # Converte para formato Excel
                for i, m in enumerate(multas, 1):
                    todas_multas_excel.append(_multa_para_excel(m, veiculo_data.placa, i))

                total_geral += valor_veiculo
                db_update_veiculo_status(consulta_id, veiculo_data.placa, {
                    "status": "completed",
                    "multas_count": len(multas),
                    "valor_total": valor_veiculo,
                    "mensagem": f"{len(multas)} multa(s) encontrada(s)",
                })

            except Exception as e:
                db_update_veiculo_status(consulta_id, veiculo_data.placa, {
                    "status": "error",
                    "mensagem": f"Erro: {str(e)}",
                })
                print(f"Erro ao processar {veiculo_data.placa}:")
                print(traceback.format_exc())

        if todas_multas_excel:
            salvar_no_excel(todas_multas_excel)
            excel_path = detran_manual.EXCEL_ARQUIVO
        else:
            excel_path = None

        db_update_consulta_status(
            consulta_id,
            status="completed",
            excel_path=excel_path,
            total_multas=len(todas_multas_excel),
            valor_total=total_geral,
        )

        print(f"Consulta {consulta_id} concluída: {len(todas_multas_excel)} multas | R$ {total_geral:.2f}")

    except Exception:
        db_update_consulta_status(consulta_id, "error")
        print(f"Erro na consulta {consulta_id}:")
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
