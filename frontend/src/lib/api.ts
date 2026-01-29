import axios from 'axios';
import { supabase } from './supabaseClient';

// Usar API Routes do Vercel como proxy (no servidor)
// Em desenvolvimento usa localhost diretamente
const API_BASE_URL = typeof window === 'undefined' 
  ? (process.env.BACKEND_URL || 'http://localhost:8000')
  : (process.env.NEXT_PUBLIC_API_URL || '/api/proxy');

export const api = axios.create({
  baseURL: API_BASE_URL,
  headers: {
    'Content-Type': 'application/json',
  },
});

export interface Veiculo {
  placa: string;
  renavam: string;
}

export interface Multa {
  placa: string;
  numero: number;
  ait: string;
  ait_originaria: string;
  motivo: string;
  data_infracao: string;
  data_vencimento: string;
  valor: string;
  valor_a_pagar: string;
  orgao_autuador: string;
  codigo_pagamento: string;
}

export interface ConsultaStatus {
  id: string;
  status: 'pending' | 'processing' | 'completed' | 'error';
  veiculos: VeiculoStatus[];
  total_multas: number;
  valor_total: number;
  created_at: string;
}

export interface VeiculoStatus {
  placa: string;
  status: 'pending' | 'processing' | 'completed' | 'error';
  multas_count: number;
  valor_total: number;
  mensagem?: string;
}

export interface ConsultaResultado {
  id: string;
  multas: Multa[];
  excel_path?: string;
  pdf_paths?: string[];
  total_multas: number;
  valor_total: number;
}

// ===== Novos modelos para APIs auxiliares =====
export interface Condutor {
  id: string;
  nome: string;
  cpf: string;
  cnh_categoria?: string;
  cnh_vencimento?: string;
  pontuacao?: number;
}

export interface VeiculoRegistro {
  placa: string;
  renavam: string;
  condutorId?: string;
  ativo?: boolean;
}

export interface IndicacaoRequest {
  ait: string;
  placa: string;
  condutorId: string;
  data_indicacao?: string;
}

export interface PagamentoRequest {
  ait: string;
  placa: string;
  codigo_pagamento?: string;
  valor?: number;
}

export interface PagamentoRegistro {
  id: string;
  ait: string;
  placa: string;
  valor_original?: number;
  valor_pago: number;
  desconto_aplicado: boolean;
  codigo_pagamento?: string;
  status: string;
  data_pagamento: string;
}

export interface DebitoItem {
  tipo: string;
  descricao: string;
  valor: string;
  vencimento?: string;
  situacao: string;
}

export interface DebitoResumo {
  placa: string;
  debitos: DebitoItem[];
}

// Obter veículos configurados no backend (detran_manual.py)
export const obterVeiculosConfig = async () => {
  const response = await api.get<Veiculo[]>('/config/veiculos');
  return response.data;
};

// Iniciar nova consulta
export const iniciarConsulta = async (veiculos: Veiculo[]) => {
  const response = await api.post<{ consulta_id: string }>('/consultas', { veiculos });
  return response.data;
};

// Obter status da consulta
export const obterStatus = async (consultaId: string) => {
  const response = await api.get<ConsultaStatus>(`/consultas/${consultaId}/status`);
  return response.data;
};

// Obter resultado da consulta
export const obterResultado = async (consultaId: string) => {
  const response = await api.get<ConsultaResultado>(`/consultas/${consultaId}/resultado`);
  return response.data;
};

// Baixar Excel
export const baixarExcel = async (consultaId: string) => {
  const response = await api.get(`/consultas/${consultaId}/excel`, {
    responseType: 'blob',
  });
  
  const url = window.URL.createObjectURL(new Blob([response.data]));
  const link = document.createElement('a');
  link.href = url;
  link.setAttribute('download', `resultado_detran_${consultaId}.xlsx`);
  document.body.appendChild(link);
  link.click();
  link.remove();
};

// Baixar PDF individual
export const baixarPDF = async (consultaId: string, pdfFilename: string) => {
  const response = await api.get(`/consultas/${consultaId}/pdf/${pdfFilename}`, {
    responseType: 'blob',
  });
  
  const url = window.URL.createObjectURL(new Blob([response.data]));
  const link = document.createElement('a');
  link.href = url;
  link.setAttribute('download', pdfFilename);
  document.body.appendChild(link);
  link.click();
  link.remove();
};

// Listar histórico de consultas
export const listarHistorico = async () => {
  const response = await api.get<ConsultaStatus[]>('/consultas/historico');
  return response.data;
};

// ===== Condutores =====
export const listarCondutores = async () => {
  const { data, error } = await supabase
    .from('condutores')
    .select('*')
    .order('created_at', { ascending: false });

  if (error) throw error;
  return (data || []) as Condutor[];
};

export const criarCondutor = async (condutor: Omit<Condutor, 'id'>) => {
  const { data, error } = await supabase
    .from('condutores')
    .insert(condutor)
    .select('*')
    .single();

  if (error) throw error;
  return data as Condutor;
};

export const removerCondutor = async (condutorId: string) => {
  const { error } = await supabase
    .from('condutores')
    .delete()
    .eq('id', condutorId);

  if (error) throw error;
  return { status: 'ok' };
};

// ===== Veículos (registro) =====
export const listarVeiculosRegistro = async () => {
  const response = await api.get<VeiculoRegistro[]>('/veiculos-registro');
  return response.data;
};

export const salvarVeiculoRegistro = async (veiculo: VeiculoRegistro) => {
  const response = await api.post<VeiculoRegistro>('/veiculos-registro', veiculo);
  return response.data;
};

// ===== Indicações =====
export const registrarIndicacao = async (payload: IndicacaoRequest) => {
  const { data, error } = await supabase
    .from('indicacoes')
    .insert({
      ait: payload.ait,
      placa: payload.placa,
      condutor_id: payload.condutorId,
      data_indicacao: payload.data_indicacao || new Date().toISOString(),
      status: 'registrado',
    })
    .select('*')
    .single();

  if (error) throw error;
  return data as { status: string } & IndicacaoRequest & { data_indicacao: string };
};

// ===== Pagamentos =====
export const pagarMulta = async (payload: PagamentoRequest) => {
  const response = await api.post<PagamentoRegistro>('/pagamentos', payload);
  return response.data;
};

export const pagarMultaEmLote = async (pagamentos: PagamentoRequest[]) => {
  const response = await api.post<{ total: number; registros: PagamentoRegistro[] }>(
    '/pagamentos/lote',
    { pagamentos }
  );
  return response.data;
};

// ===== Débitos =====
export const listarDebitos = async (placa: string) => {
  const response = await api.get<DebitoResumo>(`/debitos/${placa}`);
  return response.data;
};

// ===== Emissão de PDF de multa (após pagamento) =====
export const emitirPdfMulta = async (placa: string, renavam: string, ait: string) => {
  const response = await api.post<{ filename: string; pdf_path: string; codigo_pagamento?: string }>(
    '/pagamentos/emitir-pdf',
    { placa, renavam, ait }
  );
  return response.data;
};

export const baixarBoletoPdf = async (filename: string) => {
  const response = await api.get(`/boletos/pdf/${filename}`, { responseType: 'blob' });
  const url = window.URL.createObjectURL(new Blob([response.data]));
  const link = document.createElement('a');
  link.href = url;
  link.setAttribute('download', filename);
  document.body.appendChild(link);
  link.click();
  link.remove();
};

export const getBoletoPdfUrl = (filename: string) => `${API_BASE_URL}/boletos/pdf/${filename}`;
