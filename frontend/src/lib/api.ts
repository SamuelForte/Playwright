import axios from 'axios';

const API_BASE_URL = process.env.NEXT_PUBLIC_API_URL || 'http://localhost:8000';

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
