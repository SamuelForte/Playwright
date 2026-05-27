'use client';

import React, { Suspense } from 'react';
import { useRouter, useSearchParams } from 'next/navigation';
import {
  Box,
  Typography,
  Paper,
  Table,
  TableBody,
  TableCell,
  TableContainer,
  TableHead,
  TableRow,
  Button,
  Alert,
  LinearProgress,
  Divider,
  Chip,
} from '@mui/material';
import { useQuery } from '@tanstack/react-query';
import CheckCircleIcon from '@mui/icons-material/CheckCircle';
import Layout from '@/components/Layout';
import ProcessStatus from '@/components/ProcessStatus';
import { obterStatus, VeiculoStatus, LogEntry } from '@/lib/api';

function ProcessamentoContent() {
  const router = useRouter();
  const searchParams = useSearchParams();
  const consultaId = searchParams.get('id') || '';

  const { data, isError, error } = useQuery({
    queryKey: ['status', consultaId],
    queryFn: () => obterStatus(consultaId),
    refetchInterval: (queryData: any) => {
      if (!queryData || queryData.status === 'completed' || queryData.status === 'error') {
        return false;
      }
      return 2000;
    },
    enabled: !!consultaId,
    retry: 3,
  });

  const handleVerResultados = () => {
    router.push(`/resultados?id=${consultaId}`);
  };

  const logsFormatados = (data?.logs || []).map((log: LogEntry) => ({
    ...log,
    hora: new Date(log.timestamp).toLocaleTimeString('pt-BR', {
      hour: '2-digit',
      minute: '2-digit',
      second: '2-digit',
    }),
  }));

  const getLogColor = (nivel: LogEntry['nivel']) => {
    switch (nivel) {
      case 'error': return 'error';
      case 'warn': return 'warning';
      case 'success': return 'success';
      default: return 'info';
    }
  };

  if (!consultaId) {
    return (
      <Layout>
        <Alert severity="warning">Consulta não informada. Volte e inicie uma nova consulta.</Alert>
      </Layout>
    );
  }

  if (isError) {
    const errorMessage = error instanceof Error ? error.message : 'Erro desconhecido';
    return (
      <Layout>
        <Alert severity="error">Erro ao carregar status: {errorMessage}</Alert>
      </Layout>
    );
  }

  const statusGeral = data?.status || 'pending';
  const veiculos = data?.veiculos || [];
  const totalConcluidos = veiculos.filter(v => v.status === 'completed').length;
  const progresso = veiculos.length > 0 ? (totalConcluidos / veiculos.length) * 100 : 0;

  return (
    <Layout>
      <Box sx={{ mb: 4 }}>
        <Typography variant="h4" gutterBottom fontWeight={600}>
          Processamento em Andamento
        </Typography>
        <Typography variant="body1" color="text.secondary">
          Acompanhe o status da consulta em tempo real
        </Typography>
      </Box>

      <Box sx={{ mb: 4 }}>
        <ProcessStatus
          status={statusGeral}
          mensagem={
            statusGeral === 'processing'
              ? `Processando ${totalConcluidos} de ${veiculos.length} veículo(s)...`
              : statusGeral === 'completed'
              ? 'Consulta concluída com sucesso!'
              : statusGeral === 'error'
              ? 'Erro durante o processamento'
              : 'Aguardando início...'
          }
          progresso={progresso}
        />
      </Box>

      <Paper elevation={2} sx={{ mb: 4, p: 3 }}>
        <Box sx={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', gap: 2, mb: 2 }}>
          <Typography variant="h6" fontWeight={600}>
            Logs da extração
          </Typography>
          <Chip label={`${logsFormatados.length} evento(s)`} size="small" />
        </Box>

        <Divider sx={{ mb: 2 }} />

        {logsFormatados.length === 0 ? (
          <Typography variant="body2" color="text.secondary">
            Aguardando primeiras mensagens da extração...
          </Typography>
        ) : (
          <Box
            sx={{
              maxHeight: 360,
              overflowY: 'auto',
              display: 'flex',
              flexDirection: 'column',
              gap: 1,
              pr: 1,
            }}
          >
            {logsFormatados.map((log, index) => (
              <Box
                key={`${log.timestamp}-${index}`}
                sx={{
                  p: 1.25,
                  borderRadius: 2,
                  bgcolor: 'background.default',
                  border: '1px solid',
                  borderColor: 'divider',
                }}
              >
                <Box sx={{ display: 'flex', alignItems: 'center', gap: 1, mb: 0.5, flexWrap: 'wrap' }}>
                  <Chip
                    size="small"
                    label={log.nivel.toUpperCase()}
                    color={getLogColor(log.nivel)}
                  />
                  <Typography variant="caption" color="text.secondary">
                    {log.hora}
                  </Typography>
                </Box>
                <Typography variant="body2" sx={{ whiteSpace: 'pre-wrap', wordBreak: 'break-word' }}>
                  {log.mensagem}
                </Typography>
              </Box>
            ))}
          </Box>
        )}
      </Paper>

      <TableContainer component={Paper} elevation={2}>
        <Table>
          <TableHead>
            <TableRow sx={{ bgcolor: 'primary.main' }}>
              <TableCell sx={{ color: 'white', fontWeight: 600 }}>Placa</TableCell>
              <TableCell sx={{ color: 'white', fontWeight: 600 }}>Status</TableCell>
              <TableCell sx={{ color: 'white', fontWeight: 600 }}>Multas</TableCell>
              <TableCell sx={{ color: 'white', fontWeight: 600 }}>Valor Total</TableCell>
            </TableRow>
          </TableHead>
          <TableBody>
            {veiculos.length === 0 ? (
              <TableRow>
                <TableCell colSpan={4} align="center">
                  <LinearProgress sx={{ my: 2 }} />
                  <Typography color="text.secondary">Carregando...</Typography>
                </TableCell>
              </TableRow>
            ) : (
              veiculos.map((veiculo: VeiculoStatus, index: number) => (
                <TableRow key={index}>
                  <TableCell sx={{ fontWeight: 600 }}>{veiculo.placa}</TableCell>
                  <TableCell>
                    <ProcessStatus status={veiculo.status} mensagem={veiculo.mensagem} />
                  </TableCell>
                  <TableCell>{veiculo.multas_count || '-'}</TableCell>
                  <TableCell>
                    {veiculo.valor_total
                      ? new Intl.NumberFormat('pt-BR', {
                          style: 'currency',
                          currency: 'BRL',
                        }).format(veiculo.valor_total)
                      : '-'}
                  </TableCell>
                </TableRow>
              ))
            )}
          </TableBody>
        </Table>
      </TableContainer>

      {statusGeral === 'completed' && (
        <Box sx={{ mt: 4, textAlign: 'center' }}>
          <Button
            variant="contained"
            size="large"
            startIcon={<CheckCircleIcon />}
            onClick={handleVerResultados}
            sx={{ minWidth: 250 }}
          >
            Ver Resultados
          </Button>
        </Box>
      )}
    </Layout>
  );
}

export default function Processamento() {
  return (
    <Suspense fallback={<Layout><Alert severity="info">Carregando...</Alert></Layout>}>
      <ProcessamentoContent />
    </Suspense>
  );
}
