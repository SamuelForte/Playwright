'use client';

import React, { useEffect, useState } from 'react';
import { useParams, useRouter } from 'next/navigation';
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
} from '@mui/material';
import { useQuery } from '@tanstack/react-query';
import CheckCircleIcon from '@mui/icons-material/CheckCircle';
import Layout from '@/components/Layout';
import ProcessStatus from '@/components/ProcessStatus';
import { obterStatus, VeiculoStatus } from '@/lib/api';

export default function Processamento() {
  const params = useParams();
  const router = useRouter();
  const consultaId = params.id as string;

  const { data, isError, error } = useQuery({
    queryKey: ['status', consultaId],
    queryFn: () => obterStatus(consultaId),
    refetchInterval: (data: any) => {
      // Para de atualizar quando completar ou dar erro
      if (!data || data.status === 'completed' || data.status === 'error') {
        return false;
      }
      return 2000; // Atualiza a cada 2 segundos
    },
    enabled: !!consultaId,
  });

  const handleVerResultados = () => {
    router.push(`/resultados/${consultaId}`);
  };

  if (isError) {
    return (
      <Layout>
        <Alert severity="error">
          Erro ao carregar status: {(error as Error).message}
        </Alert>
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
                    <ProcessStatus
                      status={veiculo.status}
                      mensagem={veiculo.mensagem}
                    />
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
