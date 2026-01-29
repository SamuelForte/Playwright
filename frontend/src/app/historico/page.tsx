'use client';

import React from 'react';
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
  IconButton,
  Chip,
} from '@mui/material';
import { useQuery } from '@tanstack/react-query';
import { useRouter } from 'next/navigation';
import VisibilityIcon from '@mui/icons-material/Visibility';
import { format } from 'date-fns';
import { ptBR } from 'date-fns/locale';
import Layout from '@/components/Layout';
import { listarHistorico } from '@/lib/api';

export default function Historico() {
  const router = useRouter();

  const { data: consultas, isLoading } = useQuery({
    queryKey: ['historico'],
    queryFn: listarHistorico,
  });

  const handleVerConsulta = (consultaId: string) => {
    router.push(`/resultados/${consultaId}`);
  };

  const formatarData = (dataString: string) => {
    try {
      const data = new Date(dataString);
      return format(data, "dd/MM/yyyy 'às' HH:mm", { locale: ptBR });
    } catch {
      return dataString;
    }
  };

  const formatarValor = (valor: number) => {
    return new Intl.NumberFormat('pt-BR', {
      style: 'currency',
      currency: 'BRL',
    }).format(valor);
  };

  const getStatusChip = (status: string) => {
    const configs = {
      completed: { label: 'Concluído', color: 'success' as const },
      processing: { label: 'Processando', color: 'primary' as const },
      error: { label: 'Erro', color: 'error' as const },
      pending: { label: 'Pendente', color: 'default' as const },
    };

    const config = configs[status as keyof typeof configs] || configs.pending;

    return <Chip label={config.label} color={config.color} size="small" />;
  };

  return (
    <Layout>
      <Box sx={{ mb: 4 }}>
        <Typography variant="h4" gutterBottom fontWeight={600}>
          Histórico de Consultas
        </Typography>
        <Typography variant="body1" color="text.secondary">
          Consultas realizadas anteriormente
        </Typography>
      </Box>

      <TableContainer component={Paper} elevation={2}>
        <Table>
          <TableHead>
            <TableRow sx={{ bgcolor: 'primary.main' }}>
              <TableCell sx={{ color: 'white', fontWeight: 600 }}>Data</TableCell>
              <TableCell sx={{ color: 'white', fontWeight: 600 }}>Veículos</TableCell>
              <TableCell sx={{ color: 'white', fontWeight: 600 }}>Multas</TableCell>
              <TableCell sx={{ color: 'white', fontWeight: 600 }}>Valor Total</TableCell>
              <TableCell sx={{ color: 'white', fontWeight: 600 }}>Status</TableCell>
              <TableCell sx={{ color: 'white', fontWeight: 600, textAlign: 'center' }}>
                Ações
              </TableCell>
            </TableRow>
          </TableHead>
          <TableBody>
            {isLoading ? (
              <TableRow>
                <TableCell colSpan={6} align="center" sx={{ py: 4 }}>
                  <Typography color="text.secondary">Carregando...</Typography>
                </TableCell>
              </TableRow>
            ) : !consultas || consultas.length === 0 ? (
              <TableRow>
                <TableCell colSpan={6} align="center" sx={{ py: 4 }}>
                  <Typography color="text.secondary">
                    Nenhuma consulta realizada ainda
                  </Typography>
                </TableCell>
              </TableRow>
            ) : (
              consultas.map((consulta) => (
                <TableRow key={consulta.id} hover>
                  <TableCell>{formatarData(consulta.created_at)}</TableCell>
                  <TableCell>{consulta.veiculos?.length || 0}</TableCell>
                  <TableCell>{consulta.total_multas}</TableCell>
                  <TableCell sx={{ fontWeight: 600 }}>
                    {formatarValor(consulta.valor_total)}
                  </TableCell>
                  <TableCell>{getStatusChip(consulta.status)}</TableCell>
                  <TableCell align="center">
                    <IconButton
                      color="primary"
                      onClick={() => handleVerConsulta(consulta.id)}
                      disabled={consulta.status !== 'completed'}
                    >
                      <VisibilityIcon />
                    </IconButton>
                  </TableCell>
                </TableRow>
              ))
            )}
          </TableBody>
        </Table>
      </TableContainer>
    </Layout>
  );
}
