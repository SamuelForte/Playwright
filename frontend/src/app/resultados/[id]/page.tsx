'use client';

import React from 'react';
import { useParams } from 'next/navigation';
import {
  Box,
  Typography,
  Button,
  Alert,
  Grid,
  Paper,
} from '@mui/material';
import { useQuery } from '@tanstack/react-query';
import DownloadIcon from '@mui/icons-material/Download';
import PictureAsPdfIcon from '@mui/icons-material/PictureAsPdf';
import TableViewIcon from '@mui/icons-material/TableView';
import Layout from '@/components/Layout';
import MultasTable from '@/components/MultasTable';
import StatusCard from '@/components/StatusCard';
import { obterResultado, baixarExcel, baixarPDF } from '@/lib/api';
import WarningAmberIcon from '@mui/icons-material/WarningAmber';
import AttachMoneyIcon from '@mui/icons-material/AttachMoney';

export default function Resultados() {
  const params = useParams();
  const consultaId = params.id as string;

  const { data, isLoading, isError, error } = useQuery({
    queryKey: ['resultado', consultaId],
    queryFn: () => obterResultado(consultaId),
    enabled: !!consultaId,
  });

  const handleBaixarExcel = async () => {
    try {
      await baixarExcel(consultaId);
    } catch (error) {
      console.error('Erro ao baixar Excel:', error);
    }
  };

  const handleBaixarPDF = async (ait: string) => {
    try {
      // Implementar lógica para encontrar o PDF correto pelo AIT
      // Por enquanto, apenas um exemplo
      if (data?.pdf_paths && data.pdf_paths.length > 0) {
        await baixarPDF(consultaId, data.pdf_paths[0]);
      }
    } catch (error) {
      console.error('Erro ao baixar PDF:', error);
    }
  };

  if (isLoading) {
    return (
      <Layout>
        <Box sx={{ textAlign: 'center', py: 8 }}>
          <Typography>Carregando resultados...</Typography>
        </Box>
      </Layout>
    );
  }

  if (isError) {
    return (
      <Layout>
        <Alert severity="error">
          Erro ao carregar resultados: {(error as Error).message}
        </Alert>
      </Layout>
    );
  }

  const multas = data?.multas || [];
  const totalMultas = data?.total_multas || 0;
  const valorTotal = data?.valor_total || 0;

  const formatarValor = (valor: number) => {
    return new Intl.NumberFormat('pt-BR', {
      style: 'currency',
      currency: 'BRL',
    }).format(valor);
  };

  return (
    <Layout>
      <Box sx={{ mb: 4 }}>
        <Typography variant="h4" gutterBottom fontWeight={600}>
          Resultados da Consulta
        </Typography>
        <Typography variant="body1" color="text.secondary">
          Resumo e detalhamento das multas encontradas
        </Typography>
      </Box>

      <Grid container spacing={3} sx={{ mb: 4 }}>
        <Grid item xs={12} md={4}>
          <StatusCard
            title="Multas Encontradas"
            value={totalMultas}
            icon={<WarningAmberIcon />}
            color="warning"
          />
        </Grid>
        <Grid item xs={12} md={4}>
          <StatusCard
            title="Valor Total"
            value={formatarValor(valorTotal)}
            icon={<AttachMoneyIcon />}
            color="error"
          />
        </Grid>
        <Grid item xs={12} md={4}>
          <StatusCard
            title="PDFs Gerados"
            value={data?.pdf_paths?.length || 0}
            icon={<PictureAsPdfIcon />}
            color="success"
          />
        </Grid>
      </Grid>

      <Paper elevation={2} sx={{ p: 3, mb: 3 }}>
        <Box sx={{ display: 'flex', gap: 2, flexWrap: 'wrap' }}>
          <Button
            variant="contained"
            startIcon={<TableViewIcon />}
            onClick={handleBaixarExcel}
            size="large"
          >
            Baixar Excel
          </Button>
          <Button
            variant="outlined"
            startIcon={<PictureAsPdfIcon />}
            size="large"
            disabled={!data?.pdf_paths || data.pdf_paths.length === 0}
          >
            Baixar Todos os PDFs
          </Button>
        </Box>
      </Paper>

      {multas.length === 0 ? (
        <Paper elevation={2} sx={{ p: 6, textAlign: 'center' }}>
          <Typography variant="h6" color="text.secondary">
            Nenhuma multa encontrada
          </Typography>
        </Paper>
      ) : (
        <MultasTable multas={multas} onDownloadPDF={handleBaixarPDF} />
      )}
    </Layout>
  );
}
