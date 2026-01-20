'use client';

import React, { useEffect, useState } from 'react';
import { Grid, Typography, Box } from '@mui/material';
import DirectionsCarIcon from '@mui/icons-material/DirectionsCar';
import WarningAmberIcon from '@mui/icons-material/WarningAmber';
import AttachMoneyIcon from '@mui/icons-material/AttachMoney';
import PictureAsPdfIcon from '@mui/icons-material/PictureAsPdf';
import Layout from '@/components/Layout';
import StatusCard from '@/components/StatusCard';

export default function Dashboard() {
  const [stats, setStats] = useState({
    veiculosConsultados: 0,
    multasEncontradas: 0,
    valorTotal: 0,
    pdfsGerados: 0,
  });

  useEffect(() => {
    // Aqui você pode buscar estatísticas reais da API
    // Por enquanto, usando dados de exemplo
    setStats({
      veiculosConsultados: 0,
      multasEncontradas: 0,
      valorTotal: 0,
      pdfsGerados: 0,
    });
  }, []);

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
          Dashboard
        </Typography>
        <Typography variant="body1" color="text.secondary">
          Resumo das consultas realizadas hoje
        </Typography>
      </Box>

      <Grid container spacing={3}>
        <Grid item xs={12} sm={6} md={3}>
          <StatusCard
            title="Veículos Consultados Hoje"
            value={stats.veiculosConsultados}
            icon={<DirectionsCarIcon />}
            color="primary"
          />
        </Grid>
        <Grid item xs={12} sm={6} md={3}>
          <StatusCard
            title="Multas Encontradas"
            value={stats.multasEncontradas}
            icon={<WarningAmberIcon />}
            color="warning"
          />
        </Grid>
        <Grid item xs={12} sm={6} md={3}>
          <StatusCard
            title="Valor Total"
            value={formatarValor(stats.valorTotal)}
            icon={<AttachMoneyIcon />}
            color="error"
          />
        </Grid>
        <Grid item xs={12} sm={6} md={3}>
          <StatusCard
            title="PDFs Gerados"
            value={stats.pdfsGerados}
            icon={<PictureAsPdfIcon />}
            color="success"
          />
        </Grid>
      </Grid>

      <Box sx={{ mt: 6, textAlign: 'center', py: 8 }}>
        <Typography variant="h5" color="text.secondary" gutterBottom>
          Nenhuma consulta realizada hoje
        </Typography>
        <Typography variant="body2" color="text.secondary">
          Clique em "Nova Consulta" para começar
        </Typography>
      </Box>
    </Layout>
  );
}
