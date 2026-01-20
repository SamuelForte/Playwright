'use client';

import React, { useMemo, useState } from 'react';
import { Box, Button, Typography, Paper, Alert } from '@mui/material';
import { useRouter } from 'next/navigation';
import PlayArrowIcon from '@mui/icons-material/PlayArrow';
import Layout from '@/components/Layout';
import { Veiculo, iniciarConsulta } from '@/lib/api';

export default function NovaConsulta() {
  const router = useRouter();
  const [loading, setLoading] = useState(false);
  const [erro, setErro] = useState('');

  // Lista fixa de veículos (sincronizada com VEICULOS do detran_manual.py)
  const veiculosFixos = useMemo<Veiculo[]>(
    () => [
      { placa: 'SBA7F09', renavam: '01365705622' },
      { placa: 'TIF1J98', renavam: '01450499292' },
    ],
    []
  );

  const handleIniciarConsulta = async () => {
    setErro('');
    setLoading(true);

    try {
      const { consulta_id } = await iniciarConsulta(veiculosFixos);
      router.push(`/processamento/${consulta_id}`);
    } catch (error: any) {
      setErro(error.response?.data?.detail || 'Erro ao iniciar consulta');
      setLoading(false);
    }
  };

  return (
    <Layout>
      <Box sx={{ mb: 4 }}>
        <Typography variant="h4" gutterBottom fontWeight={600}>
          Nova Consulta
        </Typography>
        <Typography variant="body1" color="text.secondary">
          Consulta automática usando os veículos configurados na automação
        </Typography>
      </Box>

      {erro && (
        <Alert severity="error" sx={{ mb: 3 }} onClose={() => setErro('')}>
          {erro}
        </Alert>
      )}

      <Paper elevation={2} sx={{ mb: 4, p: 3 }}>
        <Typography variant="subtitle1" fontWeight={600} gutterBottom>
          Veículos prontos para consulta
        </Typography>
        {veiculosFixos.map((v, index) => (
          <Box key={`${v.placa}-${v.renavam}-${index}`} sx={{ display: 'flex', gap: 2, py: 0.5 }}>
            <Typography fontWeight={600}>Placa:</Typography>
            <Typography>{v.placa}</Typography>
            <Typography fontWeight={600}>Renavam:</Typography>
            <Typography>{v.renavam}</Typography>
          </Box>
        ))}
      </Paper>

      <Box sx={{ display: 'flex', justifyContent: 'center', gap: 2 }}>
        <Button
          variant="contained"
          size="large"
          startIcon={<PlayArrowIcon />}
          onClick={handleIniciarConsulta}
          disabled={loading || veiculosFixos.length === 0}
          sx={{ minWidth: 250, py: 1.5 }}
        >
          {loading ? 'Iniciando...' : 'Iniciar Consulta Automática'}
        </Button>
      </Box>

      <Box sx={{ mt: 2, textAlign: 'center' }}>
        <Typography variant="body2" color="text.secondary">
          {veiculosFixos.length} veículo(s) serão consultado(s) automaticamente
        </Typography>
      </Box>
    </Layout>
  );
}
