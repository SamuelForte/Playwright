'use client';

import React, { useEffect, useState } from 'react';
import { Box, Button, Typography, Paper, Alert } from '@mui/material';
import { useRouter } from 'next/navigation';
import PlayArrowIcon from '@mui/icons-material/PlayArrow';
import Layout from '@/components/Layout';
import { Veiculo, iniciarConsulta } from '@/lib/api';

export default function NovaConsulta() {
  const router = useRouter();
  const [loading, setLoading] = useState(false);
  const [erro, setErro] = useState('');
  const [veiculosAtivos, setVeiculosAtivos] = useState<Veiculo[]>([
    { placa: 'SBA7F09', renavam: '01365705622' },
    { placa: 'TIF1J98', renavam: '01450499292' },
  ]);

  // Carregar veículos do localStorage quando o componente monta
  useEffect(() => {
    const veiculosSalvos = localStorage.getItem('veiculos');
    if (veiculosSalvos) {
      try {
        const parsed = JSON.parse(veiculosSalvos);
        if (parsed.length > 0) {
          setVeiculosAtivos(parsed);
        }
      } catch (e) {
        console.error('Erro ao carregar veículos salvos:', e);
      }
    }
  }, []);

  const handleIniciarConsulta = async () => {
    setErro('');
    setLoading(true);

    try {
      console.log('Iniciando consulta com veículos:', veiculosAtivos);
      const { consulta_id } = await iniciarConsulta(veiculosAtivos);
      console.log('Consulta iniciada com sucesso. ID:', consulta_id);
      router.push(`/processamento/${consulta_id}`);
    } catch (error: any) {
      console.error('Erro ao iniciar consulta:', error);
      const mensagemErro = error.response?.data?.detail 
        || error.message 
        || 'Erro ao iniciar consulta. Verifique se o backend está rodando em http://localhost:8000';
      setErro(mensagemErro);
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
        {veiculosAtivos.length === 0 ? (
          <Typography color="text.secondary">
            Nenhum veículo cadastrado. Vá para o Dashboard e adicione veículos.
          </Typography>
        ) : (
          <Box sx={{ display: 'flex', flexDirection: 'column', gap: 1 }}>
            {veiculosAtivos.map((v, index) => (
              <Box
                key={`${v.placa}-${v.renavam}-${index}`}
                sx={{
                  border: '1px solid #e0e0e0',
                  borderRadius: '8px',
                  p: 0.75,
                  backgroundColor: '#f9f9f9',
                  display: 'flex',
                  flexDirection: 'column',
                  gap: 0.75,
                  '&:hover': {
                    backgroundColor: '#f5f5f5',
                    borderColor: '#d0d0d0',
                  },
                }}
              >
                <Box sx={{ display: 'flex', flexDirection: 'column', gap: 0.125 }}>
                  <Typography variant="caption" color="text.secondary" fontWeight={600} sx={{ fontSize: '0.65rem' }}>
                    Placa
                  </Typography>
                  <Typography variant="caption" fontWeight={600} sx={{ color: '#1976d2' }}>
                    {v.placa}
                  </Typography>
                </Box>
                <Box sx={{ display: 'flex', flexDirection: 'column', gap: 0.125 }}>
                  <Typography variant="caption" color="text.secondary" fontWeight={600} sx={{ fontSize: '0.65rem' }}>
                    Renavam
                  </Typography>
                  <Typography variant="caption" fontWeight={600} sx={{ color: '#1976d2' }}>
                    {v.renavam}
                  </Typography>
                </Box>
              </Box>
            ))}
          </Box>
        )}
      </Paper>

      <Box sx={{ display: 'flex', justifyContent: 'center', gap: 2 }}>
        <Button
          variant="contained"
          size="large"
          startIcon={<PlayArrowIcon />}
          onClick={handleIniciarConsulta}
          disabled={loading || veiculosAtivos.length === 0}
          sx={{ minWidth: 250, py: 1.5 }}
        >
          {loading ? 'Iniciando...' : 'Iniciar Consulta Automática'}
        </Button>
      </Box>

      <Box sx={{ mt: 2, textAlign: 'center' }}>
        <Typography variant="body2" color="text.secondary">
          {veiculosAtivos.length} veículo(s) serão consultado(s) automaticamente
        </Typography>
      </Box>
    </Layout>
  );
}
