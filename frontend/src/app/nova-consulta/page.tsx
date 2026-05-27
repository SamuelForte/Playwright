'use client';

import React, { useEffect, useState } from 'react';
import { Box, Button, Typography, Paper, Alert } from '@mui/material';
import { useRouter } from 'next/navigation';
import PlayArrowIcon from '@mui/icons-material/PlayArrow';
import Layout from '@/components/Layout';
import { Veiculo, iniciarConsulta } from '@/lib/api';

const withTimeout = async <T,>(promise: Promise<T>, timeoutMs = 15000): Promise<T> => {
  return await Promise.race([
    promise,
    new Promise<T>((_, reject) =>
      setTimeout(() => reject(new Error('Tempo limite ao iniciar consulta.')), timeoutMs)
    ),
  ]);
};

export default function NovaConsulta() {
  const router = useRouter();
  const isLocalDev =
    typeof window !== 'undefined' &&
    (window.location.hostname === 'localhost' || window.location.hostname === '127.0.0.1');
  const [loading, setLoading] = useState(false);
  const [erro, setErro] = useState('');
  const [veiculosAtivos, setVeiculosAtivos] = useState<Veiculo[]>([]);

  // Carregar veículos do localStorage quando o componente monta
  useEffect(() => {
    const veiculosSalvos = localStorage.getItem('veiculos');
    if (veiculosSalvos) {
      try {
        const parsed = JSON.parse(veiculosSalvos);
        if (Array.isArray(parsed)) {
          const normalizados: Veiculo[] = parsed
            .map((item: any) => ({
              placa: String(item?.placa || '').trim().toUpperCase(),
              renavam: String(item?.renavam || '').trim(),
            }))
            .filter((item: Veiculo) => item.placa.length > 0 && item.renavam.length > 0);

          setVeiculosAtivos(normalizados);
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
      const { consulta_id } = await withTimeout(iniciarConsulta(veiculosAtivos), 15000);
      const consultaId = consulta_id;

      console.log('Consulta iniciada com sucesso. ID:', consultaId);
      router.push(`/processamento?id=${consultaId}`);
    } catch (error: any) {
      console.error('Erro ao iniciar consulta:', error);
      const rawMessage = String(error?.message || '');
      const isNetworkError =
        rawMessage.toLowerCase().includes('failed to fetch')
        || rawMessage.toLowerCase().includes('network error');

      const mensagemErro = error.response?.data?.detail 
        || (isNetworkError
          ? (isLocalDev
            ? 'Sem conexão com o backend local. Confirme se a API está em http://localhost:8000.'
            : 'Sem conexão com o backend público. Configure NEXT_PUBLIC_API_URL com a URL correta e valide CORS no backend.')
          : error.message)
        || (isLocalDev
          ? 'Erro ao iniciar consulta. Verifique backend em http://localhost:8000 e frontend em http://localhost:3000'
          : 'Erro ao iniciar consulta. Verifique NEXT_PUBLIC_API_URL e CORS do backend para o domínio publicado.');
      setErro(mensagemErro);
    } finally {
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
