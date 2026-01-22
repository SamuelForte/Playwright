'use client';

import React from 'react';
import { useParams } from 'next/navigation';
import { addDays, differenceInCalendarDays, parse } from 'date-fns';
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

const parseDataBR = (value: string) => {
  if (!value) return null;
  const parsed = parse(value, 'dd/MM/yyyy', new Date());
  if (!Number.isNaN(parsed.getTime())) return parsed;
  const fallback = new Date(value);
  return Number.isNaN(fallback.getTime()) ? null : fallback;
};

const dentroPrazoIndicacao = (dataNotificacao: string) => {
  const data = parseDataBR(dataNotificacao);
  if (!data) return false;
  const hoje = new Date();
  const diff = differenceInCalendarDays(hoje, data);
  return diff >= 0 && diff <= 15;
};

const prazoIndicacao = (dataNotificacao: string) => {
  const data = parseDataBR(dataNotificacao);
  if (!data) return null;
  return addDays(data, 15);
};

export default function Resultados() {
  const params = useParams();
  const consultaId = params.id as string;

  const { data, isLoading, isError, error } = useQuery({
    queryKey: ['resultado', consultaId],
    queryFn: () => obterResultado(consultaId),
    enabled: !!consultaId,
  });

  // Salva multas no localStorage para o dashboard poder exibir notificações
  React.useEffect(() => {
    if (data?.multas && data.multas.length > 0) {
      localStorage.setItem('ultimas_multas', JSON.stringify(data.multas));
    }
  }, [data?.multas]);

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
  const placasAptas = Array.from(
    new Set(multas.filter(m => dentroPrazoIndicacao(m.data_infracao)).map(m => m.placa))
  );
  const prazosPorPlaca = (() => {
    const map = new Map<string, { notificacao: Date; prazo: Date }>();
    multas.forEach(m => {
      const notificacao = parseDataBR(m.data_infracao);
      if (!notificacao) return;
      const prazo = addDays(notificacao, 15);
      const atual = map.get(m.placa);
      if (!atual || notificacao > atual.notificacao) {
        map.set(m.placa, { notificacao, prazo });
      }
    });
    return Array.from(map.entries()).map(([placa, valores]) => ({
      placa,
      notificacao: valores.notificacao,
      prazo: valores.prazo,
    }));
  })();

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

      {placasAptas.length > 0 && (
        <Alert severity="success" sx={{ mb: 3 }}>
          <Typography fontWeight={600} sx={{ mb: prazosPorPlaca.length ? 1 : 0 }}>
            {placasAptas.length === 1
              ? `O veículo ${placasAptas[0]} está apto a indicar condutor (até 15 dias da notificação).`
              : `Os veículos ${placasAptas.join(', ')} estão aptos a indicar condutor (até 15 dias da notificação).`}
          </Typography>
          {prazosPorPlaca.length > 0 && (
            <Box component="ul" sx={{ pl: 2, mb: 0, mt: 0.5 }}>
              {prazosPorPlaca.map(item => (
                <Box component="li" key={item.placa} sx={{ lineHeight: 1.4 }}>
                  <Typography variant="body2">
                    {item.placa}: notificação em {item.notificacao.toLocaleDateString('pt-BR')} • prazo até {item.prazo.toLocaleDateString('pt-BR')}
                  </Typography>
                </Box>
              ))}
            </Box>
          )}
        </Alert>
      )}

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
