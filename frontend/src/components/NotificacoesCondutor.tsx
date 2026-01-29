'use client';

import React, { useMemo } from 'react';
import { useRouter } from 'next/navigation';
import {
  Alert,
  Box,
  Button,
  Paper,
  Typography,
  Table,
  TableBody,
  TableCell,
  TableContainer,
  TableHead,
  TableRow,
  Chip,
} from '@mui/material';
import NotificationsActiveIcon from '@mui/icons-material/NotificationsActive';
import HowToRegIcon from '@mui/icons-material/HowToReg';

interface Condutor {
  id: string;
  nome: string;
  cpf: string;
}

interface Veiculo {
  placa: string;
  renavam: string;
  condutorId?: string;
}

interface Multa {
  placa: string;
  numero: number;
  ait: string;
  motivo: string;
  data_infracao: string;
  data_vencimento: string;
  valor_a_pagar: string;
}

interface NotificacoesCondutorProps {
  condutores: Condutor[];
  veiculos: Veiculo[];
  multas: Multa[];
}

export default function NotificacoesCondutor({
  condutores,
  veiculos,
  multas,
}: NotificacoesCondutorProps) {
  const router = useRouter();

  // Agrupar notificações por condutor
  const notificacoesPorCondutor = useMemo(() => {
    const map = new Map<string, { condutor: Condutor; multas: Multa[] }>();

    veiculos.forEach((veiculo) => {
      if (!veiculo.condutorId) return;

      const condutor = condutores.find((c) => c.id === veiculo.condutorId);
      if (!condutor) return;

      const multasDoVeiculo = multas.filter((m) => m.placa === veiculo.placa);

      if (multasDoVeiculo.length === 0) return;

      if (!map.has(condutor.id)) {
        map.set(condutor.id, { condutor, multas: [] });
      }

      const entry = map.get(condutor.id)!;
      entry.multas.push(...multasDoVeiculo);
    });

    return Array.from(map.values());
  }, [condutores, veiculos, multas]);

  if (notificacoesPorCondutor.length === 0) {
    return null;
  }

  const totalNotificacoes = notificacoesPorCondutor.reduce(
    (acc, n) => acc + n.multas.length,
    0
  );

  return (
    <Paper
      elevation={3}
      sx={{
        p: 3,
        mb: 4,
        backgroundColor: '#e3f2fd',
        borderLeft: '4px solid #1976d2',
      }}
    >
      <Box sx={{ display: 'flex', gap: 2, alignItems: 'flex-start' }}>
        <NotificationsActiveIcon
          sx={{ color: '#1976d2', fontSize: 32, mt: 0.5 }}
        />
        <Box sx={{ flex: 1 }}>
          <Typography variant="h6" sx={{ color: '#1976d2', fontWeight: 600 }}>
            🔔 {totalNotificacoes} Notificação(ões) de Multa
          </Typography>
          <Typography variant="body2" color="text.secondary" sx={{ mt: 1, mb: 2 }}>
            Condutores associados aos veículos foram notificados sobre novas multas.
            Eles podem indicar o condutor real dentro do prazo de 15 dias da notificação.
          </Typography>

          {notificacoesPorCondutor.map(({ condutor, multas: multasCondutor }) => (
            <Paper
              key={condutor.id}
              elevation={1}
              sx={{ p: 2, mb: 2, backgroundColor: 'white' }}
            >
              <Box
                sx={{
                  display: 'flex',
                  justifyContent: 'space-between',
                  alignItems: 'center',
                  mb: 2,
                }}
              >
                <Box>
                  <Typography variant="subtitle1" fontWeight={600}>
                    Condutor: {condutor.nome}
                  </Typography>
                  <Typography variant="body2" color="text.secondary">
                    CPF: {condutor.cpf} · {multasCondutor.length} multa(s) registrada(s)
                  </Typography>
                </Box>
                <Button
                  variant="contained"
                  size="small"
                  startIcon={<HowToRegIcon />}
                  onClick={() => {
                    // Pre-fill com primeira multa do condutor
                    const primeiraMulta = multasCondutor[0];
                    if (primeiraMulta) {
                      router.push(`/indicacao?ait=${primeiraMulta.ait}&placa=${primeiraMulta.placa}`);
                    } else {
                      router.push('/indicacao');
                    }
                  }}
                >
                  Indicar Condutor
                </Button>
              </Box>

              <TableContainer sx={{ maxHeight: 300 }}>
                <Table size="small">
                  <TableHead>
                    <TableRow sx={{ backgroundColor: '#f5f5f5' }}>
                      <TableCell sx={{ fontWeight: 600 }}>Placa</TableCell>
                      <TableCell sx={{ fontWeight: 600 }}>AIT</TableCell>
                      <TableCell sx={{ fontWeight: 600 }}>Descrição</TableCell>
                      <TableCell sx={{ fontWeight: 600 }}>Data Infração</TableCell>
                      <TableCell sx={{ fontWeight: 600 }}>Valor</TableCell>
                      <TableCell sx={{ fontWeight: 600, textAlign: 'center' }}>Ação</TableCell>
                    </TableRow>
                  </TableHead>
                  <TableBody>
                    {multasCondutor.map((multa, index) => (
                      <TableRow key={index} hover>
                        <TableCell>
                          <Chip
                            label={multa.placa}
                            size="small"
                            color="primary"
                            variant="outlined"
                          />
                        </TableCell>
                        <TableCell sx={{ fontFamily: 'monospace', fontSize: '0.875rem' }}>
                          {multa.ait}
                        </TableCell>
                        <TableCell sx={{ maxWidth: 300 }}>
                          <Typography variant="body2" noWrap>
                            {multa.motivo}
                          </Typography>
                        </TableCell>
                        <TableCell>{multa.data_infracao}</TableCell>
                        <TableCell sx={{ fontWeight: 600, color: 'error.main' }}>
                          {multa.valor_a_pagar}
                        </TableCell>
                        <TableCell sx={{ textAlign: 'center' }}>
                          <Button
                            size="small"
                            variant="outlined"
                            startIcon={<HowToRegIcon />}
                            onClick={() => router.push(`/indicacao?ait=${multa.ait}&placa=${multa.placa}`)}
                          >
                            Indicar
                          </Button>
                        </TableCell>
                      </TableRow>
                    ))}
                  </TableBody>
                </Table>
              </TableContainer>
            </Paper>
          ))}

          <Alert severity="info" sx={{ mt: 2 }}>
            <Typography variant="body2" fontWeight={600}>
              ⏰ Atenção ao prazo!
            </Typography>
            <Typography variant="body2">
              A indicação do condutor real deve ser feita em até 15 dias da data da
              notificação da infração. Após esse prazo, a multa e os pontos serão
              aplicados ao proprietário do veículo.
            </Typography>
          </Alert>
        </Box>
      </Box>
    </Paper>
  );
}
