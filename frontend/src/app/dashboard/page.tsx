'use client';

import React, { useEffect, useState } from 'react';
import { differenceInCalendarDays, addDays, parse, isAfter, isBefore, startOfDay, endOfDay } from 'date-fns';
import { 
  Grid, Typography, Box, Paper, Button, TextField, Dialog,
  DialogTitle, DialogContent, DialogActions, Table, TableBody,
  TableCell, TableContainer, TableHead, TableRow, IconButton, Chip,
  Tabs, Tab, MenuItem, Select, FormControl, InputLabel
} from '@mui/material';
import DirectionsCarIcon from '@mui/icons-material/DirectionsCar';
import WarningAmberIcon from '@mui/icons-material/WarningAmber';
import AttachMoneyIcon from '@mui/icons-material/AttachMoney';
import PictureAsPdfIcon from '@mui/icons-material/PictureAsPdf';
import AddIcon from '@mui/icons-material/Add';
import DeleteIcon from '@mui/icons-material/Delete';
import PersonIcon from '@mui/icons-material/Person';
import ReceiptIcon from '@mui/icons-material/Receipt';
import ClearIcon from '@mui/icons-material/Clear';
import ErrorIcon from '@mui/icons-material/Error';
import Layout from '@/components/Layout';
import StatusCard from '@/components/StatusCard';
import { Veiculo } from '@/lib/api';

interface Condutor {
  id: string;
  nome: string;
  cpf: string;
}

interface VeiculoComCondutor extends Veiculo {
  condutorId?: string;
}

interface Multa {
  placa: string;
  numero: number;
  ait: string;
  ait_originaria: string;
  motivo: string;
  data_infracao: string;
  data_vencimento: string;
  valor: string;
  valor_a_pagar: string;
  orgao_autuador: string;
  codigo_pagamento: string;
}

export default function Dashboard() {
  const [stats, setStats] = useState({
    veiculosConsultados: 0,
    multasEncontradas: 0,
    valorTotal: 0,
    pdfsGerados: 0,
  });

  const [abaSelecionada, setAbaSelecionada] = useState(0);

  const [condutores, setCondutores] = useState<Condutor[]>([]);
  const [veiculos, setVeiculos] = useState<VeiculoComCondutor[]>([
    { placa: 'SBA7F09', renavam: '01365705622' },
    { placa: 'TIF1J98', renavam: '01450499292' },
  ]);

  const [indicacaoPrazo, setIndicacaoPrazo] = useState<Record<string, { notificacao: Date; prazo: Date }>>({});

  const [openDialogCondutor, setOpenDialogCondutor] = useState(false);
  const [novoCondutorNome, setNovoCondutorNome] = useState('');
  const [novoCondutorCPF, setNovoCondutorCPF] = useState('');

  const [openDialogVeiculo, setOpenDialogVeiculo] = useState(false);
  const [novaPlaca, setNovaPlaca] = useState('');
  const [novoRenavam, setNovoRenavam] = useState('');
  const [condutorSelecionado, setCondutorSelecionado] = useState('');

  const [multas, setMultas] = useState<Multa[]>([]);
  const [filtroTipoMulta, setFiltroTipoMulta] = useState('');
  const [dataInicio, setDataInicio] = useState('');
  const [dataFim, setDataFim] = useState('');
  const [filtroPlaca, setFiltroPlaca] = useState('');
  const [filtroArtigo, setFiltroArtigo] = useState('');

  const artigosCTB = [
    { codigo: 'Art. 161', descricao: 'Infringir qualquer norma do CTB' },
    { codigo: 'Art. 162', descricao: 'Dirigir sem CNH / CNH vencida / categoria diferente' },
    { codigo: 'Art. 163', descricao: 'Entregar veículo a pessoa sem CNH' },
    { codigo: 'Art. 164', descricao: 'Permitir que pessoa sem CNH dirija' },
    { codigo: 'Art. 165', descricao: 'Dirigir sob influência de álcool ou drogas' },
    { codigo: 'Art. 165-A', descricao: 'Recusar teste do bafômetro' },
    { codigo: 'Art. 166', descricao: 'Confiar veículo a condutor alcoolizado' },
    { codigo: 'Art. 167', descricao: 'Deixar de usar cinto de segurança' },
    { codigo: 'Art. 168', descricao: 'Transporte inadequado de criança' },
    { codigo: 'Art. 169', descricao: 'Dirigir sem atenção ou sem cuidados indispensáveis' },
    { codigo: 'Art. 170', descricao: 'Dirigir ameaçando pedestres ou veículos' },
    { codigo: 'Art. 171', descricao: 'Usar veículo para arremessar água ou detritos' },
    { codigo: 'Art. 172', descricao: 'Manobra perigosa (arrancada brusca, derrapagem)' },
    { codigo: 'Art. 173', descricao: 'Disputar corrida (racha)' },
    { codigo: 'Art. 174', descricao: 'Promover ou participar de competição não autorizada' },
    { codigo: 'Art. 175', descricao: 'Exibição de manobra perigosa' },
    { codigo: 'Art. 176', descricao: 'Deixar de prestar socorro em acidente' },
    { codigo: 'Art. 177', descricao: 'Deixar de prestar informações após acidente' },
    { codigo: 'Art. 178', descricao: 'Abandonar veículo em via pública' },
    { codigo: 'Art. 181', descricao: 'Estacionar em local proibido' },
    { codigo: 'Art. 182', descricao: 'Parar em local proibido' },
    { codigo: 'Art. 183', descricao: 'Parar sobre faixa de pedestres' },
    { codigo: 'Art. 184', descricao: 'Transitar em calçada, passeio ou acostamento' },
    { codigo: 'Art. 185', descricao: 'Transitar na contramão' },
    { codigo: 'Art. 186', descricao: 'Transitar em faixa exclusiva' },
    { codigo: 'Art. 187', descricao: 'Transitar em locais proibidos' },
    { codigo: 'Art. 188', descricao: 'Ultrapassar pela direita' },
    { codigo: 'Art. 189', descricao: 'Ultrapassar pela contramão em local proibido' },
    { codigo: 'Art. 190', descricao: 'Forçar passagem entre veículos' },
    { codigo: 'Art. 191', descricao: 'Forçar ultrapassagem perigosa' },
    { codigo: 'Art. 192', descricao: 'Não guardar distância de segurança' },
    { codigo: 'Art. 193', descricao: 'Caminhão fora da faixa adequada' },
    { codigo: 'Art. 194', descricao: 'Marcha à ré perigosa' },
    { codigo: 'Art. 195', descricao: 'Desobedecer ordens da autoridade de trânsito' },
    { codigo: 'Art. 196', descricao: 'Avançar sinal vermelho' },
    { codigo: 'Art. 197', descricao: 'Não sinalizar conversão' },
    { codigo: 'Art. 198', descricao: 'Transpor bloqueio viário/policial' },
    { codigo: 'Art. 199', descricao: 'Ultrapassar pela contramão veículo parado' },
    { codigo: 'Art. 200', descricao: 'Uso indevido da buzina' },
    { codigo: 'Art. 201', descricao: 'Deixar de reduzir velocidade' },
    { codigo: 'Art. 202', descricao: 'Ultrapassar em interseções' },
    { codigo: 'Art. 203', descricao: 'Ultrapassagens proibidas' },
    { codigo: 'Art. 204', descricao: 'Não dar preferência ao pedestre' },
    { codigo: 'Art. 205', descricao: 'Não dar preferência em cruzamentos' },
    { codigo: 'Art. 206', descricao: 'Executar retorno proibido' },
    { codigo: 'Art. 208', descricao: 'Avançar sinal vermelho ou parada obrigatória' },
    { codigo: 'Art. 209', descricao: 'Transpor bloqueio ferroviário' },
    { codigo: 'Art. 210', descricao: 'Transitar com veículo excessivamente lento' },
    { codigo: 'Art. 211', descricao: 'Ultrapassar veículos em cortejo, desfile ou fila' },
    { codigo: 'Art. 212', descricao: 'Não parar para pedestre na faixa' },
    { codigo: 'Art. 213', descricao: 'Deixar de dar passagem a veículo de emergência' },
    { codigo: 'Art. 214', descricao: 'Não dar preferência ao pedestre' },
    { codigo: 'Art. 215', descricao: 'Avançar sinal destinado a pedestre' },
    { codigo: 'Art. 216', descricao: 'Parar sobre faixa de pedestre' },
    { codigo: 'Art. 217', descricao: 'Transitar sobre passeio ou calçada' },
    { codigo: 'Art. 218', descricao: 'Excesso de velocidade' },
    { codigo: 'Art. 219', descricao: 'Transitar em velocidade incompatível' },
    { codigo: 'Art. 220', descricao: 'Não reduzir velocidade próximo a escolas/hospitais' },
    { codigo: 'Art. 230', descricao: 'Conduzir veículo em desacordo com normas' },
    { codigo: 'Art. 231', descricao: 'Transitar com excesso de peso/dimensões' },
    { codigo: 'Art. 232', descricao: 'Conduzir sem documentos obrigatórios' },
    { codigo: 'Art. 233', descricao: 'Deixar de efetuar registro/licenciamento' },
    { codigo: 'Art. 234', descricao: 'Uso de documento falsificado ou adulterado' },
    { codigo: 'Art. 235', descricao: 'Transporte irregular de pessoas' },
    { codigo: 'Art. 236', descricao: 'Excesso de passageiros' },
    { codigo: 'Art. 237', descricao: 'Evadir fiscalização' },
    { codigo: 'Art. 238', descricao: 'Uso indevido de veículo oficial' },
    { codigo: 'Art. 239', descricao: 'Transporte irregular de criança' },
    { codigo: 'Art. 240', descricao: 'Deixar de sinalizar carga' },
    { codigo: 'Art. 241', descricao: 'Uso de equipamento proibido' },
    { codigo: 'Art. 244', descricao: 'Infrações cometidas por motociclistas' },
    { codigo: 'Art. 245', descricao: 'Conduzir motocicleta sem equipamentos obrigatórios' },
    { codigo: 'Art. 247', descricao: 'Deixar de sinalizar obstáculo na via' },
    { codigo: 'Art. 248', descricao: 'Transportar carga em desacordo' },
    { codigo: 'Art. 249', descricao: 'Falta de identificação da carga' },
    { codigo: 'Art. 250', descricao: 'Equipamentos obrigatórios inoperantes' },
    { codigo: 'Art. 251', descricao: 'Veículo com características alteradas' },
    { codigo: 'Art. 252', descricao: 'Dirigir usando celular ou fones' },
    { codigo: 'Art. 253', descricao: 'Bloquear via pública' },
    { codigo: 'Art. 253-A', descricao: 'Interromper via sem autorização' },
    { codigo: 'Art. 255', descricao: 'Infração cometida por pedestre ou ciclista' },
  ];

  useEffect(() => {
    // Carregar veículos do localStorage
    const veiculosSalvos = localStorage.getItem('veiculos');
    if (veiculosSalvos) {
      setVeiculos(JSON.parse(veiculosSalvos));
    }

    // Carregar condutores do localStorage
    const condutoresSalvos = localStorage.getItem('condutores');
    if (condutoresSalvos) {
      setCondutores(JSON.parse(condutoresSalvos));
    }

    // Carregar multas e contar
    const multasSalvasRaw = localStorage.getItem('ultimas_multas');
    let totalMultas = 0;
    let valorTotalMultas = 0;
    let multasCarregadas: Multa[] = [];

    if (multasSalvasRaw) {
      try {
        multasCarregadas = JSON.parse(multasSalvasRaw);
        setMultas(multasCarregadas);
        totalMultas = multasCarregadas.length;

        // Tenta somar os valores das multas
        multasCarregadas.forEach(m => {
          const valor = m.valor_a_pagar || m.valor || '0';
          const valorNumerico = parseFloat(valor.replace(/[^\d,-]/g, '').replace(',', '.'));
          if (!isNaN(valorNumerico)) {
            valorTotalMultas += valorNumerico;
          }
        });
      } catch (e) {
        console.error('Erro ao carregar multas do localStorage', e);
      }
    }

    // Aqui você pode buscar estatísticas reais da API
    setStats({
      veiculosConsultados: veiculosSalvos ? JSON.parse(veiculosSalvos).length : 0,
      multasEncontradas: totalMultas,
      valorTotal: valorTotalMultas,
      pdfsGerados: 0,
    });

    // Função para carregar as notificações
    const carregarNotificacoes = () => {
      const multasSalvasRaw = localStorage.getItem('ultimas_multas');
      if (multasSalvasRaw) {
        try {
          const multasSalvas = JSON.parse(multasSalvasRaw) as Array<{ placa: string; data_infracao?: string }>;
          const map: Record<string, { notificacao: Date; prazo: Date }> = {};

          const parseDataBR = (value: string) => {
            if (!value) return null;
            const parsed = parse(value, 'dd/MM/yyyy', new Date());
            if (!Number.isNaN(parsed.getTime())) return parsed;
            const fallback = new Date(value);
            return Number.isNaN(fallback.getTime()) ? null : fallback;
          };

          multasSalvas.forEach((m) => {
            const data = m.data_infracao ? parseDataBR(m.data_infracao) : null;
            if (!data) return;
            const diff = differenceInCalendarDays(new Date(), data);
            if (diff < 0 || diff > 15) return; // fora do prazo
            const prazo = addDays(data, 15);
            const atual = map[m.placa];
            if (!atual || data > atual.notificacao) {
              map[m.placa] = { notificacao: data, prazo };
            }
          });

          setIndicacaoPrazo(map);
        } catch (e) {
          console.error('Erro ao carregar últimas multas do localStorage', e);
        }
      }
    };

    carregarNotificacoes();

    // Recarrega a cada 5 segundos para pegar atualizações
    const interval = setInterval(carregarNotificacoes, 5000);
    return () => clearInterval(interval);
  }, []);

  const handleAddCondutor = () => {
    if (!novoCondutorNome.trim() || !novoCondutorCPF.trim()) {
      alert('Preencha nome e CPF do condutor');
      return;
    }

    const novoCondutor: Condutor = {
      id: Date.now().toString(),
      nome: novoCondutorNome,
      cpf: novoCondutorCPF
    };

    const condutoresAtualizados = [...condutores, novoCondutor];
    setCondutores(condutoresAtualizados);
    localStorage.setItem('condutores', JSON.stringify(condutoresAtualizados));

    setNovoCondutorNome('');
    setNovoCondutorCPF('');
    setOpenDialogCondutor(false);
  };

  const handleDeleteCondutor = (id: string) => {
    const condutoresAtualizados = condutores.filter(c => c.id !== id);
    setCondutores(condutoresAtualizados);
    localStorage.setItem('condutores', JSON.stringify(condutoresAtualizados));
  };

  const handleAddVeiculo = () => {
    if (!novaPlaca.trim() || !novoRenavam.trim()) {
      alert('Preencha placa e renavam');
      return;
    }

    const novoVeiculo: VeiculoComCondutor = {
      placa: novaPlaca.toUpperCase(),
      renavam: novoRenavam,
      condutorId: condutorSelecionado || undefined
    };

    const veiculosAtualizados = [...veiculos, novoVeiculo];
    setVeiculos(veiculosAtualizados);
    localStorage.setItem('veiculos', JSON.stringify(veiculosAtualizados));

    setNovaPlaca('');
    setNovoRenavam('');
    setCondutorSelecionado('');
    setOpenDialogVeiculo(false);
  };

  const handleDeleteVeiculo = (index: number) => {
    const veiculosAtualizados = veiculos.filter((_, i) => i !== index);
    setVeiculos(veiculosAtualizados);
    localStorage.setItem('veiculos', JSON.stringify(veiculosAtualizados));
  };

  const parseDataBR = (value: string) => {
    if (!value) return null;
    const parsed = parse(value, 'dd/MM/yyyy', new Date());
    if (!Number.isNaN(parsed.getTime())) return parsed;
    const fallback = new Date(value);
    return Number.isNaN(fallback.getTime()) ? null : fallback;
  };

  const multasFiltradas = multas.filter(m => {
    // Filtro por placa
    if (filtroPlaca && m.placa !== filtroPlaca) {
      return false;
    }

    // Filtro por artigo CTB
    if (filtroArtigo && !m.motivo.toLowerCase().includes(filtroArtigo.toLowerCase())) {
      return false;
    }

    // Filtro por tipo de multa (motivo)
    if (filtroTipoMulta && !m.motivo.toLowerCase().includes(filtroTipoMulta.toLowerCase())) {
      return false;
    }

    // Filtro por data de infração
    if (dataInicio || dataFim) {
      const dataInfracao = parseDataBR(m.data_infracao);
      if (!dataInfracao) return false;

      if (dataInicio) {
        const inicio = startOfDay(new Date(dataInicio));
        if (isBefore(dataInfracao, inicio)) return false;
      }

      if (dataFim) {
        const fim = endOfDay(new Date(dataFim));
        if (isAfter(dataInfracao, fim)) return false;
      }
    }

    return true;
  });

  const tiposMultaUnicos = Array.from(new Set(multas.map(m => m.motivo).filter(m => m)));
  const placasUnicas = Array.from(new Set(multas.map(m => m.placa).filter(m => m)));

  // Calcular multas vencendo nos próximos 7 dias
  const multasVencendoEmBreve = multas.filter(m => {
    const dataVencimento = parseDataBR(m.data_vencimento);
    if (!dataVencimento) return false;

    const hoje = new Date();
    const diasAteVencimento = differenceInCalendarDays(dataVencimento, hoje);

    // Mostra multas que vencerão de hoje até 7 dias
    return diasAteVencimento >= 0 && diasAteVencimento <= 7;
  }).sort((a, b) => {
    const dataA = parseDataBR(a.data_vencimento);
    const dataB = parseDataBR(b.data_vencimento);
    if (!dataA || !dataB) return 0;
    return dataA.getTime() - dataB.getTime();
  });

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
          Gerenciamento de veículos e resumo de consultas
        </Typography>
      </Box>

      <Grid container spacing={3} sx={{ mb: 4 }}>
        <Grid item xs={12} sm={6} md={3}>
          <StatusCard
            title="Veículos Ativos"
            value={veiculos.length}
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

      {/* Alerta de Multas Vencendo em Breve */}
      {multasVencendoEmBreve.length > 0 && (
        <Paper elevation={3} sx={{ p: 3, mb: 4, backgroundColor: '#ffebee', borderLeft: '4px solid #d32f2f' }}>
          <Box sx={{ display: 'flex', gap: 2, alignItems: 'flex-start' }}>
            <ErrorIcon sx={{ color: '#d32f2f', fontSize: 32, mt: 0.5 }} />
            <Box sx={{ flex: 1 }}>
              <Typography variant="h6" sx={{ color: '#d32f2f', fontWeight: 600 }}>
                ⚠️ {multasVencendoEmBreve.length} Multa(s) Vencendo nos Próximos 7 Dias!
              </Typography>
              <Typography variant="body2" color="text.secondary" sx={{ mt: 1, mb: 2 }}>
                Verifique os vencimentos abaixo para evitar atrasos:
              </Typography>
              <TableContainer sx={{ backgroundColor: 'white', borderRadius: 1 }}>
                <Table size="small">
                  <TableHead>
                    <TableRow sx={{ backgroundColor: '#f5f5f5' }}>
                      <TableCell sx={{ fontWeight: 600 }}>Placa</TableCell>
                      <TableCell sx={{ fontWeight: 600 }}>AIT</TableCell>
                      <TableCell sx={{ fontWeight: 600 }}>Vencimento</TableCell>
                      <TableCell sx={{ fontWeight: 600 }}>Dias</TableCell>
                      <TableCell sx={{ fontWeight: 600 }}>Valor</TableCell>
                    </TableRow>
                  </TableHead>
                  <TableBody>
                    {multasVencendoEmBreve.map((multa, index) => {
                      const dataVencimento = parseDataBR(multa.data_vencimento);
                      const diasAteVencimento = dataVencimento 
                        ? differenceInCalendarDays(dataVencimento, new Date())
                        : -1;
                      const isVencida = diasAteVencimento < 0;
                      const isUrgente = diasAteVencimento <= 3;

                      return (
                        <TableRow 
                          key={index} 
                          sx={{ 
                            backgroundColor: isVencida ? '#ffcdd2' : isUrgente ? '#fff9c4' : 'inherit',
                            '&:hover': { backgroundColor: '#f5f5f5' }
                          }}
                        >
                          <TableCell sx={{ fontWeight: 600 }}>
                            <Chip label={multa.placa} size="small" color="primary" variant="outlined" />
                          </TableCell>
                          <TableCell sx={{ fontFamily: 'monospace', fontSize: '0.875rem' }}>
                            {multa.ait}
                          </TableCell>
                          <TableCell>
                            {multa.data_vencimento}
                          </TableCell>
                          <TableCell>
                            <Chip
                              label={
                                isVencida 
                                  ? `VENCIDA (${Math.abs(diasAteVencimento)}d atrás)` 
                                  : `${diasAteVencimento} dia${diasAteVencimento !== 1 ? 's' : ''}`
                              }
                              size="small"
                              color={isVencida ? 'error' : isUrgente ? 'warning' : 'default'}
                              variant="filled"
                            />
                          </TableCell>
                          <TableCell sx={{ fontWeight: 600, color: 'error.main' }}>
                            {multa.valor_a_pagar}
                          </TableCell>
                        </TableRow>
                      );
                    })}
                  </TableBody>
                </Table>
              </TableContainer>
            </Box>
          </Box>
        </Paper>
      )}

      {/* Abas para Veículos e Condutores */}
      <Paper elevation={2} sx={{ p: 3, mb: 4 }}>
        <Tabs value={abaSelecionada} onChange={(e, newValue) => setAbaSelecionada(newValue)}>
          <Tab label="Veículos Ativos" icon={<DirectionsCarIcon />} iconPosition="start" />
          <Tab label="Condutores" icon={<PersonIcon />} iconPosition="start" />
          <Tab label="Multas" icon={<ReceiptIcon />} iconPosition="start" />
        </Tabs>

        {/* Aba Veículos */}
        {abaSelecionada === 0 && (
          <Box sx={{ mt: 3 }}>
            <Box sx={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', mb: 2 }}>
              <Typography variant="h6" fontWeight={600}>
                Veículos Ativos para Consulta
              </Typography>
              <Button
                variant="contained"
                startIcon={<AddIcon />}
                onClick={() => setOpenDialogVeiculo(true)}
                size="small"
              >
                Adicionar Veículo
              </Button>
            </Box>

            {veiculos.length === 0 ? (
              <Typography color="text.secondary" sx={{ py: 3, textAlign: 'center' }}>
                Nenhum veículo cadastrado. Clique em "Adicionar Veículo" para começar.
              </Typography>
            ) : (
              <TableContainer>
                <Table>
                  <TableHead>
                    <TableRow sx={{ bgcolor: 'primary.main' }}>
                      <TableCell sx={{ color: 'white', fontWeight: 600 }}>Placa</TableCell>
                      <TableCell sx={{ color: 'white', fontWeight: 600 }}>RENAVAM</TableCell>
                      <TableCell sx={{ color: 'white', fontWeight: 600 }}>Condutor</TableCell>
                      <TableCell sx={{ color: 'white', fontWeight: 600, textAlign: 'center' }}>Status</TableCell>
                      <TableCell sx={{ color: 'white', fontWeight: 600, textAlign: 'center' }}>Ação</TableCell>
                    </TableRow>
                  </TableHead>
                  <TableBody>
                    {veiculos.map((veiculo, index) => {
                      const condutor = veiculo.condutorId ? condutores.find(c => c.id === veiculo.condutorId) : null;
                      return (
                        <TableRow key={index} hover>
                          <TableCell sx={{ fontWeight: 600 }}>{veiculo.placa}</TableCell>
                          <TableCell sx={{ fontFamily: 'monospace' }}>{veiculo.renavam}</TableCell>
                          <TableCell>{condutor ? `${condutor.nome} (${condutor.cpf})` : '-'}</TableCell>
                          <TableCell sx={{ textAlign: 'center' }}>
                            {indicacaoPrazo[veiculo.placa] ? (
                              <Box sx={{ display: 'flex', flexDirection: 'column', gap: 0.5, alignItems: 'center' }}>
                                <Chip label="Prazo de indicação" color="error" size="small" />
                                <Typography variant="caption" color="error">
                                  Notificação: {indicacaoPrazo[veiculo.placa].notificacao.toLocaleDateString('pt-BR')} · Prazo até {indicacaoPrazo[veiculo.placa].prazo.toLocaleDateString('pt-BR')}
                                </Typography>
                              </Box>
                            ) : (
                              <Chip label="Ativo" color="success" size="small" />
                            )}
                          </TableCell>
                          <TableCell sx={{ textAlign: 'center' }}>
                            <IconButton
                              size="small"
                              color="error"
                              onClick={() => handleDeleteVeiculo(index)}
                            >
                              <DeleteIcon />
                            </IconButton>
                          </TableCell>
                        </TableRow>
                      );
                    })}
                  </TableBody>
                </Table>
              </TableContainer>
            )}
          </Box>
        )}

        {/* Aba Condutores */}
        {abaSelecionada === 1 && (
          <Box sx={{ mt: 3 }}>
            <Box sx={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', mb: 2 }}>
              <Typography variant="h6" fontWeight={600}>
                Condutores Cadastrados
              </Typography>
              <Button
                variant="contained"
                startIcon={<AddIcon />}
                onClick={() => setOpenDialogCondutor(true)}
                size="small"
              >
                Adicionar Condutor
              </Button>
            </Box>

            {condutores.length === 0 ? (
              <Typography color="text.secondary" sx={{ py: 3, textAlign: 'center' }}>
                Nenhum condutor cadastrado. Clique em "Adicionar Condutor" para começar.
              </Typography>
            ) : (
              <TableContainer>
                <Table>
                  <TableHead>
                    <TableRow sx={{ bgcolor: 'primary.main' }}>
                      <TableCell sx={{ color: 'white', fontWeight: 600 }}>Nome</TableCell>
                      <TableCell sx={{ color: 'white', fontWeight: 600 }}>CPF</TableCell>
                      <TableCell sx={{ color: 'white', fontWeight: 600, textAlign: 'center' }}>Ação</TableCell>
                    </TableRow>
                  </TableHead>
                  <TableBody>
                    {condutores.map((condutor) => (
                      <TableRow key={condutor.id} hover>
                        <TableCell sx={{ fontWeight: 600 }}>{condutor.nome}</TableCell>
                        <TableCell sx={{ fontFamily: 'monospace' }}>{condutor.cpf}</TableCell>
                        <TableCell sx={{ textAlign: 'center' }}>
                          <IconButton
                            size="small"
                            color="error"
                            onClick={() => handleDeleteCondutor(condutor.id)}
                          >
                            <DeleteIcon />
                          </IconButton>
                        </TableCell>
                      </TableRow>
                    ))}
                  </TableBody>
                </Table>
              </TableContainer>
            )}
          </Box>
        )}

        {/* Aba Multas */}
        {abaSelecionada === 2 && (
          <Box sx={{ mt: 3 }}>
            <Typography variant="h6" fontWeight={600} sx={{ mb: 2 }}>
              Filtrar Multas
            </Typography>

            <Paper elevation={1} sx={{ p: 2, mb: 3 }}>
              <Grid container spacing={2}>
                <Grid item xs={12} sm={6} md={3}>
                  <FormControl fullWidth size="small">
                    <InputLabel>Veículo</InputLabel>
                    <Select
                      value={filtroPlaca}
                      label="Veículo"
                      onChange={(e) => setFiltroPlaca(e.target.value)}
                    >
                      <MenuItem value="">Todos</MenuItem>
                      {placasUnicas.map(placa => (
                        <MenuItem key={placa} value={placa}>
                          {placa}
                        </MenuItem>
                      ))}
                    </Select>
                  </FormControl>
                </Grid>

                <Grid item xs={12} sm={6} md={3}>
                  <FormControl fullWidth size="small">
                    <InputLabel>Artigo CTB</InputLabel>
                    <Select
                      value={filtroArtigo}
                      label="Artigo CTB"
                      onChange={(e) => setFiltroArtigo(e.target.value)}
                    >
                      <MenuItem value="">Todos</MenuItem>
                      {artigosCTB.map(artigo => (
                        <MenuItem key={artigo.codigo} value={artigo.codigo}>
                          {artigo.codigo} - {artigo.descricao}
                        </MenuItem>
                      ))}
                    </Select>
                  </FormControl>
                </Grid>

                <Grid item xs={12} sm={6} md={3}>
                  <TextField
                    label="Data Início"
                    type="date"
                    value={dataInicio}
                    onChange={(e) => setDataInicio(e.target.value)}
                    fullWidth
                    size="small"
                    InputLabelProps={{ shrink: true }}
                  />
                </Grid>

                <Grid item xs={12} sm={6} md={3}>
                  <TextField
                    label="Data Fim"
                    type="date"
                    value={dataFim}
                    onChange={(e) => setDataFim(e.target.value)}
                    fullWidth
                    size="small"
                    InputLabelProps={{ shrink: true }}
                  />
                </Grid>

                <Grid item xs={12}>
                  <Button
                    startIcon={<ClearIcon />}
                    onClick={() => {
                      setFiltroPlaca('');
                      setFiltroArtigo('');
                      setFiltroTipoMulta('');
                      setDataInicio('');
                      setDataFim('');
                    }}
                    size="small"
                  >
                    Limpar Filtros
                  </Button>
                </Grid>
              </Grid>
            </Paper>

            <Typography variant="subtitle2" sx={{ mb: 2 }}>
              {multasFiltradas.length} multa(s) encontrada(s)
            </Typography>

            {multasFiltradas.length === 0 ? (
              <Typography color="text.secondary" sx={{ py: 3, textAlign: 'center' }}>
                Nenhuma multa encontrada com os filtros aplicados.
              </Typography>
            ) : (
              <TableContainer component={Paper} elevation={2}>
                <Table size="small">
                  <TableHead>
                    <TableRow sx={{ bgcolor: 'primary.main' }}>
                      <TableCell sx={{ color: 'white', fontWeight: 600 }}>Placa</TableCell>
                      <TableCell sx={{ color: 'white', fontWeight: 600 }}>AIT</TableCell>
                      <TableCell sx={{ color: 'white', fontWeight: 600 }}>Descrição</TableCell>
                      <TableCell sx={{ color: 'white', fontWeight: 600 }}>Data Infração</TableCell>
                      <TableCell sx={{ color: 'white', fontWeight: 600 }}>Valor</TableCell>
                      <TableCell sx={{ color: 'white', fontWeight: 600 }}>Valor a Pagar</TableCell>
                    </TableRow>
                  </TableHead>
                  <TableBody>
                    {multasFiltradas.map((multa, index) => (
                      <TableRow key={index} hover>
                        <TableCell sx={{ fontWeight: 600 }}>
                          <Chip label={multa.placa} size="small" color="primary" variant="outlined" />
                        </TableCell>
                        <TableCell sx={{ fontFamily: 'monospace', fontSize: '0.875rem' }}>
                          {multa.ait}
                        </TableCell>
                        <TableCell sx={{ maxWidth: 300 }}>
                          <Typography variant="body2">{multa.motivo}</Typography>
                        </TableCell>
                        <TableCell>{multa.data_infracao}</TableCell>
                        <TableCell>{multa.valor}</TableCell>
                        <TableCell sx={{ fontWeight: 600, color: 'error.main' }}>
                          {multa.valor_a_pagar}
                        </TableCell>
                      </TableRow>
                    ))}
                  </TableBody>
                </Table>
              </TableContainer>
            )}
          </Box>
        )}
      </Paper>

      <Box sx={{ textAlign: 'center', py: 4 }}>
        <Typography variant="body2" color="text.secondary">
          Estes veículos serão consultados nas próximas automações
        </Typography>
      </Box>

      {/* Dialog para Adicionar Veículo */}
      <Dialog open={openDialogVeiculo} onClose={() => setOpenDialogVeiculo(false)} maxWidth="sm" fullWidth>
        <DialogTitle sx={{ fontWeight: 600 }}>Adicionar Novo Veículo</DialogTitle>
        <DialogContent sx={{ pt: 2 }}>
          <TextField
            label="Placa (ex: ABC1234)"
            value={novaPlaca}
            onChange={(e) => setNovaPlaca(e.target.value.toUpperCase())}
            fullWidth
            margin="normal"
            placeholder="ABC1234"
            inputProps={{ maxLength: 7 }}
          />
          <TextField
            label="RENAVAM (11 dígitos)"
            value={novoRenavam}
            onChange={(e) => setNovoRenavam(e.target.value)}
            fullWidth
            margin="normal"
            placeholder="12345678901"
            inputProps={{ maxLength: 11, type: 'number' }}
          />
          <FormControl fullWidth margin="normal">
            <InputLabel>Condutor (Opcional)</InputLabel>
            <Select
              value={condutorSelecionado}
              label="Condutor (Opcional)"
              onChange={(e) => setCondutorSelecionado(e.target.value)}
            >
              <MenuItem value="">Nenhum</MenuItem>
              {condutores.map(condutor => (
                <MenuItem key={condutor.id} value={condutor.id}>
                  {condutor.nome} ({condutor.cpf})
                </MenuItem>
              ))}
            </Select>
          </FormControl>
        </DialogContent>
        <DialogActions sx={{ p: 2 }}>
          <Button onClick={() => setOpenDialogVeiculo(false)}>Cancelar</Button>
          <Button 
            onClick={handleAddVeiculo} 
            variant="contained"
          >
            Adicionar
          </Button>
        </DialogActions>
      </Dialog>

      {/* Dialog para Adicionar Condutor */}
      <Dialog open={openDialogCondutor} onClose={() => setOpenDialogCondutor(false)} maxWidth="sm" fullWidth>
        <DialogTitle sx={{ fontWeight: 600 }}>Adicionar Novo Condutor</DialogTitle>
        <DialogContent sx={{ pt: 2 }}>
          <TextField
            label="Nome Completo"
            value={novoCondutorNome}
            onChange={(e) => setNovoCondutorNome(e.target.value)}
            fullWidth
            margin="normal"
            placeholder="João Silva"
          />
          <TextField
            label="CPF (11 dígitos)"
            value={novoCondutorCPF}
            onChange={(e) => setNovoCondutorCPF(e.target.value)}
            fullWidth
            margin="normal"
            placeholder="12345678901"
            inputProps={{ maxLength: 11, type: 'number' }}
          />
        </DialogContent>
        <DialogActions sx={{ p: 2 }}>
          <Button onClick={() => setOpenDialogCondutor(false)}>Cancelar</Button>
          <Button 
            onClick={handleAddCondutor} 
            variant="contained"
          >
            Adicionar
          </Button>
        </DialogActions>
      </Dialog>
    </Layout>
  );
}
