'use client';

import React, { useEffect, useState } from 'react';
import {
  Box,
  Typography,
  Paper,
  TextField,
  Button,
  Card,
  CardContent,
  Grid,
  Alert,
  CircularProgress,
  Chip,
  Dialog,
  DialogTitle,
  DialogContent,
  DialogActions,
} from '@mui/material';
import { useQuery, useMutation } from '@tanstack/react-query';
import Layout from '@/components/Layout';
import { supabase } from '@/lib/supabaseClient';

interface Indicacao {
  id: string;
  ait: string;
  placa: string;
  condutor_id: string;
  data_indicacao: string;
  status: string;
  defesa_texto?: string;
  comprovante_url?: string;
  resposta_status: string;
  data_resposta?: string;
  observacoes?: string;
}

interface Condutor {
  id: string;
  nome: string;
  cpf: string;
}

export default function CondutorIndicacaoPage() {
  const [condutorSelecionado, setCondutorSelecionado] = useState<Condutor | null>(null);
  const [indicacoesDoConductor, setIndicacoesDoConductor] = useState<Indicacao[]>([]);
  const [indicacaoSelecionada, setIndicacaoSelecionada] = useState<Indicacao | null>(null);
  const [openFormDialog, setOpenFormDialog] = useState(false);
  const [carregando, setCarregando] = useState(false);
  const [mensagem, setMensagem] = useState('');
  const [erro, setErro] = useState('');

  // Formulário
  const [defesaTexto, setDefesaTexto] = useState('');
  const [comprovanteUrl, setComprovanteUrl] = useState('');
  const [observacoes, setObservacoes] = useState('');

  // Buscar condutores
  const condutoresQuery = useQuery({
    queryKey: ['condutores'],
    queryFn: async () => {
      const { data, error } = await supabase
        .from('condutores')
        .select('*')
        .order('created_at', { ascending: false });
      if (error) throw error;
      return data as Condutor[];
    },
  });

  // Buscar indicações de um condutor
  const buscarIndicacoes = async (condutorId: string) => {
    try {
      setCarregando(true);
      const { data, error } = await supabase
        .from('indicacoes')
        .select('*')
        .eq('condutor_id', condutorId)
        .order('data_indicacao', { ascending: false });

      if (error) throw error;
      setIndicacoesDoConductor(data || []);
      setMensagem('');
    } catch (e: any) {
      setErro(e?.message || 'Erro ao carregar indicações');
    } finally {
      setCarregando(false);
    }
  };

  // Atualizar indicação
  const atualizarIndicacao = async () => {
    if (!indicacaoSelecionada) return;

    try {
      setCarregando(true);
      const { error } = await supabase
        .from('indicacoes')
        .update({
          defesa_texto: defesaTexto.trim(),
          comprovante_url: comprovanteUrl.trim(),
          observacoes: observacoes.trim(),
          resposta_status: 'respondido',
          data_resposta: new Date().toISOString(),
        })
        .eq('id', indicacaoSelecionada.id);

      if (error) throw error;

      setMensagem('Indicação preenchida com sucesso!');
      setOpenFormDialog(false);
      buscarIndicacoes(indicacaoSelecionada.condutor_id);
      resetarFormulario();
    } catch (e: any) {
      setErro(e?.message || 'Erro ao salvar indicação');
    } finally {
      setCarregando(false);
    }
  };

  const resetarFormulario = () => {
    setDefesaTexto('');
    setComprovanteUrl('');
    setObservacoes('');
  };

  const handleSelecionarCondutor = (condutor: Condutor) => {
    setCondutorSelecionado(condutor);
    buscarIndicacoes(condutor.id);
  };

  const handleAbrirFormulario = (indicacao: Indicacao) => {
    setIndicacaoSelecionada(indicacao);
    setDefesaTexto(indicacao.defesa_texto || '');
    setComprovanteUrl(indicacao.comprovante_url || '');
    setObservacoes(indicacao.observacoes || '');
    setOpenFormDialog(true);
  };

  const condutores = condutoresQuery.data || [];

  return (
    <Layout>
      <Box sx={{ mb: 4 }}>
        <Typography variant="h4" gutterBottom fontWeight={600}>
          Responder Indicação
        </Typography>
        <Typography variant="body1" color="text.secondary">
          Selecione um condutor para visualizar suas indicações e preencher o formulário de resposta.
        </Typography>
      </Box>

      {erro && (
        <Alert severity="error" sx={{ mb: 2 }} onClose={() => setErro('')}>
          {erro}
        </Alert>
      )}
      {mensagem && (
        <Alert severity="success" sx={{ mb: 2 }} onClose={() => setMensagem('')}>
          {mensagem}
        </Alert>
      )}

      <Grid container spacing={3}>
        {/* Coluna de Condutores */}
        <Grid item xs={12} md={4}>
          <Paper elevation={2} sx={{ p: 3 }}>
            <Typography variant="h6" fontWeight={600} gutterBottom>
              Condutores
            </Typography>

            {condutoresQuery.isLoading ? (
              <Box sx={{ display: 'flex', justifyContent: 'center', py: 3 }}>
                <CircularProgress />
              </Box>
            ) : condutores.length === 0 ? (
              <Typography color="text.secondary" variant="body2">
                Nenhum condutor cadastrado.
              </Typography>
            ) : (
              <Box sx={{ display: 'flex', flexDirection: 'column', gap: 1 }}>
                {condutores.map((condutor) => (
                  <Card
                    key={condutor.id}
                    sx={{
                      cursor: 'pointer',
                      bgcolor:
                        condutorSelecionado?.id === condutor.id
                          ? 'primary.light'
                          : 'inherit',
                      border:
                        condutorSelecionado?.id === condutor.id
                          ? '2px solid'
                          : '1px solid',
                      borderColor: 'divider',
                      transition: 'all 0.2s',
                      '&:hover': { boxShadow: 2 },
                    }}
                    onClick={() => handleSelecionarCondutor(condutor)}
                  >
                    <CardContent sx={{ py: 1.5 }}>
                      <Typography fontWeight={600} variant="body2">
                        {condutor.nome}
                      </Typography>
                      <Typography variant="caption" color="text.secondary">
                        CPF: {condutor.cpf}
                      </Typography>
                    </CardContent>
                  </Card>
                ))}
              </Box>
            )}
          </Paper>
        </Grid>

        {/* Coluna de Indicações */}
        <Grid item xs={12} md={8}>
          {!condutorSelecionado ? (
            <Paper elevation={2} sx={{ p: 3, textAlign: 'center' }}>
              <Typography color="text.secondary">
                Selecione um condutor à esquerda para visualizar suas indicações.
              </Typography>
            </Paper>
          ) : (
            <Box>
              <Paper elevation={2} sx={{ p: 3, mb: 2 }}>
                <Box sx={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', mb: 2 }}>
                  <Box>
                    <Typography variant="h6" fontWeight={600}>
                      Indicações de {condutorSelecionado.nome}
                    </Typography>
                    <Typography variant="body2" color="text.secondary">
                      CPF: {condutorSelecionado.cpf}
                    </Typography>
                  </Box>
                </Box>

                {carregando ? (
                  <Box sx={{ display: 'flex', justifyContent: 'center', py: 3 }}>
                    <CircularProgress />
                  </Box>
                ) : indicacoesDoConductor.length === 0 ? (
                  <Typography color="text.secondary" variant="body2" sx={{ py: 2 }}>
                    Nenhuma indicação encontrada para este condutor.
                  </Typography>
                ) : (
                  <Box sx={{ display: 'flex', flexDirection: 'column', gap: 2 }}>
                    {indicacoesDoConductor.map((indicacao) => (
                      <Card key={indicacao.id} variant="outlined">
                        <CardContent>
                          <Box sx={{ display: 'flex', justifyContent: 'space-between', alignItems: 'start', mb: 2 }}>
                            <Box>
                              <Typography fontWeight={600}>
                                AIT {indicacao.ait}
                              </Typography>
                              <Typography variant="body2" color="text.secondary">
                                Placa: {indicacao.placa}
                              </Typography>
                            </Box>
                            <Chip
                              label={indicacao.resposta_status === 'respondido' ? 'Respondido' : 'Pendente'}
                              color={indicacao.resposta_status === 'respondido' ? 'success' : 'warning'}
                              size="small"
                            />
                          </Box>

                          <Typography variant="caption" color="text.secondary">
                            Data da indicação: {new Date(indicacao.data_indicacao).toLocaleDateString('pt-BR')}
                          </Typography>

                          {indicacao.resposta_status === 'respondido' && (
                            <>
                              <Box sx={{ mt: 2, pt: 2, borderTop: '1px solid #eee' }}>
                                <Typography variant="subtitle2" fontWeight={600}>
                                  Defesa fornecida:
                                </Typography>
                                <Typography variant="body2" sx={{ mt: 1, whiteSpace: 'pre-wrap' }}>
                                  {indicacao.defesa_texto}
                                </Typography>
                              </Box>
                              {indicacao.comprovante_url && (
                                <Box sx={{ mt: 2 }}>
                                  <Typography variant="subtitle2" fontWeight={600}>
                                    Comprovante: <a href={indicacao.comprovante_url} target="_blank" rel="noopener noreferrer">Visualizar</a>
                                  </Typography>
                                </Box>
                              )}
                            </>
                          )}

                          <Box sx={{ mt: 2, display: 'flex', gap: 1 }}>
                            <Button
                              size="small"
                              variant="contained"
                              onClick={() => handleAbrirFormulario(indicacao)}
                            >
                              {indicacao.resposta_status === 'respondido' ? 'Editar' : 'Responder'}
                            </Button>
                          </Box>
                        </CardContent>
                      </Card>
                    ))}
                  </Box>
                )}
              </Paper>
            </Box>
          )}
        </Grid>
      </Grid>

      {/* Dialog do Formulário */}
      <Dialog open={openFormDialog} onClose={() => setOpenFormDialog(false)} maxWidth="sm" fullWidth>
        <DialogTitle sx={{ fontWeight: 600 }}>
          Preencher Resposta - AIT {indicacaoSelecionada?.ait}
        </DialogTitle>
        <DialogContent sx={{ pt: 2 }}>
          <TextField
            label="Defesa / Resposta"
            value={defesaTexto}
            onChange={(e) => setDefesaTexto(e.target.value)}
            fullWidth
            multiline
            rows={5}
            margin="normal"
            placeholder="Descreva sua defesa ou justificativa..."
            helperText="Campo obrigatório"
          />
          <TextField
            label="URL do Comprovante (opcional)"
            value={comprovanteUrl}
            onChange={(e) => setComprovanteUrl(e.target.value)}
            fullWidth
            margin="normal"
            placeholder="https://..."
            helperText="Link para documento ou imagem que comprove sua defesa"
          />
          <TextField
            label="Observações (opcional)"
            value={observacoes}
            onChange={(e) => setObservacoes(e.target.value)}
            fullWidth
            multiline
            rows={3}
            margin="normal"
            placeholder="Informações adicionais..."
          />
        </DialogContent>
        <DialogActions sx={{ p: 2 }}>
          <Button onClick={() => setOpenFormDialog(false)}>Cancelar</Button>
          <Button
            onClick={atualizarIndicacao}
            variant="contained"
            disabled={carregando || !defesaTexto.trim()}
          >
            {carregando ? 'Salvando...' : 'Salvar Resposta'}
          </Button>
        </DialogActions>
      </Dialog>
    </Layout>
  );
}
