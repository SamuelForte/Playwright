"use client";

import { useEffect, useMemo, useState } from "react";
import { useSearchParams } from "next/navigation";
import {
  Alert,
  Box,
  Button,
  Grid,
  MenuItem,
  Paper,
  Stack,
  TextField,
  Typography,
} from "@mui/material";
import { useMutation, useQuery, useQueryClient } from "@tanstack/react-query";
import Layout from "@/components/Layout";
import {
  Condutor,
  Multa,
  criarCondutor,
  listarCondutores,
  registrarIndicacao,
} from "@/lib/api";

interface CondutorCreate {
  nome: string;
  cpf: string;
}

export default function IndicacaoContent() {
  const queryClient = useQueryClient();
  const searchParams = useSearchParams();
  const [multasRecentes, setMultasRecentes] = useState<Multa[]>([]);
  const [condutorSelecionado, setCondutorSelecionado] = useState("");
  const [multaSelecionada, setMultaSelecionada] = useState("");
  const [placaManual, setPlacaManual] = useState("");
  const [aitManual, setAitManual] = useState("");
  const [mensagem, setMensagem] = useState("");
  const [erro, setErro] = useState("");

  const condutoresQuery = useQuery({
    queryKey: ["condutores"],
    queryFn: listarCondutores,
  });

  const criarCondutorMutation = useMutation({
    mutationFn: criarCondutor,
    onSuccess: () => {
      queryClient.invalidateQueries({ queryKey: ["condutores"] });
    },
  });

  const registrarIndicacaoMutation = useMutation({
    mutationFn: registrarIndicacao,
  });

  useEffect(() => {
    const salvas = localStorage.getItem("ultimas_multas");
    if (!salvas) return;
    try {
      const parsed = JSON.parse(salvas) as Multa[];
      setMultasRecentes(parsed);
      
      // Pre-fill from URL params
      const aitParam = searchParams.get("ait");
      const placaParam = searchParams.get("placa");
      
      if (aitParam) {
        const multaEncontrada = parsed.find(m => m.ait === aitParam);
        if (multaEncontrada) {
          setMultaSelecionada(aitParam);
        } else {
          setAitManual(aitParam);
        }
      }
      
      if (placaParam && !aitParam) {
        setPlacaManual(placaParam);
      }
    } catch (e) {
      console.error("Erro ao carregar ultimas multas", e);
    }
  }, [searchParams]);

  const multaAtual = useMemo(
    () => multasRecentes.find((m) => m.ait === multaSelecionada),
    [multaSelecionada, multasRecentes]
  );

  const handleCriarCondutor = async (event: React.FormEvent<HTMLFormElement>) => {
    event.preventDefault();
    const formData = new FormData(event.currentTarget);
    const nome = String(formData.get("nome") || "").trim();
    const cpf = String(formData.get("cpf") || "").trim();
    if (!nome || !cpf) {
      setErro("Informe nome e CPF do condutor.");
      return;
    }
    setErro("");
    setMensagem("");
    try {
      await criarCondutorMutation.mutateAsync({ nome, cpf });
      setMensagem("Condutor cadastrado com sucesso.");
      event.currentTarget.reset();
    } catch (e: any) {
      const msg = e?.message || e?.error?.message || "Erro ao salvar condutor.";
      setErro(msg);
    }
  };

  const handleRegistrarIndicacao = async () => {
    setErro("");
    setMensagem("");

    const condutorId = condutorSelecionado;
    const ait = multaSelecionada || aitManual.trim();
    const placa = multaAtual?.placa || placaManual.trim().toUpperCase();

    if (!condutorId) {
      setErro("Selecione um condutor.");
      return;
    }
    if (!ait || !placa) {
      setErro("Informe AIT e placa da multa a indicar.");
      return;
    }

    try {
      await registrarIndicacaoMutation.mutateAsync({
        ait,
        placa,
        condutorId,
        data_indicacao: new Date().toISOString(),
      });
      setMensagem("Indicação registrada com sucesso.");
      setAitManual("");
      setPlacaManual("");
    } catch (e: any) {
      const msg = e?.response?.data?.detail || e?.message || "Erro ao registrar indicação";
      setErro(msg);
    }
  };

  const nenhumaMultaLocal = multasRecentes.length === 0;
  const condutores = condutoresQuery.data || [];

  return (
    <Layout>
      <Box sx={{ mb: 4 }}>
        <Typography variant="h4" gutterBottom fontWeight={600}>
          Indicação do Condutor
        </Typography>
        <Typography variant="body1" color="text.secondary">
          Associe um condutor a uma infração (AIT) dentro do prazo de 15 dias.
        </Typography>
      </Box>

      {erro && (
        <Alert severity="error" sx={{ mb: 2 }} onClose={() => setErro("")}>
          {erro}
        </Alert>
      )}
      {mensagem && (
        <Alert severity="success" sx={{ mb: 2 }} onClose={() => setMensagem("")}>
          {mensagem}
        </Alert>
      )}

      <Grid container spacing={3}>
        <Grid item xs={12} md={7}>
          <Paper elevation={2} sx={{ p: 3 }}>
            <Typography variant="h6" fontWeight={600} gutterBottom>
              Registrar Indicação
            </Typography>

            {nenhumaMultaLocal && (
              <Alert severity="info" sx={{ mb: 2 }}>
                Nenhuma multa recente encontrada. Você pode preencher AIT e placa manualmente.
              </Alert>
            )}

            <Stack spacing={2}>
              {!nenhumaMultaLocal && (
                <TextField
                  select
                  label="Selecionar multa"
                  value={multaSelecionada}
                  onChange={(e) => setMultaSelecionada(e.target.value)}
                  helperText="Usa as multas mais recentes do resultado de consulta"
                  fullWidth
                >
                  {multasRecentes.map((multa, idx) => (
                    <MenuItem key={`${multa.placa}-${multa.ait}-${idx}`} value={multa.ait}>
                      {multa.placa} · AIT {multa.ait} · {multa.motivo}
                    </MenuItem>
                  ))}
                </TextField>
              )}

              <Box sx={{ display: "flex", gap: 2, flexWrap: "wrap" }}>
                <TextField
                  label="Placa"
                  value={placaManual}
                  onChange={(e) => setPlacaManual(e.target.value.toUpperCase())}
                  fullWidth
                  placeholder={multaAtual?.placa || "ABC1D23"}
                  helperText="Preencha se não escolheu uma multa acima"
                />
                <TextField
                  label="AIT"
                  value={aitManual}
                  onChange={(e) => setAitManual(e.target.value)}
                  fullWidth
                  placeholder={multaAtual?.ait || "000000000"}
                  helperText="Preencha se não escolheu uma multa acima"
                />
              </Box>

              <TextField
                select
                label="Condutor"
                value={condutorSelecionado}
                onChange={(e) => setCondutorSelecionado(e.target.value)}
                fullWidth
                helperText="Cadastre o condutor ao lado se ainda não existir"
              >
                {condutores.map((condutor: Condutor) => (
                  <MenuItem key={condutor.id} value={condutor.id}>
                    {condutor.nome} · CPF {condutor.cpf}
                  </MenuItem>
                ))}
                {condutores.length === 0 && (
                  <MenuItem disabled>Nenhum condutor cadastrado</MenuItem>
                )}
              </TextField>

              <Button
                variant="contained"
                onClick={handleRegistrarIndicacao}
                disabled={registrarIndicacaoMutation.isPending}
              >
                {registrarIndicacaoMutation.isPending ? "Registrando..." : "Registrar Indicação"}
              </Button>
            </Stack>
          </Paper>
        </Grid>

        <Grid item xs={12} md={5}>
          <Paper elevation={2} sx={{ p: 3, mb: 3 }}>
            <Typography variant="h6" fontWeight={600} gutterBottom>
              Condutores
            </Typography>
            <form onSubmit={handleCriarCondutor}>
              <Stack spacing={2}>
                <TextField name="nome" label="Nome" fullWidth required />
                <TextField name="cpf" label="CPF" fullWidth required />
                <Button
                  type="submit"
                  variant="outlined"
                  disabled={criarCondutorMutation.isPending}
                >
                  {criarCondutorMutation.isPending ? "Salvando..." : "Adicionar Condutor"}
                </Button>
              </Stack>
            </form>

            {condutores.length > 0 && (
              <Box sx={{ mt: 3 }}>
                <Typography variant="subtitle2" gutterBottom fontWeight={600}>
                  Condutores cadastrados
                </Typography>
                <Stack spacing={1}>
                  {condutores.map((condutor: Condutor) => (
                    <Paper key={condutor.id} variant="outlined" sx={{ p: 1.5 }}>
                      <Typography fontWeight={600}>{condutor.nome}</Typography>
                      <Typography variant="body2" color="text.secondary">
                        CPF {condutor.cpf}
                      </Typography>
                    </Paper>
                  ))}
                </Stack>
              </Box>
            )}
          </Paper>

          <Paper elevation={2} sx={{ p: 3 }}>
            <Typography variant="h6" fontWeight={600} gutterBottom>
              Multas Recentes
            </Typography>
            {multasRecentes.length === 0 ? (
              <Typography color="text.secondary">
                Nenhuma multa encontrada no histórico recente. Realize uma consulta para popular esta lista.
              </Typography>
            ) : (
              <Stack spacing={1}>
                {multasRecentes.slice(0, 5).map((multa, idx) => (
                  <Paper key={`${multa.placa}-${multa.ait}-${idx}`} variant="outlined" sx={{ p: 1.5 }}>
                    <Typography fontWeight={600}>{multa.placa}</Typography>
                    <Typography variant="body2" sx={{ fontFamily: "monospace" }}>
                      AIT {multa.ait}
                    </Typography>
                    <Typography variant="body2" color="text.secondary">
                      {multa.motivo}
                    </Typography>
                  </Paper>
                ))}
              </Stack>
            )}
          </Paper>
        </Grid>
      </Grid>
    </Layout>
  );
}
