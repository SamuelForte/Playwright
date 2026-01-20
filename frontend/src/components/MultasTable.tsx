'use client';

import React, { useState } from 'react';
import {
  Box,
  Paper,
  Table,
  TableBody,
  TableCell,
  TableContainer,
  TableHead,
  TableRow,
  TablePagination,
  TextField,
  IconButton,
  Chip,
  Typography,
  InputAdornment,
  MenuItem,
  Select,
  FormControl,
  InputLabel,
} from '@mui/material';
import SearchIcon from '@mui/icons-material/Search';
import PictureAsPdfIcon from '@mui/icons-material/PictureAsPdf';
import { Multa } from '@/lib/api';

interface MultasTableProps {
  multas: Multa[];
  onDownloadPDF?: (ait: string) => void;
}

export default function MultasTable({ multas, onDownloadPDF }: MultasTableProps) {
  const [page, setPage] = useState(0);
  const [rowsPerPage, setRowsPerPage] = useState(10);
  const [busca, setBusca] = useState('');
  const [filtroOrgao, setFiltroOrgao] = useState('todos');

  // Extrai órgãos únicos
  const orgaosUnicos = Array.from(new Set(multas.map(m => m.orgao_autuador)));

  // Filtra multas
  const multasFiltradas = multas.filter(multa => {
    const matchBusca = 
      (multa.placa || '').toLowerCase().includes(busca.toLowerCase()) ||
      (multa.ait || '').toLowerCase().includes(busca.toLowerCase()) ||
      (multa.motivo || '').toLowerCase().includes(busca.toLowerCase());
    
    const matchOrgao = filtroOrgao === 'todos' || multa.orgao_autuador === filtroOrgao;

    return matchBusca && matchOrgao;
  });

  // Paginação
  const multasPaginadas = multasFiltradas.slice(
    page * rowsPerPage,
    page * rowsPerPage + rowsPerPage
  );

  const handleChangePage = (_event: unknown, newPage: number) => {
    setPage(newPage);
  };

  const handleChangeRowsPerPage = (event: React.ChangeEvent<HTMLInputElement>) => {
    setRowsPerPage(parseInt(event.target.value, 10));
    setPage(0);
  };

  return (
    <Box>
      <Paper elevation={2} sx={{ p: 2, mb: 2 }}>
        <Box sx={{ display: 'flex', gap: 2, flexWrap: 'wrap' }}>
          <TextField
            placeholder="Buscar por placa, AIT ou descrição..."
            value={busca}
            onChange={(e) => setBusca(e.target.value)}
            InputProps={{
              startAdornment: (
                <InputAdornment position="start">
                  <SearchIcon />
                </InputAdornment>
              ),
            }}
            sx={{ flex: 1, minWidth: 250 }}
          />
          <FormControl sx={{ minWidth: 200 }}>
            <InputLabel>Órgão Autuador</InputLabel>
            <Select
              value={filtroOrgao}
              label="Órgão Autuador"
              onChange={(e) => setFiltroOrgao(e.target.value)}
            >
              <MenuItem value="todos">Todos</MenuItem>
              {orgaosUnicos.map(orgao => (
                <MenuItem key={orgao} value={orgao}>
                  {orgao}
                </MenuItem>
              ))}
            </Select>
          </FormControl>
        </Box>
      </Paper>

      <TableContainer component={Paper} elevation={2}>
        <Table>
          <TableHead>
            <TableRow sx={{ bgcolor: 'primary.main' }}>
              <TableCell sx={{ color: 'white', fontWeight: 600 }}>Placa</TableCell>
              <TableCell sx={{ color: 'white', fontWeight: 600 }}>#</TableCell>
              <TableCell sx={{ color: 'white', fontWeight: 600 }}>AIT</TableCell>
              <TableCell sx={{ color: 'white', fontWeight: 600 }}>Órgão</TableCell>
              <TableCell sx={{ color: 'white', fontWeight: 600, minWidth: 300 }}>Descrição</TableCell>
              <TableCell sx={{ color: 'white', fontWeight: 600 }}>Data Infração</TableCell>
              <TableCell sx={{ color: 'white', fontWeight: 600 }}>Vencimento</TableCell>
              <TableCell sx={{ color: 'white', fontWeight: 600 }}>Valor</TableCell>
              <TableCell sx={{ color: 'white', fontWeight: 600 }}>Valor a Pagar</TableCell>
              <TableCell sx={{ color: 'white', fontWeight: 600, textAlign: 'center' }}>PDF</TableCell>
            </TableRow>
          </TableHead>
          <TableBody>
            {multasPaginadas.length === 0 ? (
              <TableRow>
                <TableCell colSpan={10} align="center" sx={{ py: 4 }}>
                  <Typography color="text.secondary">
                    {busca || filtroOrgao !== 'todos' 
                      ? 'Nenhuma multa encontrada com os filtros aplicados'
                      : 'Nenhuma multa encontrada'
                    }
                  </Typography>
                </TableCell>
              </TableRow>
            ) : (
              multasPaginadas.map((multa, index) => (
                <TableRow key={index} hover>
                  <TableCell>
                    <Chip label={multa.placa} size="small" color="primary" variant="outlined" />
                  </TableCell>
                  <TableCell>{multa.numero}</TableCell>
                  <TableCell sx={{ fontFamily: 'monospace', fontSize: '0.875rem' }}>
                    {multa.ait}
                  </TableCell>
                  <TableCell>
                    <Typography variant="body2" sx={{ maxWidth: 150 }}>
                      {multa.orgao_autuador}
                    </Typography>
                  </TableCell>
                  <TableCell>
                    <Typography variant="body2" sx={{ maxWidth: 400 }}>
                      {multa.motivo}
                    </Typography>
                  </TableCell>
                  <TableCell>{multa.data_infracao}</TableCell>
                  <TableCell>{multa.data_vencimento}</TableCell>
                  <TableCell sx={{ fontWeight: 500 }}>{multa.valor}</TableCell>
                  <TableCell sx={{ fontWeight: 600, color: 'error.main' }}>
                    {multa.valor_a_pagar}
                  </TableCell>
                  <TableCell align="center">
                    {onDownloadPDF && (
                      <IconButton
                        size="small"
                        color="error"
                        onClick={() => onDownloadPDF(multa.ait)}
                      >
                        <PictureAsPdfIcon />
                      </IconButton>
                    )}
                  </TableCell>
                </TableRow>
              ))
            )}
          </TableBody>
        </Table>
        <TablePagination
          component="div"
          count={multasFiltradas.length}
          page={page}
          onPageChange={handleChangePage}
          rowsPerPage={rowsPerPage}
          onRowsPerPageChange={handleChangeRowsPerPage}
          rowsPerPageOptions={[5, 10, 25, 50]}
          labelRowsPerPage="Linhas por página:"
          labelDisplayedRows={({ from, to, count }) => `${from}-${to} de ${count}`}
        />
      </TableContainer>
    </Box>
  );
}
