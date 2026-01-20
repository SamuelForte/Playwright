'use client';

import React, { useState } from 'react';
import {
  Box,
  TextField,
  Button,
  IconButton,
  Paper,
  Typography,
  List,
  ListItem,
  ListItemText,
  ListItemSecondaryAction,
  Alert,
} from '@mui/material';
import DeleteIcon from '@mui/icons-material/Delete';
import AddIcon from '@mui/icons-material/Add';
import { Veiculo } from '@/lib/api';

interface VehicleFormProps {
  onVeiculosChange: (veiculos: Veiculo[]) => void;
  veiculos: Veiculo[];
}

export default function VehicleForm({ onVeiculosChange, veiculos }: VehicleFormProps) {
  const [placa, setPlaca] = useState('');
  const [renavam, setRenavam] = useState('');
  const [erro, setErro] = useState('');

  const formatarPlaca = (value: string) => {
    // Remove caracteres não alfanuméricos e converte para maiúsculas
    let formatted = value.replace(/[^A-Z0-9]/gi, '').toUpperCase();
    
    // Formato Mercosul: AAA0A00 ou Antigo: AAA0000
    if (formatted.length <= 7) {
      return formatted;
    }
    return formatted.substring(0, 7);
  };

  const formatarRenavam = (value: string) => {
    // Remove não números
    let formatted = value.replace(/\D/g, '');
    
    // Máximo 11 dígitos
    if (formatted.length <= 11) {
      return formatted;
    }
    return formatted.substring(0, 11);
  };

  const validarPlaca = (placa: string) => {
    // Placa antiga: AAA0000 ou Mercosul: AAA0A00
    const regexAntiga = /^[A-Z]{3}\d{4}$/;
    const regexMercosul = /^[A-Z]{3}\d[A-Z]\d{2}$/;
    
    return regexAntiga.test(placa) || regexMercosul.test(placa);
  };

  const validarRenavam = (renavam: string) => {
    return renavam.length === 11 && /^\d+$/.test(renavam);
  };

  const handleAdicionarVeiculo = () => {
    setErro('');

    if (!placa || !renavam) {
      setErro('Preencha placa e RENAVAM');
      return;
    }

    if (!validarPlaca(placa)) {
      setErro('Placa inválida. Use formato AAA0000 ou AAA0A00');
      return;
    }

    if (!validarRenavam(renavam)) {
      setErro('RENAVAM deve ter 11 dígitos');
      return;
    }

    // Verifica se já existe
    if (veiculos.some(v => v.placa === placa)) {
      setErro('Placa já adicionada');
      return;
    }

    const novosVeiculos = [...veiculos, { placa, renavam }];
    onVeiculosChange(novosVeiculos);
    
    setPlaca('');
    setRenavam('');
  };

  const handleRemoverVeiculo = (index: number) => {
    const novosVeiculos = veiculos.filter((_, i) => i !== index);
    onVeiculosChange(novosVeiculos);
  };

  const handleKeyPress = (e: React.KeyboardEvent) => {
    if (e.key === 'Enter') {
      handleAdicionarVeiculo();
    }
  };

  return (
    <Paper elevation={2} sx={{ p: 3 }}>
      <Typography variant="h6" gutterBottom>
        Adicionar Veículos
      </Typography>

      <Box sx={{ display: 'flex', gap: 2, mb: 2 }}>
        <TextField
          label="Placa"
          placeholder="ABC1234"
          value={placa}
          onChange={(e) => setPlaca(formatarPlaca(e.target.value))}
          onKeyPress={handleKeyPress}
          fullWidth
          inputProps={{ maxLength: 7 }}
        />
        <TextField
          label="RENAVAM"
          placeholder="01365705622"
          value={renavam}
          onChange={(e) => setRenavam(formatarRenavam(e.target.value))}
          onKeyPress={handleKeyPress}
          fullWidth
          inputProps={{ maxLength: 11 }}
        />
        <Button
          variant="contained"
          startIcon={<AddIcon />}
          onClick={handleAdicionarVeiculo}
          sx={{ minWidth: 140 }}
        >
          Adicionar
        </Button>
      </Box>

      {erro && (
        <Alert severity="error" sx={{ mb: 2 }}>
          {erro}
        </Alert>
      )}

      {veiculos.length > 0 && (
        <>
          <Typography variant="subtitle2" color="text.secondary" sx={{ mt: 3, mb: 1 }}>
            Veículos adicionados ({veiculos.length})
          </Typography>
          <List>
            {veiculos.map((veiculo, index) => (
              <ListItem
                key={index}
                sx={{
                  bgcolor: 'background.default',
                  mb: 1,
                  borderRadius: 1,
                }}
              >
                <ListItemText
                  primary={veiculo.placa}
                  secondary={`RENAVAM: ${veiculo.renavam}`}
                />
                <ListItemSecondaryAction>
                  <IconButton
                    edge="end"
                    onClick={() => handleRemoverVeiculo(index)}
                    color="error"
                  >
                    <DeleteIcon />
                  </IconButton>
                </ListItemSecondaryAction>
              </ListItem>
            ))}
          </List>
        </>
      )}
    </Paper>
  );
}
