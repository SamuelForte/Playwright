'use client';

import React, { useCallback } from 'react';
import { useDropzone } from 'react-dropzone';
import {
  Box,
  Paper,
  Typography,
  Button,
  List,
  ListItem,
  ListItemText,
  Alert,
} from '@mui/material';
import CloudUploadIcon from '@mui/icons-material/CloudUpload';
import * as XLSX from 'xlsx';
import { Veiculo } from '@/lib/api';

interface FileUploadProps {
  onVeiculosImportados: (veiculos: Veiculo[]) => void;
}

export default function FileUpload({ onVeiculosImportados }: FileUploadProps) {
  const [preview, setPreview] = React.useState<Veiculo[]>([]);
  const [erro, setErro] = React.useState('');

  const processarArquivo = async (file: File) => {
    setErro('');
    
    try {
      const data = await file.arrayBuffer();
      const workbook = XLSX.read(data);
      const sheetName = workbook.SheetNames[0];
      const worksheet = workbook.Sheets[sheetName];
      const jsonData = XLSX.utils.sheet_to_json(worksheet);

      const veiculos: Veiculo[] = [];

      for (const row of jsonData as any[]) {
        const placa = String(row.placa || row.Placa || '').trim().toUpperCase();
        const renavam = String(row.renavam || row.Renavam || row.RENAVAM || '').trim();

        if (placa && renavam) {
          veiculos.push({ placa, renavam });
        }
      }

      if (veiculos.length === 0) {
        setErro('Nenhum veículo válido encontrado. Certifique-se de que o arquivo tem colunas "placa" e "renavam".');
        return;
      }

      setPreview(veiculos);
    } catch (error) {
      setErro('Erro ao processar arquivo. Certifique-se de que é um arquivo Excel válido.');
    }
  };

  const onDrop = useCallback((acceptedFiles: File[]) => {
    if (acceptedFiles.length > 0) {
      processarArquivo(acceptedFiles[0]);
    }
  }, []);

  const { getRootProps, getInputProps, isDragActive } = useDropzone({
    onDrop,
    accept: {
      'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet': ['.xlsx'],
      'application/vnd.ms-excel': ['.xls'],
      'text/csv': ['.csv'],
    },
    maxFiles: 1,
  });

  const handleConfirmarImportacao = () => {
    onVeiculosImportados(preview);
    setPreview([]);
  };

  return (
    <Box>
      <Paper
        {...getRootProps()}
        elevation={2}
        sx={{
          p: 4,
          textAlign: 'center',
          cursor: 'pointer',
          border: '2px dashed',
          borderColor: isDragActive ? 'primary.main' : 'grey.300',
          bgcolor: isDragActive ? 'action.hover' : 'background.paper',
          transition: 'all 0.2s',
          '&:hover': {
            borderColor: 'primary.main',
            bgcolor: 'action.hover',
          },
        }}
      >
        <input {...getInputProps()} />
        <CloudUploadIcon sx={{ fontSize: 48, color: 'primary.main', mb: 2 }} />
        <Typography variant="h6" gutterBottom>
          {isDragActive ? 'Solte o arquivo aqui' : 'Arraste um arquivo Excel ou clique para selecionar'}
        </Typography>
        <Typography variant="body2" color="text.secondary">
          Formatos aceitos: .xlsx, .xls, .csv
        </Typography>
        <Typography variant="caption" display="block" sx={{ mt: 1 }} color="text.secondary">
          O arquivo deve conter as colunas: <strong>placa</strong> e <strong>renavam</strong>
        </Typography>
      </Paper>

      {erro && (
        <Alert severity="error" sx={{ mt: 2 }}>
          {erro}
        </Alert>
      )}

      {preview.length > 0 && (
        <Paper elevation={2} sx={{ mt: 3, p: 2 }}>
          <Typography variant="h6" gutterBottom>
            Preview - {preview.length} veículo(s) encontrado(s)
          </Typography>
          
          <List sx={{ maxHeight: 300, overflow: 'auto', mb: 2 }}>
            {preview.slice(0, 10).map((veiculo, index) => (
              <ListItem key={index} divider>
                <ListItemText
                  primary={veiculo.placa}
                  secondary={`RENAVAM: ${veiculo.renavam}`}
                />
              </ListItem>
            ))}
            {preview.length > 10 && (
              <ListItem>
                <ListItemText
                  secondary={`... e mais ${preview.length - 10} veículo(s)`}
                  sx={{ textAlign: 'center', fontStyle: 'italic' }}
                />
              </ListItem>
            )}
          </List>

          <Button
            variant="contained"
            fullWidth
            onClick={handleConfirmarImportacao}
          >
            Confirmar Importação
          </Button>
        </Paper>
      )}
    </Box>
  );
}
