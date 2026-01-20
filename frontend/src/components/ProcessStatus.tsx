'use client';

import React from 'react';
import {
  Box,
  Paper,
  Typography,
  LinearProgress,
  Chip,
} from '@mui/material';
import CheckCircleIcon from '@mui/icons-material/CheckCircle';
import ErrorIcon from '@mui/icons-material/Error';
import HourglassEmptyIcon from '@mui/icons-material/HourglassEmpty';

interface ProcessStatusProps {
  status: 'pending' | 'processing' | 'completed' | 'error';
  mensagem?: string;
  progresso?: number;
}

export default function ProcessStatus({ status, mensagem, progresso }: ProcessStatusProps) {
  const getStatusConfig = () => {
    switch (status) {
      case 'pending':
        return {
          icon: <HourglassEmptyIcon />,
          label: 'Aguardando',
          color: 'default' as const,
          bgcolor: '#FFF3E0',
        };
      case 'processing':
        return {
          icon: <HourglassEmptyIcon />,
          label: 'Processando',
          color: 'primary' as const,
          bgcolor: '#E3F2FD',
        };
      case 'completed':
        return {
          icon: <CheckCircleIcon />,
          label: 'Concluído',
          color: 'success' as const,
          bgcolor: '#E8F5E9',
        };
      case 'error':
        return {
          icon: <ErrorIcon />,
          label: 'Erro',
          color: 'error' as const,
          bgcolor: '#FFEBEE',
        };
    }
  };

  const config = getStatusConfig();

  return (
    <Paper
      elevation={1}
      sx={{
        p: 2,
        bgcolor: config.bgcolor,
        borderLeft: `4px solid`,
        borderColor: `${config.color}.main`,
      }}
    >
      <Box sx={{ display: 'flex', alignItems: 'center', mb: 1 }}>
        <Chip
          icon={config.icon}
          label={config.label}
          color={config.color}
          size="small"
          sx={{ mr: 2 }}
        />
        {mensagem && (
          <Typography variant="body2" color="text.secondary">
            {mensagem}
          </Typography>
        )}
      </Box>
      
      {status === 'processing' && (
        <LinearProgress
          variant={progresso !== undefined ? 'determinate' : 'indeterminate'}
          value={progresso}
          sx={{ mt: 1, borderRadius: 1 }}
        />
      )}
    </Paper>
  );
}
