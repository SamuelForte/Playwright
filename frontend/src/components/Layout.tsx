'use client';

import React, { useEffect, useState } from 'react';
import Link from 'next/link';
import {
  AppBar,
  Toolbar,
  Typography,
  Button,
  Box,
  Container,
} from '@mui/material';
import DirectionsCarIcon from '@mui/icons-material/DirectionsCar';
import DashboardIcon from '@mui/icons-material/Dashboard';
import AddCircleOutlineIcon from '@mui/icons-material/AddCircleOutline';
import HistoryIcon from '@mui/icons-material/History';
import HowToRegIcon from '@mui/icons-material/HowToReg';
import VerifiedUserIcon from '@mui/icons-material/VerifiedUser';

export default function Layout({ children }: { children: React.ReactNode }) {
  const [basePath, setBasePath] = useState('');

  useEffect(() => {
    const pathname = window.location.pathname;
    setBasePath(pathname.startsWith('/Playwright') ? '/Playwright' : '');
  }, []);

  const withBasePath = (path: string) => {
    const normalized = path.endsWith('/') ? path : `${path}/`;
    return `${basePath}${normalized}`;
  };

  return (
    <Box sx={{ display: 'flex', flexDirection: 'column', minHeight: '100vh' }}>
      <AppBar position="static" elevation={2}>
        <Toolbar>
          <DirectionsCarIcon sx={{ mr: 2 }} />
          <Typography variant="h6" component="div" sx={{ flexGrow: 1 }}>
            DETRAN-CE Consulta
          </Typography>
          
          <Box sx={{ display: 'flex', gap: 1 }}>
            <Button
              color="inherit"
              component={Link}
              href={withBasePath('/dashboard')}
              startIcon={<DashboardIcon />}
            >
              Dashboard
            </Button>
            <Button
              color="inherit"
              component={Link}
              href={withBasePath('/nova-consulta')}
              startIcon={<AddCircleOutlineIcon />}
            >
              Nova Consulta
            </Button>
            <Button
              color="inherit"
              component={Link}
              href={withBasePath('/historico')}
              startIcon={<HistoryIcon />}
            >
              Histórico
            </Button>
            <Button
              color="inherit"
              component={Link}
              href={withBasePath('/indicacao')}
              startIcon={<HowToRegIcon />}
            >
              Indicação
            </Button>
            <Button
              color="inherit"
              component={Link}
              href={withBasePath('/condutor-indicacao')}
              startIcon={<VerifiedUserIcon />}
            >
              Responder
            </Button>
          </Box>
        </Toolbar>
      </AppBar>

      <Container maxWidth="xl" sx={{ mt: 4, mb: 4, flex: 1 }}>
        {children}
      </Container>

      <Box
        component="footer"
        sx={{
          py: 3,
          px: 2,
          mt: 'auto',
          backgroundColor: 'background.paper',
          borderTop: '1px solid',
          borderColor: 'divider',
        }}
      >
        <Container maxWidth="xl">
          <Typography variant="body2" color="text.secondary" align="center">
            © {new Date().getFullYear()} DETRAN-CE Consulta - Sistema de Consulta de Multas
          </Typography>
        </Container>
      </Box>
    </Box>
  );
}
