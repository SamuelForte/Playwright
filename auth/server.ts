import express, { Request, Response, NextFunction } from 'express';
import cors from 'cors';
import jwt from 'jsonwebtoken';
import axios from 'axios';
import { PrismaClient } from '@prisma/client';
import pino from 'pino';
import dotenv from 'dotenv';

dotenv.config();

const app = express();
const prisma = new PrismaClient();
const logger = pino({ level: 'info' });

// ==================== CONFIGURAÇÃO ====================

const PORT = process.env.AUTH_PORT || 4000;
const GOOGLE_CLIENT_ID = process.env.GOOGLE_CLIENT_ID!;
const GOOGLE_CLIENT_SECRET = process.env.GOOGLE_CLIENT_SECRET!;
const JWT_SECRET = process.env.JWT_SECRET!;
const CLIENT_URL = process.env.CLIENT_URL || 'http://localhost:3000';
const REDIRECT_URI = `http://localhost:${PORT}/auth/google/callback`;

if (!GOOGLE_CLIENT_ID || !GOOGLE_CLIENT_SECRET || !JWT_SECRET) {
  logger.error('Missing required environment variables');
  process.exit(1);
}

// ==================== MIDDLEWARES ====================

app.use(cors({ origin: CLIENT_URL, credentials: true }));
app.use(express.json());

// Middleware de logging
app.use((req: Request, res: Response, next: NextFunction) => {
  logger.info({ method: req.method, url: req.url }, 'Request received');
  next();
});

// ==================== TIPOS ====================

interface GoogleTokenResponse {
  access_token: string;
  refresh_token?: string;
  expires_in: number;
  scope: string;
  token_type: string;
  id_token: string;
}

interface GoogleUserInfo {
  id: string;
  email: string;
  verified_email: boolean;
  name: string;
  given_name: string;
  family_name: string;
  picture: string;
  locale: string;
}

interface JWTPayload {
  userId: string;
  email: string;
}

interface AuthRequest extends Request {
  user?: JWTPayload;
}

// ==================== FUNÇÕES AUXILIARES ====================

function generateJWT(userId: string, email: string): string {
  return jwt.sign({ userId, email }, JWT_SECRET, { expiresIn: '7d' });
}

function verifyJWT(token: string): JWTPayload {
  return jwt.verify(token, JWT_SECRET) as JWTPayload;
}

async function exchangeCodeForTokens(code: string): Promise<GoogleTokenResponse> {
  const { data } = await axios.post<GoogleTokenResponse>(
    'https://oauth2.googleapis.com/token',
    {
      code,
      client_id: GOOGLE_CLIENT_ID,
      client_secret: GOOGLE_CLIENT_SECRET,
      redirect_uri: REDIRECT_URI,
      grant_type: 'authorization_code',
    },
    {
      headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
    }
  );
  return data;
}

async function fetchGoogleUserInfo(accessToken: string): Promise<GoogleUserInfo> {
  const { data } = await axios.get<GoogleUserInfo>(
    'https://www.googleapis.com/oauth2/v2/userinfo',
    {
      headers: { Authorization: `Bearer ${accessToken}` },
    }
  );
  return data;
}

// ==================== MIDDLEWARE DE AUTENTICAÇÃO ====================

export const authMiddleware = (req: AuthRequest, res: Response, next: NextFunction) => {
  try {
    const authHeader = req.headers.authorization;

    if (!authHeader || !authHeader.startsWith('Bearer ')) {
      logger.warn('Missing or invalid Authorization header');
      return res.status(401).json({ error: 'Token não fornecido' });
    }

    const token = authHeader.substring(7);
    const payload = verifyJWT(token);

    req.user = payload;
    next();
  } catch (error) {
    if (error instanceof jwt.JsonWebTokenError) {
      logger.warn({ error: error.message }, 'Invalid JWT token');
      return res.status(401).json({ error: 'Token inválido' });
    }
    if (error instanceof jwt.TokenExpiredError) {
      logger.warn('JWT token expired');
      return res.status(401).json({ error: 'Token expirado' });
    }
    logger.error({ error }, 'Auth middleware error');
    return res.status(500).json({ error: 'Erro na autenticação' });
  }
};

// ==================== ROTAS DE AUTENTICAÇÃO ====================

/**
 * GET /auth/google
 * Redireciona para a tela de consentimento do Google
 */
app.get('/auth/google', (req: Request, res: Response) => {
  const scopes = [
    'https://www.googleapis.com/auth/userinfo.email',
    'https://www.googleapis.com/auth/userinfo.profile',
  ];

  const authUrl = `https://accounts.google.com/o/oauth2/v2/auth?${new URLSearchParams({
    client_id: GOOGLE_CLIENT_ID,
    redirect_uri: REDIRECT_URI,
    response_type: 'code',
    scope: scopes.join(' '),
    access_type: 'offline',
    prompt: 'consent',
  }).toString()}`;

  logger.info('Redirecting to Google OAuth');
  res.redirect(authUrl);
});

/**
 * GET /auth/google/callback
 * Recebe o código do Google, troca por tokens e cria/atualiza usuário
 */
app.get('/auth/google/callback', async (req: Request, res: Response) => {
  try {
    const { code } = req.query;

    if (!code || typeof code !== 'string') {
      logger.warn('Authorization code missing');
      return res.redirect(`${CLIENT_URL}/login?error=no_code`);
    }

    // Troca código por tokens
    logger.info('Exchanging code for tokens');
    const tokens = await exchangeCodeForTokens(code);

    // Busca informações do usuário
    logger.info('Fetching user info from Google');
    const googleUser = await fetchGoogleUserInfo(tokens.access_token);

    if (!googleUser.verified_email) {
      logger.warn({ email: googleUser.email }, 'Email not verified');
      return res.redirect(`${CLIENT_URL}/login?error=email_not_verified`);
    }

    // Verifica se usuário já existe
    let user = await prisma.user.findUnique({
      where: { email: googleUser.email },
    });

    if (user) {
      // Atualiza refreshToken
      logger.info({ userId: user.id }, 'Updating existing user');
      user = await prisma.user.update({
        where: { id: user.id },
        data: {
          refreshToken: tokens.refresh_token || user.refreshToken,
          picture: googleUser.picture,
          name: googleUser.name,
        },
      });
    } else {
      // Cria novo usuário
      logger.info({ email: googleUser.email }, 'Creating new user');
      user = await prisma.user.create({
        data: {
          email: googleUser.email,
          name: googleUser.name,
          picture: googleUser.picture,
          googleId: googleUser.id,
          refreshToken: tokens.refresh_token || null,
        },
      });
    }

    // Gera JWT próprio
    const accessToken = generateJWT(user.id, user.email);

    // Redireciona para o frontend com o token
    const redirectUrl = `${CLIENT_URL}/auth/callback?token=${accessToken}`;
    logger.info({ userId: user.id }, 'Authentication successful');
    res.redirect(redirectUrl);
  } catch (error) {
    if (axios.isAxiosError(error)) {
      logger.error(
        { error: error.response?.data || error.message },
        'Google OAuth error'
      );
      return res.redirect(`${CLIENT_URL}/login?error=google_auth_failed`);
    }
    logger.error({ error }, 'Callback error');
    res.redirect(`${CLIENT_URL}/login?error=server_error`);
  }
});

/**
 * POST /auth/logout
 * Invalida o refreshToken do usuário
 */
app.post('/auth/logout', authMiddleware, async (req: AuthRequest, res: Response) => {
  try {
    const userId = req.user!.userId;

    await prisma.user.update({
      where: { id: userId },
      data: { refreshToken: null },
    });

    logger.info({ userId }, 'User logged out');
    res.status(200).json({ message: 'Logout realizado com sucesso' });
  } catch (error) {
    logger.error({ error }, 'Logout error');
    res.status(500).json({ error: 'Erro ao realizar logout' });
  }
});

// ==================== ROTAS PROTEGIDAS (EXEMPLO) ====================

/**
 * GET /auth/me
 * Retorna informações do usuário autenticado
 */
app.get('/auth/me', authMiddleware, async (req: AuthRequest, res: Response) => {
  try {
    const userId = req.user!.userId;

    const user = await prisma.user.findUnique({
      where: { id: userId },
      select: {
        id: true,
        email: true,
        name: true,
        picture: true,
        createdAt: true,
      },
    });

    if (!user) {
      logger.warn({ userId }, 'User not found');
      return res.status(404).json({ error: 'Usuário não encontrado' });
    }

    logger.info({ userId }, 'User info retrieved');
    res.status(200).json({ user });
  } catch (error) {
    logger.error({ error }, 'Get user info error');
    res.status(500).json({ error: 'Erro ao buscar informações do usuário' });
  }
});

/**
 * GET /auth/health
 * Health check
 */
app.get('/auth/health', (req: Request, res: Response) => {
  res.status(200).json({ status: 'ok', timestamp: new Date().toISOString() });
});

// ==================== ERROR HANDLER ====================

app.use((err: Error, req: Request, res: Response, next: NextFunction) => {
  logger.error({ error: err.message, stack: err.stack }, 'Unhandled error');
  res.status(500).json({ error: 'Erro interno do servidor' });
});

// ==================== INICIALIZAÇÃO ====================

const server = app.listen(PORT, () => {
  logger.info(`Auth server running on http://localhost:${PORT}`);
  logger.info(`Google OAuth callback: ${REDIRECT_URI}`);
  logger.info(`Client URL: ${CLIENT_URL}`);
});

// Graceful shutdown
process.on('SIGINT', async () => {
  logger.info('Shutting down gracefully...');
  await prisma.$disconnect();
  server.close(() => {
    logger.info('Server closed');
    process.exit(0);
  });
});

process.on('SIGTERM', async () => {
  logger.info('SIGTERM received, shutting down...');
  await prisma.$disconnect();
  server.close(() => {
    logger.info('Server closed');
    process.exit(0);
  });
});

export default app;
