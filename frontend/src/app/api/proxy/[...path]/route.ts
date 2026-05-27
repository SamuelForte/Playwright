import { NextRequest, NextResponse } from 'next/server';

const configuredBackend = process.env.BACKEND_URL || process.env.NEXT_PUBLIC_API_URL;
const LOCAL_BACKEND_URL = configuredBackend && /^https?:\/\//i.test(configuredBackend)
  ? configuredBackend
  : 'http://localhost:8000';
const REMOTE_BACKEND_URL = 'https://detran-api-playwright-production.up.railway.app';

async function forwardRequest(request: NextRequest, targetBaseUrl: string, pathSegments: string[]) {
  const targetUrl = new URL(pathSegments.join('/'), `${targetBaseUrl.replace(/\/$/, '')}/`);
  targetUrl.search = request.nextUrl.search;

  const headers = new Headers(request.headers);
  headers.delete('host');
  headers.delete('connection');
  headers.delete('content-length');

  const init: RequestInit = {
    method: request.method,
    headers,
    cache: 'no-store',
  };

  if (!['GET', 'HEAD'].includes(request.method)) {
    init.body = await request.text();
  }

  const response = await fetch(targetUrl.toString(), init);
  const responseHeaders = new Headers(response.headers);
  responseHeaders.delete('content-encoding');
  responseHeaders.delete('transfer-encoding');

  return new NextResponse(response.body, {
    status: response.status,
    headers: responseHeaders,
  });
}

async function proxyRequest(request: NextRequest, pathSegments: string[]) {
  try {
    return await forwardRequest(request, LOCAL_BACKEND_URL, pathSegments);
  } catch (localError) {
    console.warn(`Falha no backend local em ${LOCAL_BACKEND_URL}; tentando Railway.`, localError);

    try {
      return await forwardRequest(request, REMOTE_BACKEND_URL, pathSegments);
    } catch (remoteError) {
      console.error(`Falha ao encaminhar requisição para ${REMOTE_BACKEND_URL}:`, remoteError);

      return NextResponse.json(
        {
          detail: `Não foi possível conectar ao backend local em ${LOCAL_BACKEND_URL} nem ao backend remoto.`,
        },
        { status: 502 }
      );
    }
  }
}

export async function GET(request: NextRequest, context: { params: Promise<{ path: string[] }> }) {
  const { path } = await context.params;
  return proxyRequest(request, path);
}

export async function POST(request: NextRequest, context: { params: Promise<{ path: string[] }> }) {
  const { path } = await context.params;
  return proxyRequest(request, path);
}

export async function PUT(request: NextRequest, context: { params: Promise<{ path: string[] }> }) {
  const { path } = await context.params;
  return proxyRequest(request, path);
}

export async function PATCH(request: NextRequest, context: { params: Promise<{ path: string[] }> }) {
  const { path } = await context.params;
  return proxyRequest(request, path);
}

export async function DELETE(request: NextRequest, context: { params: Promise<{ path: string[] }> }) {
  const { path } = await context.params;
  return proxyRequest(request, path);
}

export async function OPTIONS() {
  return new NextResponse(null, {
    status: 204,
    headers: {
      'Access-Control-Allow-Origin': '*',
      'Access-Control-Allow-Methods': 'GET,POST,PUT,PATCH,DELETE,OPTIONS',
      'Access-Control-Allow-Headers': 'Content-Type, Authorization',
    },
  });
}
