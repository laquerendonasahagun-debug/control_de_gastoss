import { authenticate, clearSession, createSession, getSession, publicUser } from '../lib/auth.js';

function json(response, status, payload) {
  response.status(status).json(payload);
}

export default async function handler(request, response) {
  response.setHeader('Cache-Control', 'no-store, max-age=0');
  response.setHeader('Content-Type', 'application/json; charset=utf-8');

  try {
    if (request.method === 'GET') {
      const session = getSession(request);
      return session
        ? json(response, 200, { user: publicUser(session) })
        : json(response, 401, { error: 'Inicia sesión para continuar.' });
    }

    if (request.method === 'POST') {
      const body = typeof request.body === 'string' ? JSON.parse(request.body) : request.body;
      const user = authenticate(body?.username, body?.password);
      if (!user) return json(response, 401, { error: 'Usuario o contraseña incorrectos.' });
      response.setHeader('Set-Cookie', createSession(request, user));
      return json(response, 200, { user: publicUser(user) });
    }

    if (request.method === 'DELETE') {
      response.setHeader('Set-Cookie', clearSession(request));
      return json(response, 200, { signedOut: true });
    }

    response.setHeader('Allow', 'GET, POST, DELETE');
    return json(response, 405, { error: 'Método no permitido.' });
  } catch (error) {
    console.error('Error en la API de autenticación:', error);
    return json(response, 503, { error: 'El inicio de sesión todavía no está configurado.' });
  }
}
