import { createHmac, randomBytes, timingSafeEqual } from 'node:crypto';

const COOKIE_NAME = 'lq_tepeapulco_session';
const SESSION_DURATION_SECONDS = 60 * 60 * 8;

function secret() {
  const value = String(process.env.AUTH_SECRET || '');
  if (value.length < 32) throw new Error('AUTH_SECRET no está configurada correctamente.');
  return value;
}

function encode(value) {
  return Buffer.from(value).toString('base64url');
}

function sign(payload) {
  return createHmac('sha256', secret()).update(payload).digest('base64url');
}

function secureEqual(left, right) {
  const leftBuffer = Buffer.from(String(left || ''));
  const rightBuffer = Buffer.from(String(right || ''));
  return leftBuffer.length === rightBuffer.length && timingSafeEqual(leftBuffer, rightBuffer);
}

function cookieValue(request) {
  const cookies = String(request.headers?.cookie || '').split(';');
  const sessionCookie = cookies.find(cookie => cookie.trim().startsWith(`${COOKIE_NAME}=`));
  return sessionCookie ? decodeURIComponent(sessionCookie.trim().slice(COOKIE_NAME.length + 1)) : '';
}

function isSecureRequest(request) {
  return String(request.headers?.['x-forwarded-proto'] || '').split(',')[0].trim() === 'https' || process.env.NODE_ENV === 'production';
}

function cookieOptions(request, maxAge) {
  return [
    `${COOKIE_NAME}=`,
    'Path=/',
    'HttpOnly',
    'SameSite=Strict',
    isSecureRequest(request) ? 'Secure' : '',
    `Max-Age=${maxAge}`,
  ].filter(Boolean).join('; ');
}

export function configuredUsers() {
  const users = [
    {
      username: String(process.env.APP_ADMIN_USERNAME || 'administrador').trim(),
      password: String(process.env.APP_ADMIN_PASSWORD || ''),
      role: 'admin',
      displayName: 'Administrador',
    },
    {
      username: String(process.env.APP_EMPLOYEE_USERNAME || 'empleado').trim(),
      password: String(process.env.APP_EMPLOYEE_PASSWORD || ''),
      role: 'employee',
      displayName: 'Empleado',
    },
  ];
  if (users.some(user => !user.username || !user.password)) throw new Error('Las cuentas de acceso no están configuradas.');
  return users;
}

export function authenticate(username, password) {
  const normalizedUsername = String(username || '').trim().toLowerCase();
  return configuredUsers().find(user => secureEqual(user.username.toLowerCase(), normalizedUsername) && secureEqual(user.password, password)) || null;
}

export function createSession(request, user) {
  const payload = encode(JSON.stringify({
    username: user.username,
    role: user.role,
    displayName: user.displayName,
    expiresAt: Math.floor(Date.now() / 1000) + SESSION_DURATION_SECONDS,
    nonce: randomBytes(12).toString('base64url'),
  }));
  const token = `${payload}.${sign(payload)}`;
  return cookieOptions(request, SESSION_DURATION_SECONDS).replace(`${COOKIE_NAME}=`, `${COOKIE_NAME}=${encodeURIComponent(token)}`);
}

export function clearSession(request) {
  return cookieOptions(request, 0);
}

export function getSession(request) {
  try {
    const token = cookieValue(request);
    const [payload, signature] = token.split('.');
    if (!payload || !signature || !secureEqual(sign(payload), signature)) return null;
    const session = JSON.parse(Buffer.from(payload, 'base64url').toString('utf8'));
    if (!session.expiresAt || session.expiresAt <= Math.floor(Date.now() / 1000)) return null;
    if (!['admin', 'employee'].includes(session.role)) return null;
    return session;
  } catch {
    return null;
  }
}

export function publicUser(user) {
  return { username: user.username, role: user.role, displayName: user.displayName };
}

export function requireSession(request, response, roles = ['admin', 'employee']) {
  const session = getSession(request);
  if (!session) {
    response.status(401).json({ error: 'Tu sesión no es válida. Inicia sesión nuevamente.' });
    return null;
  }
  if (!roles.includes(session.role)) {
    response.status(403).json({ error: 'No tienes permiso para realizar esta acción.' });
    return null;
  }
  return session;
}
