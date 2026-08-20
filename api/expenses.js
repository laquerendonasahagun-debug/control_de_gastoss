import { neon } from '@neondatabase/serverless';

const MAX_BATCH_SIZE = 100;
const DATE_PATTERN = /^\d{4}-\d{2}-\d{2}$/;
let schemaPromise;

const databaseUrl = () => process.env.DATABASE_URL || process.env.POSTGRES_URL;

function getSql() {
  const connectionString = databaseUrl();
  if (!connectionString) throw new Error('DATABASE_URL no está configurada.');
  return neon(connectionString);
}

function ensureSchema(sql) {
  if (!schemaPromise) {
    schemaPromise = sql`
      CREATE TABLE IF NOT EXISTS tepeapulco_expenses (
        id TEXT PRIMARY KEY,
        expense_date DATE NOT NULL,
        category TEXT NOT NULL,
        amount NUMERIC(12, 2) NOT NULL CHECK (amount > 0),
        note TEXT NOT NULL DEFAULT 'Sin nota',
        payment TEXT NOT NULL DEFAULT 'Efectivo',
        spender TEXT NOT NULL,
        expense_type TEXT NOT NULL,
        period_id TEXT NOT NULL,
        week_index INTEGER NOT NULL CHECK (week_index >= 0),
        source TEXT NOT NULL DEFAULT 'manual',
        created_at TIMESTAMPTZ NOT NULL DEFAULT NOW(),
        updated_at TIMESTAMPTZ NOT NULL DEFAULT NOW()
      )
    `.catch(error => {
      schemaPromise = undefined;
      throw error;
    });
  }
  return schemaPromise;
}

function cleanText(value, fallback, maxLength) {
  const text = String(value ?? '').trim() || fallback;
  return text.slice(0, maxLength);
}

function normalizeEntry(entry) {
  const date = cleanText(entry?.date, '', 10);
  const amount = Number(entry?.amount);
  const weekIndex = Number(entry?.weekIndex);
  if (!DATE_PATTERN.test(date)) throw new Error('La fecha del gasto no es válida.');
  if (!Number.isFinite(amount) || amount <= 0) throw new Error('El monto debe ser mayor a cero.');
  if (!Number.isInteger(weekIndex) || weekIndex < 0) throw new Error('La semana del gasto no es válida.');

  const category = cleanText(entry?.category, '', 80);
  const spender = cleanText(entry?.spender, 'Sin especificar', 160);
  const periodId = cleanText(entry?.periodId, '', 100);
  if (!category || !periodId) throw new Error('Faltan datos obligatorios del gasto.');

  return {
    id: cleanText(entry?.id, `manual-${crypto.randomUUID()}`, 120),
    date,
    category,
    amount: Math.round(amount * 100) / 100,
    note: cleanText(entry?.note, 'Sin nota', 1000),
    payment: cleanText(entry?.payment, 'Efectivo', 80),
    spender,
    expenseType: cleanText(entry?.expenseType, 'Operativo', 40),
    periodId,
    weekIndex,
    source: 'manual',
  };
}

function serializeRow(row) {
  return {
    id: row.id,
    date: row.date,
    category: row.category,
    amount: Number(row.amount),
    note: row.note,
    payment: row.payment,
    spender: row.spender,
    expenseType: row.expense_type,
    periodId: row.period_id,
    weekIndex: Number(row.week_index),
    source: row.source,
  };
}

async function listEntries(sql) {
  const rows = await sql`
    SELECT
      id,
      TO_CHAR(expense_date, 'YYYY-MM-DD') AS date,
      category,
      amount,
      note,
      payment,
      spender,
      expense_type,
      period_id,
      week_index,
      source
    FROM tepeapulco_expenses
    ORDER BY expense_date ASC, created_at ASC, id ASC
  `;
  return rows.map(serializeRow);
}

function json(response, status, payload) {
  response.status(status).json(payload);
}

export default async function handler(request, response) {
  response.setHeader('Cache-Control', 'no-store, max-age=0');
  response.setHeader('Content-Type', 'application/json; charset=utf-8');

  if (!databaseUrl()) return json(response, 503, { error: 'La conexión con la base de datos todavía no está configurada.' });

  try {
    const sql = getSql();
    await ensureSchema(sql);

    if (request.method === 'GET') {
      return json(response, 200, { entries: await listEntries(sql) });
    }

    if (request.method === 'POST') {
      const requestBody = typeof request.body === 'string' ? JSON.parse(request.body) : request.body;
      const rawEntries = Array.isArray(requestBody?.entries) ? requestBody.entries : [requestBody?.entry].filter(Boolean);
      if (!rawEntries.length) return json(response, 400, { error: 'No se recibieron gastos.' });
      if (rawEntries.length > MAX_BATCH_SIZE) return json(response, 400, { error: `Solo se permiten ${MAX_BATCH_SIZE} gastos por envío.` });
      const entries = rawEntries.map(normalizeEntry);

      await sql.transaction(entries.map(entry => sql`
        INSERT INTO tepeapulco_expenses (
          id, expense_date, category, amount, note, payment, spender,
          expense_type, period_id, week_index, source
        ) VALUES (
          ${entry.id}, ${entry.date}, ${entry.category}, ${entry.amount}, ${entry.note},
          ${entry.payment}, ${entry.spender}, ${entry.expenseType}, ${entry.periodId},
          ${entry.weekIndex}, ${entry.source}
        )
        ON CONFLICT (id) DO UPDATE SET
          expense_date = EXCLUDED.expense_date,
          category = EXCLUDED.category,
          amount = EXCLUDED.amount,
          note = EXCLUDED.note,
          payment = EXCLUDED.payment,
          spender = EXCLUDED.spender,
          expense_type = EXCLUDED.expense_type,
          period_id = EXCLUDED.period_id,
          week_index = EXCLUDED.week_index,
          source = EXCLUDED.source,
          updated_at = NOW()
      `));

      return json(response, 200, { entries: await listEntries(sql) });
    }

    if (request.method === 'DELETE') {
      const id = cleanText(request.query?.id, '', 120);
      if (!id) return json(response, 400, { error: 'Falta el identificador del gasto.' });
      const deleted = await sql`DELETE FROM tepeapulco_expenses WHERE id = ${id} RETURNING id`;
      if (!deleted.length) return json(response, 404, { error: 'El gasto ya no existe.' });
      return json(response, 200, { deletedId: id });
    }

    response.setHeader('Allow', 'GET, POST, DELETE');
    return json(response, 405, { error: 'Método no permitido.' });
  } catch (error) {
    console.error('Error en la API de gastos:', error);
    const status = /fecha|monto|semana|obligatorios|recibieron|permiten/.test(error.message) ? 400 : 500;
    return json(response, status, { error: status === 400 ? error.message : 'No fue posible acceder a la base de datos.' });
  }
}
