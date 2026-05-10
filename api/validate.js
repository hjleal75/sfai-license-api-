// api/validate.js — llamado por el VBA del Excel
// POST /api/validate
// Headers: X-Excel-Token: <token>
// Body:    { "email": "...", "key": "..." }

import { getLicense } from '../lib/redis.js';

export default async function handler(req, res) {
  if (req.method === 'OPTIONS') return res.status(204).end();

  if (req.method !== 'POST') {
    return res.status(405).json({ valid: false, reason: 'method_not_allowed' });
  }

  // ── 1. Verificar token del Excel ─────────────────────────────
  const token = req.headers['x-excel-token'];
  if (!token || token !== process.env.EXCEL_TOKEN) {
    return res.status(401).json({ valid: false, reason: 'unauthorized' });
  }

  // ── 2. Leer body ─────────────────────────────────────────────
  const { email, key } = req.body || {};
  if (!email || !key) {
    return res.status(400).json({ valid: false, reason: 'missing_fields' });
  }

  // ── 3. Buscar licencia en Upstash ────────────────────────────
  let record;
  try {
    record = await getLicense(email);
  } catch (err) {
    console.error('Redis error:', err);
    return res.status(500).json({ valid: false, reason: 'db_error' });
  }

  if (!record) {
    return res.status(200).json({ valid: false, reason: 'not_found' });
  }

  // ── 4. Validar key ───────────────────────────────────────────
  if (record.key !== key) {
    return res.status(200).json({ valid: false, reason: 'invalid_key' });
  }

  // ── 5. Validar estado ────────────────────────────────────────
  if (!record.active) {
    return res.status(200).json({ valid: false, reason: 'revoked' });
  }

  // ── 6. Validar expiración ────────────────────────────────────
  const hoy = new Date().toISOString().split('T')[0];
  if (record.expira < hoy) {
    return res.status(200).json({ valid: false, reason: 'expired' });
  }

  // ── 7. Todo OK ───────────────────────────────────────────────
  return res.status(200).json({
    valid:   true,
    tier:    record.tier,
    expira:  record.expira,
    cliente: record.cliente,
  });
}
