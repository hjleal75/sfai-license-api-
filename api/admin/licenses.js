// api/admin/licenses.js — listar todas las licencias
// GET /api/admin/licenses
// Header: X-Admin-Token: <token>

import { listLicenses } from '../../lib/redis.js';

export default async function handler(req, res) {
  if (req.method === 'OPTIONS') return res.status(204).end();

  if (req.method !== 'GET') {
    return res.status(405).json({ ok: false });
  }

  const adminToken = req.headers['x-admin-token'];
  if (!adminToken || adminToken !== process.env.ADMIN_TOKEN) {
    return res.status(403).json({ ok: false, reason: 'forbidden' });
  }

  const licenses = await listLicenses();

  // Resumen de estadísticas
  const stats = {
    total:     licenses.length,
    activas:   licenses.filter(l => l.active).length,
    revocadas: licenses.filter(l => !l.active).length,
    vencidas:  licenses.filter(l => l.expira < new Date().toISOString().split('T')[0]).length,
    por_tier: {
      FREE:    licenses.filter(l => l.tier === 'FREE').length,
      PRO:     licenses.filter(l => l.tier === 'PRO').length,
      AGENCIA: licenses.filter(l => l.tier === 'AGENCIA').length,
    }
  };

  return res.status(200).json({ ok: true, stats, licenses });
}
