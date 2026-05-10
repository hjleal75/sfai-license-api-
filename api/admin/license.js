// api/admin/license.js — crear / actualizar / revocar / eliminar
// POST /api/admin/license
// Header: X-Admin-Token: <token>
// Body:   { action, email, key?, tier?, expira?, cliente?, active? }

import { getLicense, setLicense, deleteLicense } from '../../lib/redis.js';

export default async function handler(req, res) {
  if (req.method === 'OPTIONS') return res.status(204).end();

  if (req.method !== 'POST') {
    return res.status(405).json({ ok: false, reason: 'method_not_allowed' });
  }

  // ── Auth ─────────────────────────────────────────────────────
  const adminToken = req.headers['x-admin-token'];
  if (!adminToken || adminToken !== process.env.ADMIN_TOKEN) {
    return res.status(403).json({ ok: false, reason: 'forbidden' });
  }

  const { action, email, key, tier, expira, cliente, active } = req.body || {};

  if (!email) {
    return res.status(400).json({ ok: false, reason: 'email_required' });
  }

  const emailNorm = email.toLowerCase().trim();

  // ── create / update ──────────────────────────────────────────
  if (action === 'create' || action === 'update') {
    const record = {
      email:   emailNorm,
      key:     key     || '',
      tier:    tier    || 'FREE',
      expira:  expira  || '2099-12-31',
      cliente: cliente || '',
      active:  active !== undefined ? active : true,
      created: new Date().toISOString(),
    };
    await setLicense(emailNorm, record);
    return res.status(200).json({ ok: true, action, record });
  }

  // ── revoke ───────────────────────────────────────────────────
  if (action === 'revoke') {
    const existing = await getLicense(emailNorm);
    if (!existing) {
      return res.status(404).json({ ok: false, reason: 'not_found' });
    }
    existing.active = false;
    await setLicense(emailNorm, existing);
    return res.status(200).json({ ok: true, action: 'revoked', email: emailNorm });
  }

  // ── reactivate ───────────────────────────────────────────────
  if (action === 'reactivate') {
    const existing = await getLicense(emailNorm);
    if (!existing) {
      return res.status(404).json({ ok: false, reason: 'not_found' });
    }
    existing.active = true;
    await setLicense(emailNorm, existing);
    return res.status(200).json({ ok: true, action: 'reactivated', email: emailNorm });
  }

  // ── delete ───────────────────────────────────────────────────
  if (action === 'delete') {
    await deleteLicense(emailNorm);
    return res.status(200).json({ ok: true, action: 'deleted', email: emailNorm });
  }

  return res.status(400).json({ ok: false, reason: 'invalid_action' });
}
