// api/health.js
export default function handler(req, res) {
  if (req.method === 'OPTIONS') return res.status(204).end();
  return res.status(200).json({
    status:  'ok',
    product: 'sfai-license-api',
    version: '1.0.0',
    ts:      new Date().toISOString(),
  });
}