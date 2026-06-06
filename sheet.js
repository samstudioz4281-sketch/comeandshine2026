// api/sheet.js — Vercel Serverless Function
// Proxies all Google Sheet operations through the Apps Script Web App
// Environment variable required: APPS_SCRIPT_URL

const APPS_SCRIPT_URL = process.env.APPS_SCRIPT_URL;

export default async function handler(req, res) {
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'GET, POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');

  if (req.method === 'OPTIONS') return res.status(200).end();

  if (!APPS_SCRIPT_URL) {
    return res.status(500).json({ error: 'APPS_SCRIPT_URL not configured' });
  }

  try {
    // ── GET: fetch next serial number ──
    if (req.method === 'GET') {
      const url = `${APPS_SCRIPT_URL}?action=nextSerial`;
      const r = await fetch(url);
      const data = await r.json();
      return res.status(200).json(data);
    }

    // ── POST: save registration / save card PDF ──
    if (req.method === 'POST') {
      const body = req.body;

      const r = await fetch(APPS_SCRIPT_URL, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(body),
      });

      const data = await r.json();
      return res.status(200).json(data);
    }

    return res.status(405).json({ error: 'Method not allowed' });
  } catch (err) {
    console.error('Sheet API error:', err);
    return res.status(500).json({ error: err.message });
  }
}
