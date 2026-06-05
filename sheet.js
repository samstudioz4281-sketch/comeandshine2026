// Vercel serverless function — proxies all requests to Apps Script
// Set APPS_SCRIPT_URL in Vercel environment variables

export default async function handler(req, res) {
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'GET, POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');

  if (req.method === 'OPTIONS') return res.status(200).end();

  const SCRIPT_URL = process.env.APPS_SCRIPT_URL;
  if (!SCRIPT_URL) return res.status(500).json({ error: 'APPS_SCRIPT_URL not set' });

  try {
    if (req.method === 'GET') {
      const r = await fetch(SCRIPT_URL + '?action=nextserial');
      const data = await r.json();
      return res.status(200).json(data);
    }

    if (req.method === 'POST') {
      const body = req.body;
      const r = await fetch(SCRIPT_URL, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(body)
      });
      const text = await r.text();
      try {
        return res.status(200).json(JSON.parse(text));
      } catch {
        return res.status(200).json({ status: 'ok', raw: text });
      }
    }
  } catch (err) {
    return res.status(500).json({ status: 'error', message: err.message });
  }
}
