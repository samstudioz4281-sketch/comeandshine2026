/**
 * /api/scan  — server-side Groq proxy
 *
 * The GROQ_API_KEY env var lives only on the server.
 * The browser calls this endpoint; the key is never sent to the client.
 *
 * Body expected: { model, max_tokens, messages }   (standard Groq payload, minus auth)
 */

export default async function handler(req, res) {
  if (req.method !== 'POST') {
    return res.status(405).json({ error: 'Method not allowed' });
  }

  const apiKey = process.env.GROQ_API_KEY;
  if (!apiKey) {
    return res.status(500).json({ error: 'GROQ_API_KEY is not configured on the server.' });
  }

  try {
    const groqRes = await fetch('https://api.groq.com/openai/v1/chat/completions', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'Authorization': 'Bearer ' + apiKey,
      },
      body: JSON.stringify(req.body),
    });

    const data = await groqRes.json();

    if (!groqRes.ok) {
      return res.status(groqRes.status).json(data);
    }

    return res.status(200).json(data);
  } catch (err) {
    return res.status(500).json({ error: err.message || 'Proxy error' });
  }
}

export const config = {
  api: {
    // Allow large base64 image payloads (up to 10 MB)
    bodyParser: {
      sizeLimit: '12mb',
    },
  },
};
