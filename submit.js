/**
 * Vercel serverless proxy for the Google Apps Script backend.
 * Set APPS_SCRIPT_URL in Vercel to the deployed Apps Script web-app URL.
 */
module.exports = async function handler(req, res) {
  if (req.method !== 'POST') {
    res.setHeader('Allow', 'POST');
    return res.status(405).json({ ok: false, error: 'Method not allowed' });
  }

  const target = process.env.APPS_SCRIPT_URL;
  if (!target) {
    return res.status(503).json({ ok: false, error: 'The application endpoint is not configured.' });
  }

  try {
    const response = await fetch(target, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(req.body || {})
    });
    const text = await response.text();
    let payload;
    try { payload = JSON.parse(text); } catch (e) { payload = { ok: false, error: 'Invalid backend response.' }; }
    return res.status(response.ok ? 200 : 502).json(payload);
  } catch (error) {
    console.error('Apps Script proxy error:', error);
    return res.status(502).json({ ok: false, error: 'Unable to reach the application endpoint.' });
  }
};
