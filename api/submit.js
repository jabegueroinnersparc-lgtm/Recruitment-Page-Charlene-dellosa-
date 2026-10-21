/**
 * Vercel serverless proxy for the Google Apps Script backend.
 */

module.exports = async function handler(req, res) {
  if (req.method !== 'POST') {
    res.setHeader('Allow', 'POST');
    return res.status(405).json({
      ok: false,
      error: 'Method not allowed'
    });
  }

  const target =
    'https://script.google.com/macros/s/AKfycbwNT9B4YqIATpLPQTTTgnRhqfAuL8ilXivpe3RGjLo0teuuVFJMKg5qc-xaSGselNBZ/exec';

  try {
    const response = await fetch(target, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json'
      },
      body: JSON.stringify(req.body || {})
    });

    const text = await response.text();

    let payload;

    try {
      payload = text ? JSON.parse(text) : {
        ok: false,
        error: 'Empty response from Google Apps Script.'
      };
    } catch (parseError) {
      console.error('Invalid Google Apps Script response:', text);

      payload = {
        ok: false,
        error: 'Invalid response from Google Apps Script.'
      };
    }

    return res.status(response.ok ? 200 : 502).json(payload);
  } catch (error) {
    console.error('Google Apps Script connection error:', error);

    return res.status(502).json({
      ok: false,
      error: 'Unable to connect to Google Apps Script.'
    });
  }
};
