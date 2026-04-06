const webpush = require('web-push');

const VAPID_PUBLIC_KEY = process.env.VAPID_PUBLIC_KEY;
const VAPID_PRIVATE_KEY = process.env.VAPID_PRIVATE_KEY;
const WEBAPP_URL        = process.env.WEBAPP_URL;

if (VAPID_PUBLIC_KEY && VAPID_PRIVATE_KEY) {
  webpush.setVapidDetails('mailto:m.elamraoui@ocpgroup.ma', VAPID_PUBLIC_KEY, VAPID_PRIVATE_KEY);
}

exports.handler = async (event) => {
  const headers = { 'Access-Control-Allow-Origin': '*', 'Access-Control-Allow-Headers': 'Content-Type' };

  if (event.httpMethod === 'OPTIONS') return { statusCode: 204, headers };
  if (event.httpMethod !== 'POST') return { statusCode: 405, headers, body: 'Method Not Allowed' };

  try {
    const { title, body, targetProfile } = JSON.parse(event.body || '{}');
    if (!title) return { statusCode: 400, headers, body: 'Missing title' };

    // Récupérer les abonnements depuis Google Sheets
    const resp = await fetch(`${WEBAPP_URL}?action=getSubscriptions&profile=${encodeURIComponent(targetProfile || 'admin')}&t=${Date.now()}`);
    if (!resp.ok) return { statusCode: 502, headers, body: 'Failed to fetch subscriptions' };

    const subscriptions = await resp.json();
    if (!Array.isArray(subscriptions) || subscriptions.length === 0) {
      return { statusCode: 200, headers, body: JSON.stringify({ sent: 0 }) };
    }

    const payload = JSON.stringify({ title, body });
    const results = await Promise.allSettled(
      subscriptions.map(sub => {
        try { return webpush.sendNotification(JSON.parse(sub.subscription), payload); }
        catch(e) { return Promise.reject(e); }
      })
    );

    const sent = results.filter(r => r.status === 'fulfilled').length;
    return { statusCode: 200, headers, body: JSON.stringify({ sent, total: subscriptions.length }) };
  } catch(e) {
    return { statusCode: 500, headers, body: e.message };
  }
};
