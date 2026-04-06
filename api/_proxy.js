// Helper générique — appelé par chaque fichier avec son URL cible
module.exports = function makeProxy(targetUrl) {
  return async function handler(req, res) {
    res.setHeader('Access-Control-Allow-Origin', '*');
    res.setHeader('Access-Control-Allow-Methods', 'GET, OPTIONS');
    if (req.method === 'OPTIONS') { res.status(204).end(); return; }

    try {
      const response = await fetch(targetUrl, {
        headers: { 'User-Agent': 'Mozilla/5.0', 'Accept': '*/*' },
        redirect: 'follow',
      });

      if (!response.ok) {
        res.status(response.status).send('Upstream error: ' + response.statusText);
        return;
      }

      const contentType = response.headers.get('content-type') || 'application/octet-stream';
      res.setHeader('Content-Type', contentType);
      res.setHeader('Cache-Control', 'public, max-age=120');

      const buffer = await response.arrayBuffer();
      res.status(200).send(Buffer.from(buffer));
    } catch (err) {
      res.status(500).json({ error: err.message });
    }
  };
};
