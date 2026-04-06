// Vercel serverless function — proxy générique Google Sheets (contourne CORS)
// Chaque route /api/proxy-* appelle ce même handler avec une URL cible différente

const TARGETS = {
  'proxy-sap':          'https://docs.google.com/spreadsheets/d/1aQAvb1DUv6Vk1Y1C-WEYgQnYN1BxujEAg8lbMt1sP3s/export?format=xlsx',
  'proxy-arrets':       'https://docs.google.com/spreadsheets/d/1C9bYkPsoYg81ARgolVDlZRwsMZk4Seff6aC7vfxoVeE/export?format=xlsx',
  'proxy-travhebdo':    'https://docs.google.com/spreadsheets/d/1C9bYkPsoYg81ARgolVDlZRwsMZk4Seff6aC7vfxoVeE/gviz/tq?tqx=out:csv&sheet=Travaux%20hebdomadaire',
  'proxy-installation': 'https://docs.google.com/spreadsheets/d/1rOPljpAHYIs_uQ5-EUnL4yVwy2ciKAn20htExnU2vG4/gviz/tq?tqx=out:csv&sheet=Installation',
  'proxy-code':         'https://docs.google.com/spreadsheets/d/1rOPljpAHYIs_uQ5-EUnL4yVwy2ciKAn20htExnU2vG4/gviz/tq?tqx=out:csv&sheet=Code',
  'proxy-calarrets':    'https://docs.google.com/spreadsheets/d/1EBACM8ou8B_9fmExToUKsMCvHL27hiwU2D0yZ_gQGOA/gviz/tq?tqx=out:csv&sheet=Planning%20des%20arrets',
};

module.exports = async function handler(req, res) {
  // Extraire le nom de la route depuis l'URL (ex: /api/proxy-sap → proxy-sap)
  const routeName = req.url.replace(/^\/api\//, '').split('?')[0];
  const targetUrl = TARGETS[routeName];

  if (!targetUrl) {
    res.status(404).json({ error: 'Unknown proxy route: ' + routeName });
    return;
  }

  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'GET, OPTIONS');

  if (req.method === 'OPTIONS') { res.status(204).end(); return; }

  try {
    const response = await fetch(targetUrl, {
      headers: {
        'User-Agent': 'Mozilla/5.0',
        'Accept': '*/*',
      },
      redirect: 'follow',
    });

    if (!response.ok) {
      res.status(response.status).send('Upstream error: ' + response.statusText);
      return;
    }

    const contentType = response.headers.get('content-type') || 'application/octet-stream';
    res.setHeader('Content-Type', contentType);
    res.setHeader('Cache-Control', 'public, max-age=120'); // cache 2 min

    const buffer = await response.arrayBuffer();
    res.status(200).send(Buffer.from(buffer));
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
};
