// Vercel serverless function — Gemini chatbot
module.exports = async function handler(req, res) {
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');
  if (req.method === 'OPTIONS') { res.status(204).end(); return; }
  if (req.method !== 'POST')    { res.status(405).json({ error: 'Method not allowed' }); return; }

  const GEMINI_API_KEY = process.env.GEMINI_API_KEY;
  if (!GEMINI_API_KEY) { res.status(500).json({ error: 'Clé API manquante' }); return; }

  try {
    const { question, context } = req.body || {};
    if (!question) { res.status(400).json({ error: 'Question manquante' }); return; }

    const systemPrompt = `Tu es un assistant de maintenance industrielle intégré dans une application SAP PM.
Tu réponds en français, de façon concise et claire.
Tu analyses uniquement les données fournies dans le contexte ci-dessous pour répondre aux questions.
Si une information n'est pas disponible dans le contexte, dis-le clairement.
Ne donne jamais de conseils médicaux, juridiques ou financiers.

=== DONNÉES DISPONIBLES ===
${context || 'Aucune donnée disponible.'}
=== FIN DES DONNÉES ===`;

    const body = {
      contents: [
        { role: 'user', parts: [{ text: systemPrompt + '\n\nQuestion : ' + question }] }
      ],
      generationConfig: {
        temperature: 0.2,
        maxOutputTokens: 1024,
      }
    };

    const response = await fetch(
      `https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-flash:generateContent?key=${GEMINI_API_KEY}`,
      { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(body) }
    );

    if (!response.ok) {
      const err = await response.text();
      res.status(response.status).json({ error: 'Erreur Gemini: ' + err });
      return;
    }

    const data = await response.json();
    const answer = data?.candidates?.[0]?.content?.parts?.[0]?.text || 'Pas de réponse.';
    res.status(200).json({ answer });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
};
