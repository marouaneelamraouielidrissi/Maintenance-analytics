// Vercel serverless function — Groq chatbot (Llama 3)
module.exports = async function handler(req, res) {
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');
  if (req.method === 'OPTIONS') { res.status(204).end(); return; }
  if (req.method !== 'POST')    { res.status(405).json({ error: 'Method not allowed' }); return; }

  const GROQ_API_KEY = process.env.GROQ_API_KEY;
  if (!GROQ_API_KEY) { res.status(500).json({ error: 'Clé API manquante' }); return; }

  try {
    const { question, context } = req.body || {};
    if (!question) { res.status(400).json({ error: 'Question manquante' }); return; }

    const systemPrompt = `Tu es un assistant de maintenance industrielle intégré dans une application SAP PM.
Tu réponds en français, de façon concise et claire.
Tu analyses uniquement les données fournies dans le contexte ci-dessous pour répondre aux questions.
Si une information n'est pas disponible dans le contexte, dis-le clairement.
Ne donne jamais de conseils médicaux, juridiques ou financiers.

Correspondance des postes de travail (corps de métier) :
- 423-ELEC  = Service Électrique
- 423-REG   = Régulation
- 421-MEC   = Mécanique
- 421-CHAU  = Mécanique (chaudronnerie)
- 421-INST  = Installation

=== DONNÉES DISPONIBLES ===
${context || 'Aucune donnée disponible.'}
=== FIN DES DONNÉES ===`;

    const body = {
      model: 'llama-3.3-70b-versatile',
      messages: [
        { role: 'system', content: systemPrompt },
        { role: 'user',   content: question }
      ],
      temperature: 0.2,
      max_tokens: 1024,
    };

    const response = await fetch('https://api.groq.com/openai/v1/chat/completions', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'Authorization': `Bearer ${GROQ_API_KEY}`
      },
      body: JSON.stringify(body)
    });

    if (!response.ok) {
      const err = await response.text();
      res.status(response.status).json({ error: 'Erreur Groq: ' + err });
      return;
    }

    const data = await response.json();
    const answer = data?.choices?.[0]?.message?.content || 'Pas de réponse.';
    res.status(200).json({ answer });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
};
