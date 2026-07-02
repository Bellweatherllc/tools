const { app } = require('@azure/functions');

// This function is the only thing that ever sees the real Anthropic API key.
// The key lives in an Azure Function App Setting (environment variable),
// never in the browser tool or in source control.

app.http('generateOpening', {
  methods: ['POST'],
  authLevel: 'anonymous', // access is limited via CORS (Function App > CORS) to the GitHub Pages origin
  handler: async (request, context) => {
    let body;
    try {
      body = await request.json();
    } catch (e) {
      return { status: 400, jsonBody: { error: 'Invalid JSON body' } };
    }

    const clientName = (body.clientName || 'the homeowners').toString().slice(0, 200);
    const scope = (body.scope || '').toString().slice(0, 6000);

    const apiKey = process.env.ANTHROPIC_API_KEY;
    if (!apiKey) {
      context.error('ANTHROPIC_API_KEY app setting is missing.');
      return { status: 500, jsonBody: { error: 'Server not configured' } };
    }

    const prompt = `You are writing a warm, professional opening paragraph for a residential design-build firm called Bellweather. The letter is addressed to a client named ${clientName}. Based on the project scope below, write 2–3 sentences that feel like a warm letter introduction — excited about the project, personal, not salesy. Do not use the word "delighted". Do not include a greeting or sign-off. Just the opening sentences.\n\nScope:\n${scope}`;

    try {
      const resp = await fetch('https://api.anthropic.com/v1/messages', {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
          'x-api-key': apiKey,
          'anthropic-version': '2023-06-01'
        },
        body: JSON.stringify({
          model: 'claude-sonnet-4-20250514',
          max_tokens: 200,
          messages: [{ role: 'user', content: prompt }]
        })
      });

      if (!resp.ok) {
        const errText = await resp.text();
        context.error('Anthropic API error:', resp.status, errText);
        return { status: 502, jsonBody: { error: 'Upstream API error' } };
      }

      const data = await resp.json();
      const opening = data.content?.[0]?.text || '';
      return { status: 200, jsonBody: { opening } };

    } catch (e) {
      context.error('Proxy failure:', e);
      return { status: 500, jsonBody: { error: 'Proxy failure' } };
    }
  }
});
