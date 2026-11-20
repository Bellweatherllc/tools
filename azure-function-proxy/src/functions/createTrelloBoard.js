const { app } = require('@azure/functions');

// This function is the only thing that ever sees the real Trello API key/token.
// They live in Azure Function App Settings (environment variables), never in the
// browser tool or in source control. The client sends a board name and a list of
// Trello usernames; this copies the project template into a new board and adds
// each person to it, server-side.

let _cachedTemplate = null; // {boardId,orgId} — resolved once per warm instance

app.http('createTrelloBoard', {
  methods: ['POST'],
  authLevel: 'anonymous', // access is limited via CORS (Function App > CORS) to the GitHub Pages origin
  handler: async (request, context) => {
    let body;
    try {
      body = await request.json();
    } catch (e) {
      return { status: 400, jsonBody: { error: 'Invalid JSON body' } };
    }

    const name = (body.name || '').toString().trim().slice(0, 200);
    const usernames = Array.isArray(body.usernames)
      ? body.usernames.map(u => (u || '').toString().trim().replace(/^@/, '')).filter(Boolean).slice(0, 50)
      : [];
    if (!name) return { status: 400, jsonBody: { error: 'Board name required' } };

    const key = process.env.TRELLO_API_KEY;
    const token = process.env.TRELLO_TOKEN;
    const templateShortlink = process.env.TRELLO_TEMPLATE_SHORTLINK;
    if (!key || !token || !templateShortlink) {
      context.error('Trello app settings missing (TRELLO_API_KEY / TRELLO_TOKEN / TRELLO_TEMPLATE_SHORTLINK).');
      return { status: 500, jsonBody: { error: 'Server not configured' } };
    }
    const qs = (params) => new URLSearchParams({ ...params, key, token }).toString();

    try {
      if (!_cachedTemplate) {
        const tRes = await fetch(`https://api.trello.com/1/boards/${templateShortlink}?${qs({ fields: 'id,idOrganization' })}`);
        if (!tRes.ok) throw new Error('Template board lookup failed: ' + tRes.status);
        const t = await tRes.json();
        _cachedTemplate = { boardId: t.id, orgId: t.idOrganization || '' };
      }

      const bRes = await fetch(`https://api.trello.com/1/boards/?${qs({
        name,
        idBoardSource: _cachedTemplate.boardId,
        idOrganization: _cachedTemplate.orgId,
        keepFromSource: 'cards',
      })}`, { method: 'POST' });
      if (!bRes.ok) throw new Error('Board create failed: ' + bRes.status + ' ' + await bRes.text());
      const board = await bRes.json();

      const failed = [];
      for (const username of usernames) {
        try {
          const mRes = await fetch(`https://api.trello.com/1/members/${encodeURIComponent(username)}?${qs({ fields: 'id' })}`);
          if (!mRes.ok) throw new Error('unknown Trello user');
          const member = await mRes.json();
          const aRes = await fetch(`https://api.trello.com/1/boards/${board.id}/members/${member.id}?${qs({ type: 'normal' })}`, { method: 'PUT' });
          if (!aRes.ok) throw new Error('add failed');
        } catch (e) {
          failed.push(username);
        }
      }

      return { status: 200, jsonBody: { board: { id: board.id, url: board.url, name: board.name }, failed } };
    } catch (e) {
      context.error('Trello proxy failure:', e);
      return { status: 502, jsonBody: { error: e.message || 'Trello proxy failure' } };
    }
  },
});
