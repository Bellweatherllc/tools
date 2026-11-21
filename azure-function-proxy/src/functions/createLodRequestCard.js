const { app } = require('@azure/functions');

// This function is the only thing that ever sees the real Trello API key/token
// for this feature — they live in Azure Function App Settings (the same
// TRELLO_API_KEY / TRELLO_TOKEN createTrelloBoard.js already uses), never in
// the browser tool or in source control. The client sends a board name, a
// list name, a card name/description, and a list of Trello usernames; this
// finds that board+list by name, creates the card, and adds each person to
// the board (so they can actually be added to the card — Trello won't let you
// add a card member who isn't already a board member) and then to the card
// itself, server-side.

app.http('createLodRequestCard', {
  methods: ['POST'],
  authLevel: 'anonymous', // access is limited via CORS (Function App > CORS) to the GitHub Pages origin
  handler: async (request, context) => {
    let body;
    try {
      body = await request.json();
    } catch (e) {
      return { status: 400, jsonBody: { error: 'Invalid JSON body' } };
    }

    const boardName = (body.boardName || '').toString().trim();
    const listName = (body.listName || '').toString().trim();
    const cardName = (body.cardName || '').toString().trim().slice(0, 200);
    const description = (body.description || '').toString().slice(0, 16000);
    const usernames = Array.isArray(body.usernames)
      ? body.usernames.map(u => (u || '').toString().trim().replace(/^@/, '')).filter(Boolean).slice(0, 50)
      : [];
    if (!boardName) return { status: 400, jsonBody: { error: 'boardName required' } };
    if (!listName) return { status: 400, jsonBody: { error: 'listName required' } };
    if (!cardName) return { status: 400, jsonBody: { error: 'cardName required' } };

    const key = process.env.TRELLO_API_KEY;
    const token = process.env.TRELLO_TOKEN;
    if (!key || !token) {
      context.error('Trello app settings missing (TRELLO_API_KEY / TRELLO_TOKEN).');
      return { status: 500, jsonBody: { error: 'Server not configured' } };
    }
    const qs = (params) => new URLSearchParams({ ...params, key, token }).toString();
    const norm = (s) => (s || '').trim().toLowerCase();

    try {
      // Find the board by exact name (case-insensitive) among boards this token
      // can see — not a fuzzy/partial search, since the two real targets (the
      // Sales Big Board, and whatever test board is in use) are known, fixed names.
      const boardsRes = await fetch(`https://api.trello.com/1/members/me/boards?${qs({ fields: 'name,id' })}`);
      if (!boardsRes.ok) throw new Error('Board lookup failed: ' + boardsRes.status);
      const boards = await boardsRes.json();
      const board = boards.find(b => norm(b.name) === norm(boardName));
      if (!board) return { status: 404, jsonBody: { error: `No board named "${boardName}" found` } };

      const listsRes = await fetch(`https://api.trello.com/1/boards/${board.id}/lists?${qs({ fields: 'name,id' })}`);
      if (!listsRes.ok) throw new Error('List lookup failed: ' + listsRes.status);
      const lists = await listsRes.json();
      const list = lists.find(l => norm(l.name) === norm(listName));
      if (!list) return { status: 404, jsonBody: { error: `No list named "${listName}" found on board "${boardName}"` } };

      // Resolve each username to a member id and make sure they're a board
      // member first — Trello rejects adding a card member who isn't already
      // on the board.
      const memberIds = [];
      const failed = [];
      for (const username of usernames) {
        try {
          const mRes = await fetch(`https://api.trello.com/1/members/${encodeURIComponent(username)}?${qs({ fields: 'id' })}`);
          if (!mRes.ok) throw new Error('unknown Trello user');
          const member = await mRes.json();
          const aRes = await fetch(`https://api.trello.com/1/boards/${board.id}/members/${member.id}?${qs({ type: 'normal' })}`, { method: 'PUT' });
          if (!aRes.ok) throw new Error('board add failed');
          memberIds.push(member.id);
        } catch (e) {
          failed.push(username);
        }
      }

      const cardParams = { idList: list.id, name: cardName, desc: description };
      if (memberIds.length) cardParams.idMembers = memberIds.join(',');
      const cRes = await fetch(`https://api.trello.com/1/cards?${qs(cardParams)}`, { method: 'POST' });
      if (!cRes.ok) throw new Error('Card create failed: ' + cRes.status + ' ' + await cRes.text());
      const card = await cRes.json();

      return { status: 200, jsonBody: { card: { id: card.id, url: card.url, name: card.name }, failed } };
    } catch (e) {
      context.error('Trello proxy failure:', e);
      return { status: 502, jsonBody: { error: e.message || 'Trello proxy failure' } };
    }
  },
});
