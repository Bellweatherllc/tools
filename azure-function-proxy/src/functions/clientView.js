const { app } = require('@azure/functions');

// Client selections link: a client opens https://bellweatherllc.github.io/tools/selections/client.html#k=<code>
// with no Microsoft sign-in. The page sends the code here; this function checks it and returns that ONE job's
// selections, read with this function's own app-only Graph access (Sites.Selected, read, BWCore only).
//
// Security:
// - The code is checked HERE, on the server. CORS is not access control.
// - A code maps to exactly one job. Codes are rows in CORE_SelectionDesignDocs (Job '_links', Collection 'clientLinks',
//   DocId = the code, Data {job, active}). Turning a link off in the Design screen deletes the row; the code stops working at once.
// - Only what the client view shows leaves here: selections, choices, allowances, project settings, mood boards,
//   confirmed meeting recaps, picture thumbnails (Kind Image only). Never internal notes, builder costs, quotes,
//   schedules, other jobs, or anything on the shared library.
// - Graph thumbnail links expire (about an hour); they are fetched fresh on every request and never stored.
//
// App settings (Function App > Configuration), never in code:
//   CLIENT_VIEW_TENANT_ID, CLIENT_VIEW_CLIENT_ID, CLIENT_VIEW_CLIENT_SECRET

const GRAPH = 'https://graph.microsoft.com/v1.0';
const LISTS = GRAPH + '/sites/bellweather.sharepoint.com:/sites/BWCore:/lists';
const SEL = {
  S: 'Title,Job,Category,Location,SelectionTotalCents,AllowanceLookupId,DesignKey,DesignStatus,DesignData',
  C: 'Title,Job,SelectionLookupId,Description,PriceCents,ChoiceStatus,DesignKey,DesignData',
  A: 'Title,Job,AllowanceCents',
  D: 'Title,Job,Collection,DocId,Data',
};
const CODE = /^[A-Za-z0-9_-]{32,64}$/;
const CLIENT_DOCS = new Set(['portal', 'boards', 'meetings']);
const TEAM_ONLY_ITEM = ['followUp', 'basis', 'calc'];          // never shown to clients (team notes, follow-ups, quantity calculators)
const TEAM_ONLY_SETTINGS = ['pipelineUrl', 'targetBudgetCents'];

let _tok = null;
async function appToken() {
  if (_tok && _tok.exp > Date.now() + 60000) return _tok.value;
  const t = process.env.CLIENT_VIEW_TENANT_ID, id = process.env.CLIENT_VIEW_CLIENT_ID, secret = process.env.CLIENT_VIEW_CLIENT_SECRET;
  if (!t || !id || !secret) throw Object.assign(new Error('not configured'), { status: 503 });
  const r = await fetch(`https://login.microsoftonline.com/${encodeURIComponent(t)}/oauth2/v2.0/token`, {
    method: 'POST', headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
    body: new URLSearchParams({ client_id: id, client_secret: secret, grant_type: 'client_credentials', scope: 'https://graph.microsoft.com/.default' }),
  });
  if (!r.ok) throw Object.assign(new Error('token ' + r.status), { status: 502 });
  const j = await r.json();
  _tok = { value: j.access_token, exp: Date.now() + (Number(j.expires_in) || 3000) * 1000 };
  return _tok.value;
}
async function g(url) {
  const r = await fetch(url, { headers: { Authorization: 'Bearer ' + await appToken() } });
  if (!r.ok) throw Object.assign(new Error('graph ' + r.status), { status: r.status === 404 ? 404 : 502 });
  return r.json();
}
async function gAll(url) { let out = [], next = url, n = 0; while (next && n++ < 20) { const j = await g(next); out = out.concat(j.value || []); next = j['@odata.nextLink'] || null; } return out; }
const enc = v => encodeURIComponent(String(v).replace(/'/g, "''"));
const byJob = (list, job, sel) => gAll(`${LISTS}/${list}/items?$filter=fields/Job eq '${enc(job)}'&$expand=fields($select=${sel})&$top=999`);
const parseJ = s => { try { const v = JSON.parse(s || '{}'); return v && typeof v === 'object' ? v : {}; } catch (e) { return {}; } };
const strip = (json, keys) => { const o = parseJ(json); for (const k of keys) delete o[k]; return JSON.stringify(o); };
async function pool(items, n, fn) { const out = []; let i = 0; await Promise.all(Array.from({ length: Math.min(n, items.length) }, async () => { while (i < items.length) { const k = i++; out[k] = await fn(items[k]); } })); return out; }

app.http('clientView', {
  methods: ['POST'],
  authLevel: 'anonymous', // the link code in the body is the access check (below); CORS only limits which pages may call
  handler: async (request, context) => {
    const headers = { 'Cache-Control': 'no-store', 'X-Robots-Tag': 'noindex, nofollow' };
    let body; try { body = await request.json(); } catch (e) { return { status: 400, headers, jsonBody: { error: 'bad_request' } }; }
    const code = String((body && body.k) || '');
    if (!CODE.test(code)) return { status: 404, headers, jsonBody: { error: 'not_found' } };
    try {
      const links = await byJob('CORE_SelectionDesignDocs', '_links', SEL.D);
      const row = links.find(r => r.fields && r.fields.Collection === 'clientLinks' && r.fields.DocId === code);
      const link = row && parseJ(row.fields.Data);
      if (!link || !link.job || link.active === false) return { status: 404, headers, jsonBody: { error: 'not_found' } };
      const job = String(link.job);
      const [S, C, A, D, P, F] = await Promise.all([
        byJob('CORE_Selections', job, SEL.S), byJob('CORE_SelectionChoices', job, SEL.C), byJob('CORE_SelectionAllowances', job, SEL.A),
        byJob('CORE_SelectionDesignDocs', job, SEL.D),
        gAll(`${LISTS}/CORE_Projects/items?$expand=fields($select=Title,JobCode,ProjectName,GanttData)&$top=999`).catch(() => []),
        gAll(`${LISTS}/CORE_SelectionFiles/items?$filter=fields/Job eq '${enc(job)}'&$expand=fields($select=Job,SelectionLookupId,ChoiceLookupId,Kind,FileLeafRef),driveItem($select=id,name)&$top=999`).catch(() => []),
      ]);
      const sels = S.filter(s => s.fields.DesignStatus !== 'removed').map(s => ({ id: s.id, fields: { ...s.fields, DesignData: strip(s.fields.DesignData, TEAM_ONLY_ITEM) } }));
      const docs = D.filter(d => CLIENT_DOCS.has(d.fields.Collection)).filter(d => d.fields.Collection !== 'meetings' || parseJ(d.fields.Data).confirmed)
        .map(d => ({ id: d.id, fields: d.fields.Collection === 'portal' ? { ...d.fields, Data: strip(d.fields.Data, TEAM_ONLY_SETTINGS) } : d.fields }));
      const pr = P.find(r => String(r.fields.JobCode || '').trim().toUpperCase() === job.trim().toUpperCase());
      const project = pr ? { ProjectName: pr.fields.ProjectName || '', GanttData: pr.fields.GanttData || '' } : null;
      const pics = F.filter(x => x.fields && (x.fields.Kind === 'Image' || (!x.fields.Kind && /\.(jpe?g|png|gif|webp|heic)$/i.test(x.fields.FileLeafRef || '')))).slice(0, 300);
      const files = (await pool(pics, 8, async x => {
        try { const t = await g(`${LISTS}/CORE_SelectionFiles/items/${x.id}/driveItem/thumbnails?$select=medium,large`); const s = t.value && t.value[0];
          return { id: x.id, sel: String(x.fields.SelectionLookupId || ''), choice: x.fields.ChoiceLookupId != null ? String(x.fields.ChoiceLookupId) : '', kind: 'Image', name: (x.driveItem && x.driveItem.name) || x.fields.FileLeafRef || '', thumb: (s && ((s.large && s.large.url) || (s.medium && s.medium.url))) || '' };
        } catch (e) { return null; } })).filter(Boolean);
      context.log(`clientView: ${job} opened`);
      return { status: 200, headers, jsonBody: { job, S: sels, C, A, D: docs, project, files } };
    } catch (e) {
      context.error('clientView failed:', e.message);
      return { status: e.status === 503 ? 503 : 502, headers, jsonBody: { error: e.status === 503 ? 'not_configured' : 'unavailable' } };
    }
  },
});
