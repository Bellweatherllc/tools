/**
 * Sherpa Lookup — Programmatic SharePoint file lookup for CORE tools
 *
 * Deploy to: chicophilly.github.io/bellweather-tools/lib/sherpa-lookup.js
 *
 * Given a project name and a source type, returns the best-matching file.
 * Handles fuzzy project matching, file-type filtering, archive exclusion,
 * and confidence scoring so callers know when to show an alert.
 *
 * USAGE:
 *   <script src="https://alcdn.msauth.net/browser/2.38.0/js/msal-browser.min.js"></script>
 *   <script src="https://chicophilly.github.io/bellweather-tools/lib/sherpa-auth.js"></script>
 *   <script src="https://chicophilly.github.io/bellweather-tools/lib/sherpa-lookup.js"></script>
 *
 *   await SherpaAuth.init();
 *   if (!SherpaAuth.isSignedIn()) await SherpaAuth.signIn();
 *
 *   const result = await SherpaLookup.find('Lebow', 'scope_outline');
 *   if (result.status === 'confident') {
 *     attachFile(result.file);
 *   } else if (result.status === 'ambiguous') {
 *     attachFile(result.file);
 *     showAlert('Multiple candidates found — review', result.candidates);
 *   } else {
 *     showNotFound();
 *   }
 *
 * API:
 *   SherpaLookup.find(projectName, sourceType, options)
 *     → { status, file, candidates, query }
 *
 *   SherpaLookup.SOURCE_TYPES — list of supported source type IDs
 *
 * Return shape:
 *   status:     'confident' | 'ambiguous' | 'none'
 *   file:       null | { name, size, modified, webUrl, folderUrl, path, ext }
 *   candidates: array of the same file shape (always populated when hits were found)
 *   query:      the search query string used (for debugging)
 */

(function(global) {
  'use strict';

  // ─────────────────────────────────────────────────────────────────────────
  // CONFIG
  // ─────────────────────────────────────────────────────────────────────────
  const CUSTOMERS_DRIVE_HOST = 'bellweather.sharepoint.com';
  const CUSTOMERS_SITE_PATH  = '/sites/BWC';
  const CUSTOMERS_LIBRARY    = 'Customers';

  // Folders whose contents should never be returned by programmatic lookup.
  // See SharePoint_Sherpa_Structure_and_Routing.md §8.
  const EXCLUDE_RE = /\/(Archive|ARCHIVE|Z Archive|SAMPLE Scope Outlines)(\/|$)/i;

  // ─────────────────────────────────────────────────────────────────────────
  // SOURCE TYPE DEFINITIONS
  // Each source type declares:
  //   - keywords: words that should appear in the filename (any one matches)
  //   - extensions: acceptable file extensions (empty = any)
  //   - folderHint: expected folder path fragment to prefer when ranking
  //     (not strict — used to boost, not filter)
  // ─────────────────────────────────────────────────────────────────────────
  const SOURCE_TYPES = {
    scope_outline: {
      label: 'Scope Outline',
      keywords: ['scope outline', 'scope'],
      extensions: ['docx', 'pdf'],
      folderHint: '1. DAs',
      folderHintAlt: '1. IDEA & FDDA',
    },
    design_agreement: {
      label: 'Design Agreement',
      keywords: ['design agreement', 'design & development agreement', 'DA', 'IDEA', 'FDDA'],
      extensions: ['docx', 'pdf'],
      folderHint: '0. Signed Contract Documents',
      folderHintAlt: '1. DAs',
    },
    construction_agreement: {
      label: 'Construction Agreement',
      keywords: ['construction agreement', 'CA'],
      extensions: ['docx', 'pdf'],
      folderHint: '0. Signed Contract Documents',
      folderHintAlt: '3. Contract & CO\'s',
    },
    selections: {
      label: 'BuilderTrend Selections',
      keywords: ['selections', 'selection'],
      extensions: ['pdf', 'xlsx', 'xls'],
      folderHint: '4. Selection Material RFPs',
    },
    estimate: {
      label: 'Estimate',
      keywords: ['estimate', 'estimating'],
      extensions: ['xlsx', 'xls'],
      folderHint: '8. Estimating',
    },
    plans: {
      label: 'Plans',
      keywords: ['plans', 'drawings'],
      extensions: ['pdf'],
      folderHint: '3. Plans',
      folderHintAlt: '3. Plans, Engineering',
    },
    permit: {
      label: 'Permit',
      keywords: ['permit'],
      extensions: ['pdf'],
      folderHint: '1. Permits',
      folderHintAlt: '7. Permits, Inspections, Fees',
    },
  };

  // ─────────────────────────────────────────────────────────────────────────
  // STATE — driveId is resolved once and cached per page load
  // ─────────────────────────────────────────────────────────────────────────
  let cachedDriveId = null;

  async function getDriveId() {
    if (cachedDriveId) return cachedDriveId;
    const url = `${SherpaAuth.GRAPH_BASE}/sites/${CUSTOMERS_DRIVE_HOST}:${CUSTOMERS_SITE_PATH}:/drives`;
    const data = await SherpaAuth.graph(url);
    const drive = data.value.find(d => d.name === CUSTOMERS_LIBRARY);
    if (!drive) throw new Error(`Sherpa: drive "${CUSTOMERS_LIBRARY}" not found`);
    cachedDriveId = drive.id;
    return cachedDriveId;
  }

  // ─────────────────────────────────────────────────────────────────────────
  // SEARCH + ENRICH — core logic shared with Sherpa's UI
  // ─────────────────────────────────────────────────────────────────────────
  async function searchAndEnrich(query) {
    const driveId = await getDriveId();
    const encoded = query.replace(/'/g, "''");
    const url = `${SherpaAuth.GRAPH_BASE}/drives/${driveId}/root/search(q='${encoded}')` +
      `?$top=50&$select=id,name,parentReference,lastModifiedDateTime,file,webUrl,size`;

    const searchResult = await SherpaAuth.graph(url);
    let hits = (searchResult.value || []).filter(i => i.file);

    // First-pass archive filter (catches hits where parentReference.path is populated)
    hits = hits.filter(i => {
      const p1 = i.parentReference?.path || '';
      const p2 = i.webUrl || '';
      return !EXCLUDE_RE.test(p1) && !EXCLUDE_RE.test(p2);
    });

    // Parallel parent-enrichment (needed for Office files whose search response has
    // empty parentReference.path — see SharePoint_Sherpa_Structure_and_Routing.md)
    const enriched = await Promise.all(hits.map(async (i) => {
      const parentId = i.parentReference?.id;
      const pDriveId = i.parentReference?.driveId || driveId;
      let parent = null;
      if (parentId) {
        try {
          parent = await SherpaAuth.graph(
            `${SherpaAuth.GRAPH_BASE}/drives/${pDriveId}/items/${parentId}?$select=name,parentReference,webUrl`
          );
        } catch { /* non-fatal */ }
      }
      return { hit: i, parent };
    }));

    // Second-pass archive filter — now reliable
    const filtered = enriched.filter(({ parent }) => {
      const parentPath = parent?.parentReference?.path || '';
      const parentName = parent?.name || '';
      return !EXCLUDE_RE.test(parentPath + '/' + parentName);
    });

    // Build uniform result shape
    return filtered.map(({ hit, parent }) => ({
      id:         hit.id,
      name:       hit.name,
      size:       hit.size || 0,
      modified:   hit.lastModifiedDateTime,
      ext:        extOf(hit.name),
      webUrl:     hit.webUrl || null,
      folderUrl:  parent?.webUrl || null,
      path:       buildParentPath(parent),
      parentName: parent?.name || '',
    }));
  }

  // ─────────────────────────────────────────────────────────────────────────
  // RANKING + SCORING
  // ─────────────────────────────────────────────────────────────────────────

  /**
   * Score a result for how well it matches the query intent.
   * Higher = better match.
   *
   * Scoring components:
   *   +100 if project name appears in filename
   *   + 50 if any source keyword appears in filename
   *   + 30 if file is in the expected folder (folderHint or folderHintAlt)
   *   + 10 if extension matches one of the expected extensions
   *   - 20 if filename contains "DRAFT" (suggests superseded)
   */
  function scoreResult(result, projectName, sourceDef) {
    let score = 0;
    const nameLower   = result.name.toLowerCase();
    const projectLower = projectName.toLowerCase();
    const pathLower   = (result.path || '').toLowerCase();

    if (nameLower.includes(projectLower)) score += 100;

    const keywordHit = sourceDef.keywords.some(kw => nameLower.includes(kw.toLowerCase()));
    if (keywordHit) score += 50;

    const folderHit = sourceDef.folderHint && pathLower.includes(sourceDef.folderHint.toLowerCase());
    const folderHitAlt = sourceDef.folderHintAlt && pathLower.includes(sourceDef.folderHintAlt.toLowerCase());
    if (folderHit || folderHitAlt) score += 30;

    if (sourceDef.extensions.length === 0 || sourceDef.extensions.includes(result.ext)) {
      score += 10;
    } else {
      // Wrong extension is a strong negative signal
      score -= 50;
    }

    if (/\bdraft\b/i.test(result.name)) score -= 20;

    return score;
  }

  /**
   * Assess confidence given the scored, sorted results.
   *
   * Rules:
   *   - If no results meet a minimum threshold (score >= 80): 'none'
   *   - If top score is >= 120 and second-best is more than 40 below: 'confident'
   *   - Otherwise: 'ambiguous'
   *
   * Rationale: score 120 means "has project name + keyword" (100 + 50 - extension mismatch
   * at worst). Gap of 40 means the second candidate is meaningfully weaker.
   */
  function assessConfidence(scored) {
    if (scored.length === 0) return 'none';
    const top = scored[0].score;
    if (top < 80) return 'none';
    if (scored.length === 1) return 'confident';
    const second = scored[1].score;
    if (top >= 120 && (top - second) >= 40) return 'confident';
    return 'ambiguous';
  }

  // ─────────────────────────────────────────────────────────────────────────
  // PUBLIC: find()
  // ─────────────────────────────────────────────────────────────────────────

  /**
   * Find the best-matching file for a project + source type.
   *
   * @param {string} projectName — last name or project identifier fragment (fuzzy match)
   * @param {string} sourceType — one of SOURCE_TYPES keys
   * @param {object} [options]
   * @param {number} [options.maxCandidates=5] — how many candidates to return in the `candidates` array
   * @returns {Promise<{status, file, candidates, query}>}
   */
  async function find(projectName, sourceType, options = {}) {
    const maxCandidates = options.maxCandidates || 5;
    const sourceDef = SOURCE_TYPES[sourceType];
    if (!sourceDef) {
      throw new Error(`SherpaLookup: unknown source type "${sourceType}". ` +
        `Valid: ${Object.keys(SOURCE_TYPES).join(', ')}`);
    }
    if (!projectName || !projectName.trim()) {
      throw new Error('SherpaLookup: projectName is required');
    }

    // Build search query: project name + primary keyword
    // Graph's search is a KQL-like ranked search, so putting both terms in the query
    // returns files matching either, with ranking preferring ones that match both.
    const query = `${projectName} ${sourceDef.keywords[0]}`.trim();

    // Run search, enrich with parent folder details
    let results;
    try {
      results = await searchAndEnrich(query);
    } catch (e) {
      console.error('SherpaLookup: search failed', e);
      return { status: 'none', file: null, candidates: [], query, error: e.message };
    }

    // Score and sort
    const scored = results.map(r => ({ ...r, score: scoreResult(r, projectName, sourceDef) }));
    scored.sort((a, b) => {
      // Primary: score descending
      if (b.score !== a.score) return b.score - a.score;
      // Tiebreaker: newest first
      return new Date(b.modified) - new Date(a.modified);
    });

    const status = assessConfidence(scored);
    const file = status === 'none' ? null : scored[0];
    const candidates = scored.slice(0, maxCandidates);

    return { status, file, candidates, query };
  }

  // ─────────────────────────────────────────────────────────────────────────
  // HELPERS — mirror the ones in Sherpa's main UI
  // ─────────────────────────────────────────────────────────────────────────

  function extOf(name) {
    const m = name.match(/\.([^.]+)$/);
    return m ? m[1].toLowerCase() : '';
  }

  function buildParentPath(parent) {
    if (!parent) return '';
    const raw = parent.parentReference?.path || '';
    let lineage = '';
    const idx = raw.indexOf('root:/');
    if (idx >= 0) {
      let p = raw.slice(idx + 6);
      if (p.startsWith('/')) p = p.slice(1);
      if (p.startsWith('Customers/')) p = p.slice(10);
      else if (p === 'Customers') p = '';
      lineage = decodeURIComponent(p);
    }
    const parts = [];
    if (lineage) parts.push(...lineage.split('/'));
    if (parent.name) parts.push(parent.name);
    return parts.join(' › ');
  }

  // ─────────────────────────────────────────────────────────────────────────
  // EXPOSE
  // ─────────────────────────────────────────────────────────────────────────
  global.SherpaLookup = {
    find,
    SOURCE_TYPES: Object.keys(SOURCE_TYPES),
    _SOURCE_DEFS: SOURCE_TYPES,  // exposed for debugging; do not depend on shape
  };

})(window);
