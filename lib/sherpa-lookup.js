/**
 * Sherpa Lookup — Programmatic SharePoint file lookup for CORE tools
 *
 * Deploy to: bellweatherllc.github.io/tools/lib/sherpa-lookup.js
 *
 * Given a project name and a source type, returns the best-matching file.
 * Format preferences are advisory — the library ranks preferred formats higher
 * but never filters out other formats. The calling tool decides what to accept
 * and how to communicate format quality to the user.
 *
 * USAGE:
 *   <script src="https://alcdn.msauth.net/browser/2.38.0/js/msal-browser.min.js"></script>
 *   <script src="https://bellweatherllc.github.io/tools/lib/sherpa-auth.js"></script>
 *   <script src="https://bellweatherllc.github.io/tools/lib/sherpa-lookup.js"></script>
 *
 *   await SherpaAuth.init();
 *   if (!SherpaAuth.isSignedIn()) await SherpaAuth.signIn();
 *
 *   const result = await SherpaLookup.find('Lebow', 'scope_outline');
 *
 *   if (result.status === 'confident') {
 *     attachFile(result.file);
 *     if (result.file.formatHint === 'acceptable') showFormatNote(result.file);
 *   } else if (result.status === 'ambiguous') {
 *     attachFile(result.file);
 *     showAlert('Multiple candidates — review', result.candidates);
 *   } else {
 *     showNotFound();
 *   }
 *
 * API:
 *   SherpaLookup.find(projectName, sourceType, options)
 *     → { status, file, candidates, query }
 *
 *   SherpaLookup.SOURCE_TYPES  — array of valid source type strings
 *   SherpaLookup._SOURCE_DEFS  — full source type definitions (for debugging)
 *
 * Result file shape:
 *   { name, size, modified, ext, webUrl, folderUrl, path, parentName, score, formatHint }
 *
 * formatHint values (set by library based on source type preferences):
 *   'preferred'   — ideal format for this source type (e.g. docx for scope outline)
 *   'acceptable'  — usable but not ideal (e.g. pdf for scope outline)
 *   'other'       — unexpected format; caller should decide whether to accept
 *
 * status values:
 *   'confident'   — one clear best match; attach silently
 *   'ambiguous'   — multiple similar candidates; surface for user review
 *   'none'        — nothing found meeting minimum relevance threshold
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
  //
  // keywords:            words to match in filename (any one = keyword hit)
  // preferredExtensions: ideal formats — listed in preference order (first = most preferred).
  //                      Used for scoring and formatHint; NOT used as a hard filter.
  // acceptableExtensions: usable but not ideal formats. Scored lower than preferred.
  //                      Files in neither list get formatHint 'other' and lowest extension score.
  // folderHint:          expected folder name fragment — boosts score when matched.
  //                      Not a filter — just a ranking signal.
  // ─────────────────────────────────────────────────────────────────────────
  const SOURCE_TYPES = {
    scope_outline: {
      label: 'Scope Outline',
      keywords: ['scope outline', 'scope'],
      preferredExtensions:   ['docx'],
      acceptableExtensions:  ['pdf'],
      folderHint:    '1. DAs',
      folderHintAlt: '1. IDEA & FDDA',
    },
    design_agreement: {
      label: 'Design Agreement',
      keywords: ['design agreement', 'design & development agreement', 'DA', 'IDEA', 'FDDA'],
      preferredExtensions:   ['docx'],
      acceptableExtensions:  ['pdf'],
      folderHint:    '0. Signed Contract Documents',
      folderHintAlt: '1. DAs',
    },
    construction_agreement: {
      label: 'Construction Agreement',
      keywords: ['construction agreement', 'CA'],
      preferredExtensions:   ['docx'],
      acceptableExtensions:  ['pdf'],
      folderHint:    '0. Signed Contract Documents',
      folderHintAlt: "3. Contract & CO's",
    },
    selections: {
      label: 'BuilderTrend Selections',
      keywords: ['selections', 'selection'],
      preferredExtensions:   ['xlsx'],
      acceptableExtensions:  ['xls', 'pdf'],
      folderHint:    '4. Selection Material RFPs',
    },
    estimate: {
      label: 'Estimate',
      keywords: ['estimate', 'estimating'],
      preferredExtensions:   ['xlsx'],
      acceptableExtensions:  ['xls'],
      folderHint:    '8. Estimating',
    },
    plans: {
      label: 'Plan Notes',
      keywords: ['plan notes', 'plans'],
      preferredExtensions:   ['txt', 'json'],
      acceptableExtensions:  [],
      folderHint:    '3. Plans',
      folderHintAlt: '3. Plans, Engineering',
    },
    permit: {
      label: 'Permit',
      keywords: ['permit'],
      preferredExtensions:   ['pdf'],
      acceptableExtensions:  [],
      folderHint:    '1. Permits',
      folderHintAlt: '7. Permits, Inspections, Fees',
    },
  };

  // ─────────────────────────────────────────────────────────────────────────
  // STATE — driveId resolved once and cached per page load
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
  // SEARCH + ENRICH
  //
  // Searches the Customers drive, enriches each hit with parent folder details
  // (needed because Graph strips parentReference.path for Office files).
  // Returns all file results — no extension filtering here. Callers filter.
  // ─────────────────────────────────────────────────────────────────────────
  async function searchAndEnrich(query) {
    const driveId = await getDriveId();

    const encoded = query.replace(/'/g, "''");
    const url = `${SherpaAuth.GRAPH_BASE}/drives/${driveId}/root/search(q='${encoded}')` +
      `?$top=50&$select=id,name,parentReference,lastModifiedDateTime,file,webUrl,size`;

    const searchResult = await SherpaAuth.graph(url);
    let hits = (searchResult.value || []).filter(i => i.file);

    // First-pass archive filter (path populated for PDFs; empty for Office files)
    hits = hits.filter(i => {
      const p1 = i.parentReference?.path || '';
      const p2 = i.webUrl || '';
      return !EXCLUDE_RE.test(p1) && !EXCLUDE_RE.test(p2);
    });

    // Parallel parent-enrichment — fetches real folder path + webUrl for every hit.
    // Necessary because Graph returns a stripped parentReference for Office files
    // (docx, xlsx, pptx) — only driveId and parentId are present, no path field.
    const enriched = await Promise.all(hits.map(async (i) => {
      const parentId = i.parentReference?.id;
      const pDriveId = i.parentReference?.driveId || driveId;
      let parent = null;
      if (parentId) {
        try {
          parent = await SherpaAuth.graph(
            `${SherpaAuth.GRAPH_BASE}/drives/${pDriveId}/items/${parentId}` +
            `?$select=name,parentReference,webUrl`
          );
        } catch { /* non-fatal — result keeps null parent */ }
      }
      return { hit: i, parent };
    }));

    // Second-pass archive filter — reliable now that parent is fetched
    const filtered = enriched.filter(({ parent }) => {
      const parentPath = parent?.parentReference?.path || '';
      const parentName = parent?.name || '';
      return !EXCLUDE_RE.test(parentPath + '/' + parentName);
    });

    // Build uniform result shape
    return filtered.map(({ hit, parent }) => ({
      id:         hit.id,
      driveId:    hit.parentReference?.driveId || driveId,
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
  // FORMAT HINT
  //
  // Categorises a file's extension relative to the source type's preferences.
  // The CA Sources page (or any other caller) uses this to show the user
  // context about format quality without making the decision for them.
  // ─────────────────────────────────────────────────────────────────────────
  function formatHintFor(ext, sourceDef) {
    if (sourceDef.preferredExtensions.includes(ext))  return 'preferred';
    if (sourceDef.acceptableExtensions.includes(ext)) return 'acceptable';
    return 'other';
  }

  // ─────────────────────────────────────────────────────────────────────────
  // SCORING
  //
  // Higher score = better match. Components:
  //   +100  project name appears in filename
  //   + 50  source keyword appears in filename
  //   + 30  file is in the expected folder (folderHint or folderHintAlt)
  //   + 20  preferred extension (first in preferredExtensions list)
  //   + 15  preferred extension (subsequent positions)
  //   + 10  acceptable extension
  //   +  0  other extension (present in results but not rewarded)
  //   -  20 filename contains "DRAFT" (suggests superseded version)
  // ─────────────────────────────────────────────────────────────────────────
  function scoreResult(result, projectName, sourceDef) {
    let score = 0;
    const nameLower    = result.name.toLowerCase();
    const projectLower = projectName.toLowerCase();
    const pathLower    = (result.path || '').toLowerCase();

    if (nameLower.includes(projectLower)) score += 100;

    const keywordHit = sourceDef.keywords.some(kw => nameLower.includes(kw.toLowerCase()));
    if (keywordHit) score += 50;

    const folderHit    = sourceDef.folderHint    && pathLower.includes(sourceDef.folderHint.toLowerCase());
    const folderHitAlt = sourceDef.folderHintAlt && pathLower.includes(sourceDef.folderHintAlt.toLowerCase());
    if (folderHit || folderHitAlt) score += 30;

    // Extension scoring — preferred formats rank above acceptable; other formats get nothing.
    const prefIdx = sourceDef.preferredExtensions.indexOf(result.ext);
    const accIdx  = sourceDef.acceptableExtensions.indexOf(result.ext);
    if (prefIdx === 0)      score += 20;
    else if (prefIdx >= 1)  score += 15;
    else if (accIdx >= 0)   score += 10;
    // 'other' extension: no bonus, no penalty — let other signals decide

    if (/\bdraft\b/i.test(result.name)) score -= 20;

    return score;
  }

  // ─────────────────────────────────────────────────────────────────────────
  // CONFIDENCE ASSESSMENT
  //
  //   'none'      — nothing cleared the minimum relevance bar (score >= 80)
  //   'confident' — top score >= 130 and gap over 2nd place >= 40
  //   'ambiguous' — everything else with at least one result above the bar
  //
  // Score 130 means "project name + keyword + preferred extension" (100+50+20 = 170
  // at best; realistically 150 when folder hint also matches). The 130 threshold
  // requires at minimum a project-name match plus keyword or format signal.
  // ─────────────────────────────────────────────────────────────────────────
  function assessConfidence(scored) {
    if (!scored.length) return 'none';
    const top = scored[0].score;
    if (top < 80) return 'none';
    if (scored.length === 1) return 'confident';
    const second = scored[1].score;
    if (top >= 130 && (top - second) >= 40) return 'confident';
    return 'ambiguous';
  }

  // ─────────────────────────────────────────────────────────────────────────
  // PUBLIC: find()
  // ─────────────────────────────────────────────────────────────────────────

  /**
   * Find the best-matching file for a project + source type.
   *
   * @param {string} projectName  — last name or identifier fragment (fuzzy match OK)
   * @param {string} sourceType   — key from SOURCE_TYPES
   * @param {object} [options]
   * @param {number} [options.maxCandidates=5] — max entries in candidates array
   * @returns {Promise<{status, file, candidates, query}>}
   */
  async function find(projectName, sourceType, options = {}) {
    const maxCandidates = options.maxCandidates || 5;
    const sourceDef = SOURCE_TYPES[sourceType];

    if (!sourceDef) {
      throw new Error(
        `SherpaLookup: unknown source type "${sourceType}". ` +
        `Valid: ${Object.keys(SOURCE_TYPES).join(', ')}`
      );
    }
    if (!projectName || !projectName.trim()) {
      throw new Error('SherpaLookup: projectName is required');
    }

    // Query: project name + primary keyword. Graph ranks results that match both
    // above those that match only one — a natural preference signal for free.
    const query = `${projectName.trim()} ${sourceDef.keywords[0]}`;

    let results;
    try {
      results = await searchAndEnrich(query);
    } catch (e) {
      console.error('SherpaLookup: search failed', e);
      return { status: 'none', file: null, candidates: [], query, error: e.message };
    }

    // Score, annotate with formatHint, sort
    const scored = results.map(r => ({
      ...r,
      score:      scoreResult(r, projectName, sourceDef),
      formatHint: formatHintFor(r.ext, sourceDef),
    }));

    scored.sort((a, b) => {
      if (b.score !== a.score) return b.score - a.score;
      return new Date(b.modified) - new Date(a.modified); // tiebreaker: newest first
    });

    const status     = assessConfidence(scored);
    const file       = status === 'none' ? null : scored[0];
    // Candidates are alternatives to the primary. When a primary exists, exclude
    // it from candidates so callers don't get the same result twice. When no
    // primary exists ('none'), return the full ranked list for diagnostic use.
    const candidates = file
      ? scored.slice(1, 1 + maxCandidates)
      : scored.slice(0, maxCandidates);

    return { status, file, candidates, query };
  }

  // ─────────────────────────────────────────────────────────────────────────
  // HELPERS
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
    _SOURCE_DEFS: SOURCE_TYPES, // for debugging; do not depend on internal shape
  };

})(window);
