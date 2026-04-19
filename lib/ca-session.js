/**
 * ca-session.js — CA Writing Tool shared session module
 * Version: 1 (stub)
 * Hosted at: bellweatherllc.github.io/tools/ca-session.js
 *
 * Owns: project selection, folder resolution, session state caching,
 *       and the projectChanged event that all screens and components listen to.
 *
 * Every CA tool screen (Setup, Sources, Generate, Document) loads this module.
 * No screen implements project resolution inline. No screen duplicates this logic.
 *
 * Usage:
 *   <script src="./ca-session.js"></script>
 *
 *   // Select a project (call from any screen's project dropdown):
 *   await CASession.selectProject(projectId);
 *
 *   // Get the current project (returns null if none selected):
 *   const project = CASession.getProject();
 *
 *   // Listen for project changes (any screen, any component):
 *   document.addEventListener('projectChanged', (e) => {
 *     const project = e.detail; // null if cleared
 *     // update UI, re-run Sherpa, etc.
 *   });
 *
 *   // Clear session (e.g. sign-out):
 *   CASession.clearProject();
 */

const CASession = (() => {

  /* ─────────────────────────────────────────────
   * CONSTANTS
   * ───────────────────────────────────────────── */

  const SESSION_KEY    = 'ca-session-project';
  const PROJECTS_SITE  = 'https://graph.microsoft.com/v1.0/sites/bellweather.sharepoint.com:/sites/BWCore';
  const PROJECTS_LIST  = 'CORE_TeamMembers'; // TODO: confirm CORE_Projects list name
  const PHASE_FOLDERS  = ['1. Active Leads', '2. Active-Design', '3. Active-Construction'];

  /* ─────────────────────────────────────────────
   * INTERNAL STATE
   * ───────────────────────────────────────────── */

  let _currentProject  = null;  // resolved project object, or null
  let _folderIndex     = [];    // [{folderName, lastName, id, driveId, phaseFolder}]
  let _folderIndexReady = false;

  /* ─────────────────────────────────────────────
   * PUBLIC: selectProject(projectId)
   *
   * Resolves a CORE_Projects record by its SharePoint list item ID,
   * matches it to a Customers library folder, caches the result,
   * and fires 'projectChanged' on document.
   *
   * Requires SherpaAuth to be initialized (sherpa-auth.js loaded + signed in).
   * ───────────────────────────────────────────── */

  async function selectProject(projectId) {
    if (!projectId) {
      _currentProject = null;
      _persist(null);
      _emit(null);
      return null;
    }

    // TODO: look up from allProjects array (passed in or fetched here)
    // For now, caller is expected to pass the full resolved project object
    // as the second argument during the transition period.
    // Final form: this function fetches from CORE_Projects by ID itself.
    throw new Error(
      'ca-session.js selectProject() stub — not yet implemented. ' +
      'CA Sources currently handles project resolution internally. ' +
      'Wire this after ca-session.js is integrated into Sources.'
    );
  }

  /* ─────────────────────────────────────────────
   * PUBLIC: setProject(resolvedProject)
   *
   * Transition-period helper. CA Sources (and eventually Setup) calls this
   * directly with an already-resolved project object until selectProject()
   * is fully implemented. All other behavior is identical — caches, emits.
   * ───────────────────────────────────────────── */

  function setProject(resolvedProject) {
    _currentProject = resolvedProject || null;
    _persist(_currentProject);
    _emit(_currentProject);
    return _currentProject;
  }

  /* ─────────────────────────────────────────────
   * PUBLIC: getProject()
   *
   * Returns the current resolved project object, or null.
   * Restores from sessionStorage on first call (survives page reload
   * within the same browser tab session).
   * ───────────────────────────────────────────── */

  function getProject() {
    if (_currentProject) return _currentProject;
    try {
      const saved = sessionStorage.getItem(SESSION_KEY);
      if (saved) {
        _currentProject = JSON.parse(saved);
        return _currentProject;
      }
    } catch (_) {}
    return null;
  }

  /* ─────────────────────────────────────────────
   * PUBLIC: clearProject()
   *
   * Clears project selection. Call on sign-out or explicit reset.
   * ───────────────────────────────────────────── */

  function clearProject() {
    _currentProject = null;
    _persist(null);
    _emit(null);
  }

  /* ─────────────────────────────────────────────
   * PUBLIC: loadFolderIndex()
   *
   * Fetches the Customers library folder list across all active phase folders
   * and builds the internal index used for folder matching.
   * Called once per session, typically on sign-in.
   *
   * Requires: SherpaAuth initialized and signed in, cachedCustomersDriveId set.
   * TODO: Accept driveId as a parameter rather than requiring it as a global.
   * ───────────────────────────────────────────── */

  async function loadFolderIndex(driveId) {
    if (_folderIndexReady) return _folderIndex;
    if (!driveId) {
      console.warn('ca-session: loadFolderIndex called without driveId — skipping');
      return [];
    }

    const fetchPhase = async (phaseName) => {
      try {
        const url = `${SherpaAuth.GRAPH_BASE}/drives/${driveId}/root:/${encodeURIComponent(phaseName)}:/children?$select=name,id,webUrl,folder&$top=500`;
        const data = await SherpaAuth.graph(url);
        return (data.value || [])
          .filter(item => item.folder)
          .map(item => ({
            phaseFolder: phaseName,
            folderName:  item.name,
            lastName:    _parseLastName(item.name),
            id:          item.id,
            driveId:     driveId,
            webUrl:      item.webUrl,
          }));
      } catch (e) {
        console.warn(`ca-session: folder list failed for "${phaseName}":`, e);
        return [];
      }
    };

    const lists = await Promise.all(PHASE_FOLDERS.map(fetchPhase));
    _folderIndex = lists.flat();
    _folderIndexReady = true;
    return _folderIndex;
  }

  /* ─────────────────────────────────────────────
   * PUBLIC: matchProjectToFolder(project)
   *
   * Returns the best-matching folder from the index for a given
   * CORE_Projects record, or null if no match.
   * Same logic currently duplicated in CA Sources — consolidate here.
   * ───────────────────────────────────────────── */

  function matchProjectToFolder(project) {
    if (!_folderIndex.length) return null;
    const cn = (project.ClientName || '').toLowerCase();
    if (!cn) return null;

    const cleaned = cn.replace(/["'][^"']*["']/g, ' ');
    const STOP = new Set(['the','and','or','of','a','an','mr','mrs','ms','dr','jr','sr','family','home','residence','household']);
    const tokens = new Set(
      cleaned.split(/[\s,&\-]+/)
        .map(t => t.replace(/\.+$/, ''))
        .filter(t => t && !STOP.has(t))
    );

    const hits = _folderIndex.filter(f => tokens.has(f.lastName.toLowerCase()));
    if (!hits.length) return null;
    if (hits.length === 1) return { ...hits[0], ambiguous: false };

    const jc = (project.JobCode || '').toLowerCase().trim();
    if (jc) {
      const jcHit = hits.find(f => f.folderName.toLowerCase().includes(jc));
      if (jcHit) return { ...jcHit, ambiguous: false };
    }
    return { ...hits[0], ambiguous: true, allMatches: hits };
  }

  /* ─────────────────────────────────────────────
   * INTERNAL HELPERS
   * ───────────────────────────────────────────── */

  function _parseLastName(folderName) {
    if (!folderName) return '';
    const m = folderName.trim().match(/^(.+?)(?:\s+-\s+|\s+-|-\s+)/);
    return (m ? m[1] : folderName).trim();
  }

  function _persist(project) {
    try {
      if (project) sessionStorage.setItem(SESSION_KEY, JSON.stringify(project));
      else sessionStorage.removeItem(SESSION_KEY);
    } catch (_) {}
  }

  function _emit(project) {
    document.dispatchEvent(new CustomEvent('projectChanged', { detail: project }));
  }

  /* ─────────────────────────────────────────────
   * EXPORT
   * ───────────────────────────────────────────── */

  return {
    selectProject,
    setProject,
    getProject,
    clearProject,
    loadFolderIndex,
    matchProjectToFolder,
  };

})();
