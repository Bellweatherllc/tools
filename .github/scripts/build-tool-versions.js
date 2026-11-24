// Writes tool-versions.json: the current version and last-modified time of every
// page in the repo, read by bw-tool-index.html to label its tiles.
//
// Run by .github/workflows/tool-versions.yml on every push to main (needs full
// git history: actions/checkout with fetch-depth: 0). Safe to run locally:
//   node .github/scripts/build-tool-versions.js
//
// Version: taken from the page itself — a CORE_VERSION / MGR_VERSION / VERSION
// constant if it has one, otherwise the first badge element that holds "vN".
// Modified: the last commit touching the page. A manual's words live in its .md
// (the .html only loads it), so a manual takes the later of its .html and .md.
const fs = require('fs');
const path = require('path');
const { execFileSync } = require('child_process');

const ROOT = path.resolve(__dirname, '..', '..');
const SKIP_DIRS = new Set(['.git', '.github', 'archive', 'node_modules', 'azure-function-proxy', 'lib', 'fonts', 'logos']);

function pages(dir, rel = '') {
  const out = [];
  for (const ent of fs.readdirSync(dir, { withFileTypes: true })) {
    if (ent.isDirectory()) {
      if (!SKIP_DIRS.has(ent.name) && !ent.name.startsWith('.')) out.push(...pages(path.join(dir, ent.name), rel + ent.name + '/'));
    } else if (ent.name.endsWith('.html')) {
      out.push(rel + ent.name);
    }
  }
  return out;
}

function lastCommit(file) {
  try {
    return execFileSync('git', ['log', '-1', '--format=%cI', '--', file], { cwd: ROOT, encoding: 'utf8' }).trim() || null;
  } catch { return null; }
}

function version(src) {
  const c = src.match(/const\s+(?:CORE_VERSION|MGR_VERSION|VERSION)\s*=\s*['"]?(v?\d+(?:\.\d+)*)['"]?\s*;/);
  if (c) return c[1].startsWith('v') ? c[1] : 'v' + c[1];
  const b = src.match(/<[^>]*(?:ver|badge)[^>]*>\s*(v\d+(?:\.\d+)*)\s*</i);
  return b ? b[1] : null;
}

const files = {};
for (const rel of pages(ROOT).sort()) {
  const src = fs.readFileSync(path.join(ROOT, rel), 'utf8');
  let modified = lastCommit(rel);
  const md = rel.replace(/\.html$/, '.md');
  if (fs.existsSync(path.join(ROOT, md)) && /fetch\(/.test(src)) {
    const mdTime = lastCommit(md);
    if (mdTime && (!modified || Date.parse(mdTime) > Date.parse(modified))) modified = mdTime;
  }
  if (!modified) continue;  // not committed yet
  files[rel] = { version: version(src), modified };
}

const out = JSON.stringify({ files }, null, 1) + '\n';
fs.writeFileSync(path.join(ROOT, 'tool-versions.json'), out);
console.log(`tool-versions.json: ${Object.keys(files).length} pages`);
