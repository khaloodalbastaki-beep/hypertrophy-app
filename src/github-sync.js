// ============================================================
// GitHub two-way sync
// ------------------------------------------------------------
// Reads/writes the workout data in a PRIVATE GitHub repo via the
// Contents API, so the same data is available on phone + Mac.
//   history.json  — settings, history (keyed by session), bodyweight, videos  (small, frequent)
//   photos.json   — exercise reference photos (base64)                        (large, rare)
// The fine-grained token (Contents: Read & Write on the data repo) is entered
// once in Settings and kept in this browser's localStorage — never in code.
// ============================================================

const API = 'https://api.github.com';
const DEFAULT_REPO = 'khaloodalbastaki-beep/hypertrophy-data';

export const getToken = () => localStorage.getItem('hyp_gh_token') || '';
export const getRepo = () => localStorage.getItem('hyp_gh_repo') || DEFAULT_REPO;
export const isConfigured = () => !!getToken();
export const setConfig = (token, repo) => {
  localStorage.setItem('hyp_gh_token', (token || '').trim());
  localStorage.setItem('hyp_gh_repo', (repo || DEFAULT_REPO).trim());
};
export const clearConfig = () => { localStorage.removeItem('hyp_gh_token'); };

const b64encode = (str) => btoa(unescape(encodeURIComponent(str)));
const b64decode = (b64) => decodeURIComponent(escape(atob((b64 || '').replace(/\n/g, ''))));

const headers = () => ({ Authorization: `Bearer ${getToken()}`, Accept: 'application/vnd.github+json' });

// GET a JSON file -> { json, sha }  (json/sha null if the file doesn't exist yet)
async function getFile(path) {
  const r = await fetch(`${API}/repos/${getRepo()}/contents/${encodeURIComponent(path)}`, { headers: headers() });
  if (r.status === 404) return { json: null, sha: null };
  if (!r.ok) throw new Error(`GET ${path} ${r.status}`);
  const data = await r.json();
  let text;
  if (data.content) text = b64decode(data.content);
  else if (data.download_url) text = await (await fetch(data.download_url)).text(); // >1MB files
  else throw new Error(`GET ${path} no content`);
  return { json: JSON.parse(text), sha: data.sha };
}

// PUT a JSON file. Pass the known sha to update; omit to create. Returns the new sha.
async function putFile(path, obj, sha, message) {
  const body = { message: message || `update ${path}`, content: b64encode(JSON.stringify(obj)) };
  if (sha) body.sha = sha;
  const r = await fetch(`${API}/repos/${getRepo()}/contents/${encodeURIComponent(path)}`, {
    method: 'PUT', headers: headers(), body: JSON.stringify(body),
  });
  if (r.status === 409 || r.status === 422) { const e = new Error('conflict'); e.conflict = true; throw e; }
  if (!r.ok) throw new Error(`PUT ${path} ${r.status}`);
  return (await r.json()).content.sha;
}

export async function testConnection() {
  const r = await fetch(`${API}/repos/${getRepo()}`, { headers: headers() });
  if (!r.ok) return { ok: false, status: r.status };
  const repo = await r.json();
  return { ok: true, canWrite: !!(repo.permissions && repo.permissions.push) };
}

// ---- merge helpers (protect offline edits; single user, last-writer-wins per session) ----
const setCount = (rec) => Object.values(rec.sets || {}).reduce((s, a) => s + (a ? a.length : 0), 0);
function pickSession(a, b) {
  const sa = setCount(a), sb = setCount(b);
  if (sa !== sb) return sa > sb ? a : b;
  return (a.endTime || a.startTime || '') >= (b.endTime || b.startTime || '') ? a : b;
}
export function mergeHistory(localH, remoteH) {
  const out = { ...(remoteH || {}) };
  for (const [k, rec] of Object.entries(localH || {})) out[k] = out[k] ? pickSession(rec, out[k]) : rec;
  return out;
}
function mergeBodyweight(a, b) {
  const m = {}; [...(a || []), ...(b || [])].forEach((e) => { if (e && e.date) m[e.date] = e; });
  return Object.values(m).sort((x, y) => (x.date < y.date ? -1 : 1));
}
// Feedback = append-only list of bug/idea reports written in-app. Union by id so
// nothing the user typed offline is lost; a 'done' status wins over 'open'; newest first.
function mergeFeedback(a, b) {
  const m = {};
  [...(a || []), ...(b || [])].forEach((e) => {
    if (!e || !e.id) return;
    const cur = m[e.id];
    if (!cur || e.status === 'done') m[e.id] = e;
  });
  return Object.values(m).sort((x, y) => ((y.createdAt || '') < (x.createdAt || '') ? -1 : 1));
}

// Pull both files and return merged data vs the passed-in local state.
export async function pullMerge(local) {
  const [h, p] = await Promise.all([getFile('history.json'), getFile('photos.json')]);
  const rH = h.json || {}; const rP = p.json || {};
  const remoteNewer = (rH.updatedAt || '') > (local.updatedAt || '');
  const mergedHistory = mergeHistory(local.history, rH.history);
  const mergedBw = mergeBodyweight(local.bodyweightLog, rH.bodyweightLog);
  const mergedFeedback = mergeFeedback(local.feedback, rH.feedback);
  const mergedNotes = { ...(rH.notes || {}), ...(local.notes || {}) };
  const mergedUnits = { ...(rH.units || {}), ...(local.units || {}) };
  return {
    shas: { history: h.sha, photos: p.sha },
    settings: rH.settings && (remoteNewer || !local.hasLocal) ? rH.settings : local.settings,
    history: mergedHistory,
    bodyweightLog: mergedBw,
    videos: { ...(rH.videos || {}), ...(local.videos || {}) },
    photos: { ...(rP.photos || {}), ...(local.photos || {}) },
    feedback: mergedFeedback,
    notes: mergedNotes,
    units: mergedUnits,
    remoteExisted: !!h.sha,
    // local has sessions/bodyweight/feedback the remote doesn't yet -> caller should push
    localAhead: JSON.stringify(mergedHistory) !== JSON.stringify(rH.history || {})
      || JSON.stringify(mergedBw) !== JSON.stringify(rH.bodyweightLog || [])
      || JSON.stringify(mergedFeedback) !== JSON.stringify(rH.feedback || [])
      || JSON.stringify(mergedNotes) !== JSON.stringify(rH.notes || {})
      || JSON.stringify(mergedUnits) !== JSON.stringify(rH.units || {}),
  };
}

export async function pushHistory(state, sha) {
  const doc = { version: 5, updatedAt: new Date().toISOString(), settings: state.settings, history: state.history, bodyweightLog: state.bodyweightLog, videos: state.videos, feedback: state.feedback || [], notes: state.notes || {}, units: state.units || {} };
  try { return await putFile('history.json', doc, sha, 'update history'); }
  catch (e) {
    if (!e.conflict) throw e;
    const cur = await getFile('history.json');
    const merged = {
      ...doc,
      history: mergeHistory(state.history, cur.json?.history),
      bodyweightLog: mergeBodyweight(state.bodyweightLog, cur.json?.bodyweightLog),
      feedback: mergeFeedback(state.feedback, cur.json?.feedback),
      notes: { ...(cur.json?.notes || {}), ...(state.notes || {}) },
      units: { ...(cur.json?.units || {}), ...(state.units || {}) },
    };
    return await putFile('history.json', merged, cur.sha, 'update history (merged)');
  }
}

export async function pushPhotos(state, sha) {
  const doc = { version: 5, updatedAt: new Date().toISOString(), photos: state.photos };
  try { return await putFile('photos.json', doc, sha, 'update photos'); }
  catch (e) {
    if (!e.conflict) throw e;
    const cur = await getFile('photos.json');
    return await putFile('photos.json', { ...doc, photos: { ...(cur.json?.photos || {}), ...state.photos } }, cur.sha, 'update photos (merged)');
  }
}
