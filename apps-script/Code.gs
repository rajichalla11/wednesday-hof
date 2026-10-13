/**
 * Wednesday 9–10 · Player Ratings backend
 * ------------------------------------------------------------
 * Lives in a PRIVATE Google Sheet (not the attendance sheet).
 * Deploy: Extensions → Apps Script → Deploy → New deployment →
 *   Web app · Execute as: Me · Who has access: Anyone
 * Then paste the web-app URL into index.html as VOTE_API.
 *
 * Privacy model
 *   Each player gets a secret link (token). We store TWO different hashes
 *   of it: id1 identifies the voter (so we can check they exist and whom
 *   they're allowed to rate) and id2 keys their ballot. Neither can be
 *   turned back into the other without the raw token, which only exists
 *   in the player's own chat once the Links tab is burned. Votes are
 *   therefore anonymous to everyone, including the sheet owner.
 */

const MAIN_SHEET_ID = '1UUcM5qxOEA-Ykr-gIG4WSHDlzpoi0TJAaRh-8HgYRaI';
const MAIN_CSV = 'https://docs.google.com/spreadsheets/d/' + MAIN_SHEET_ID + '/export?format=csv&gid=0';
const SITE_URL = 'https://wednesday910.vercel.app/';
const STATS = ['PAC', 'SHO', 'PAS', 'DRI', 'DEF', 'PHY'];
const TRAITS = ['ENGINE', 'FINISHER', 'PLAYMAKER', 'WALL', 'MAGICIAN', 'ROCKET', 'CAPTAIN', 'CLUTCH', 'SWEEPER', 'GLOVES', 'WILDCARD', 'ENFORCER', 'HEART', 'BRAIN', 'HUSTLER', 'SNIPER', 'JOKER', 'ROCK'];
const MAX_TRAITS = 3;
const MIN_SHARED = 3;      // sessions you must have shared to rate someone
const MIN_VOTES = 3;       // ballots before a rating is shown for a player with no baseline
const PRIOR_WEIGHT = 3;    // the admin's exact baseline counts like this many votes
const RMIN = 40, RMAX = 99;
const SESSION_START_COL = 8;
// Starting points by position (PAC SHO PAS DRI DEF PHY) — overwrite in the Seeds tab
const SEED_BY_POS = { attack:[72,74,64,70,52,64], midfield:[68,64,74,70,64,66], defense:[66,54,64,58,74,72], other:[65,65,65,65,65,65] };
function posClass_(pos) { const p = String(pos || '').toLowerCase(); return p.indexOf('attack') >= 0 ? 'attack' : (p.indexOf('defen') >= 0) ? 'defense' : p.indexOf('mid') >= 0 ? 'midfield' : 'other'; }
const SKIP = ['', 'player', 'games played', 'amount paid', 'balance', 'owing from previous sessions', '+1s', 'krish p', 'images', 'postion', 'position', 'preferred foot'];

/* ───────── admin: run these from the editor ───────── */

function setup() {
  const v = sheet_('Voters');  if (v.getLastRow() === 0) v.appendRow(['voter_id', 'name', 'created', 'claimed']);
  if (v.getLastColumn() < 4) v.getRange(1, 4).setValue('claimed');
  settings_();   // creates the Settings tab with a group code if missing
  const b = sheet_('Votes');   if (b.getLastRow() === 0) b.appendRow(['window', 'ballot_id', 'target', 'w'].concat(STATS, ['comment', 'updated', 'traits']));
  if (b.getLastColumn() < 4 + STATS.length + 3) b.getRange(1, 4 + STATS.length + 3).setValue('traits');
  const sr = sheet_('SelfRatings'); if (sr.getLastRow() === 0) sr.appendRow(['window', 'name'].concat(STATS, ['traits', 'note', 'updated']));
  const w = sheet_('Windows'); if (w.getLastRow() === 0) w.appendRow(['window', 'opened', 'closed']);
  const ph = sheet_('Photos'); if (ph.getLastRow() === 0) ph.appendRow(['name', 'file_id', 'updated']);
  photoFolder_();   // creates the Drive folder for player photos (and asks for the Drive permission the first time)
  const s = sheet_('Seeds');
  if (s.getLastRow() === 0) {
    s.appendRow(['name'].concat(STATS));
    players_().forEach(function (p) { s.appendRow([p.name].concat(SEED_BY_POS[posClass_(p.position)])); });
  }
  if (!props_().getProperty('window')) openWindow();
  if (rows_(sheet_('Voters')).length === 0) generateLinks();   // first run: create everyone's links too
}

/** Creates a personal link for every player who doesn't have one yet. Writes them to a "Links" tab — send them, then run burnLinks(). */
function generateLinks() {
  const voters = sheet_('Voters');
  const have = {};
  rows_(voters).forEach(function (r) { have[String(r[1]).toLowerCase()] = true; });
  const links = sheet_('Links');
  if (links.getLastRow() === 0) links.appendRow(['name', 'link']);
  let made = 0;
  players_().forEach(function (p) {
    if (have[p.name.toLowerCase()]) return;
    const token = newToken_();
    voters.appendRow([id1_(token), p.name, new Date(), '']);
    links.appendRow([p.name, SITE_URL + '#vote/' + token]);
    made++;
  });
  Logger.log(made + ' links created. Send them from the Links tab, then run burnLinks().');
}

/** Deletes the Links tab so raw tokens no longer exist in the sheet. */
function burnLinks() {
  const s = ss_().getSheetByName('Links');
  if (s) ss_().deleteSheet(s);
  cache_().remove('results');
}

/** Re-issues one player's link (old link stops working; their old ballot stays but is orphaned). */
function rotateLink(name) {
  const voters = sheet_('Voters');
  const data = voters.getDataRange().getValues();
  for (let i = data.length - 1; i >= 1; i--) if (String(data[i][1]).toLowerCase() === String(name).toLowerCase()) voters.deleteRow(i + 1);
  generateLinks();
}

function openWindow(id) {
  id = id || defaultWindow_();
  props_().setProperties({ window: id, open: '1' });
  sheet_('Windows').appendRow([id, new Date(), '']);
  cache_().remove('results');
}
function closeWindow() {
  props_().setProperty('open', '0');
  const w = sheet_('Windows'), data = w.getDataRange().getValues();
  for (let i = data.length - 1; i >= 1; i--) if (data[i][0] === props_().getProperty('window')) { w.getRange(i + 1, 3).setValue(new Date()); break; }
  cache_().remove('results');
}

/* ───────── web app ───────── */

function doGet(e) {
  const a = (e.parameter.action || 'results');
  try {
    if (a === 'me') return json_(me_(e.parameter.t));
    if (a === 'admin') return json_(admin_(e.parameter.t));
    if (a === 'results') return json_(results_());
    return json_({ ok: false, error: 'unknown action' });
  } catch (err) { return json_({ ok: false, error: String(err) }); }
}

function doPost(e) {
  let body = {};
  try { body = JSON.parse(e.postData.contents || '{}'); } catch (_) {}
  try {
    if (body.action === 'rate') return json_(rate_(body));
    if (body.action === 'claim') return json_(claim_(body));
    if (body.action === 'self') return json_(self_(body));
    if (body.action === 'setSeed') return json_(setSeed_(body));
    if (body.action === 'unrate') return json_(unrate_(body));
    if (body.action === 'photo') return json_(photo_(body));
    return json_({ ok: false, error: 'unknown action' });
  } catch (err) { return json_({ ok: false, error: String(err) }); }
}

function me_(token) {
  const name = voterName_(token);
  if (!name) return { ok: false, error: 'This link isn’t recognised.' };
  const win = props_().getProperty('window'), open = props_().getProperty('open') === '1';
  const mine = {}, feedback = [];
  const b = id2_(token);
  rows_(sheet_('Votes')).forEach(function (r) {
    if (r[0] !== win) return;
    if (r[1] === b) { const st = {}; STATS.forEach(function (k, i) { st[k] = Number(r[4 + i]); }); mine[r[2]] = { stats: st, comment: r[4 + STATS.length] || '', traits: String(r[4 + STATS.length + 2] || '').split(',').filter(Boolean) }; }
    if (r[2] === name && r[4 + STATS.length]) feedback.push(String(r[4 + STATS.length]));
  });
  const selfRow = rows_(sheet_('SelfRatings')).filter(function (r) { return r[0] === win && String(r[1]).toLowerCase() === name.toLowerCase(); }).pop();
  const self = selfRow ? rowToRating_(selfRow, 2) : null;
  return { ok: true, name: name, window: win, open: open, seeds: seeds_(), mine: mine, feedback: shuffle_(feedback), minShared: MIN_SHARED,
           traits: TRAITS, maxTraits: MAX_TRAITS, self: self, isAdmin: isAdmin_(name) };
}

function rate_(body) {
  const name = voterName_(body.t);
  if (!name) return { ok: false, error: 'This link isn’t recognised.' };
  if (props_().getProperty('open') !== '1') return { ok: false, error: 'The ratings window is closed.' };
  const target = String(body.target || '').trim();
  const ps = players_(), byName = {};
  ps.forEach(function (p) { byName[p.name.toLowerCase()] = p; });
  const me = byName[name.toLowerCase()], tp = byName[target.toLowerCase()];
  if (!tp) return { ok: false, error: 'Unknown player.' };
  if (tp === me) return { ok: false, error: 'You can’t rate yourself.' };
  const shared = shared_(me, tp);
  if (shared < MIN_SHARED) return { ok: false, error: 'You haven’t played enough sessions with ' + tp.name + '.' };
  const stats = STATS.map(function (k) { const v = Math.round(Number((body.stats || {})[k])); if (isNaN(v)) throw new Error('Missing ' + k); return Math.max(RMIN, Math.min(RMAX, v)); });
  const comment = String(body.comment || '').replace(/\s+/g, ' ').trim().slice(0, 280);
  const traits = cleanTraits_(body.traits);
  const w = shared >= 20 ? 1 : shared >= 8 ? 0.8 : 0.6;   // familiarity, bucketed so it can't identify the voter
  const win = props_().getProperty('window'), b = id2_(body.t);

  const lock = LockService.getScriptLock(); lock.waitLock(10000);
  try {
    const sh = sheet_('Votes'), data = sh.getDataRange().getValues();
    let rowIdx = -1;
    for (let i = 1; i < data.length; i++) if (data[i][0] === win && data[i][1] === b && data[i][2] === tp.name) { rowIdx = i + 1; break; }
    const row = [win, b, tp.name, w].concat(stats, [comment, new Date(), traits.join(',')]);
    if (rowIdx > 0) sh.getRange(rowIdx, 1, 1, row.length).setValues([row]); else sh.appendRow(row);
  } finally { lock.releaseLock(); }
  cache_().remove('results');
  return { ok: true };
}

function unrate_(body) {
  const name = voterName_(body.t);
  if (!name) return { ok: false, error: 'This link isn’t recognised.' };
  const win = props_().getProperty('window'), b = id2_(body.t), target = String(body.target || '');
  const sh = sheet_('Votes'), data = sh.getDataRange().getValues();
  for (let i = data.length - 1; i >= 1; i--) if (data[i][0] === win && data[i][1] === b && data[i][2] === target) sh.deleteRow(i + 1);
  cache_().remove('results');
  return { ok: true };
}

/** Self-service: a player picks their name on the site and enters the group code to get their personal link.
 *  One claim per name. An unclaimed pre-generated link is replaced. To let someone claim again, delete
 *  their row in the Voters tab. The code lives in the Settings tab (key claim_code). */
function claim_(body) {
  const name = String(body.name || '').trim(), code = String(body.code || '').trim();
  if (!name) return { ok: false, error: 'Pick your name.' };
  const key = 'claim|' + name.toLowerCase(), tries = Number(cache_().get(key) || 0);
  if (tries >= 5) return { ok: false, error: 'Too many attempts for this name. Try again in 15 minutes.' };
  const want = String(settings_().claim_code || '').trim().toLowerCase();   // blank in Settings = no code needed
  if (want && code.toLowerCase() !== want) { cache_().put(key, String(tries + 1), 900); return { ok: false, error: 'That code is not right.' }; }
  const p = players_().filter(function (x) { return x.name.toLowerCase() === name.toLowerCase(); })[0];
  if (!p) return { ok: false, error: 'That name is not on the attendance sheet.' };
  const lock = LockService.getScriptLock(); lock.waitLock(10000);
  try {
    const voters = sheet_('Voters'), data = voters.getDataRange().getValues();
    for (let i = data.length - 1; i >= 1; i--) {
      if (String(data[i][1]).toLowerCase() !== p.name.toLowerCase()) continue;
      if (data[i][3]) return { ok: false, error: 'This name has already been claimed. If that was not you, ask Raji to reset it.' };
      voters.deleteRow(i + 1);
    }
    const token = newToken_();
    voters.appendRow([id1_(token), p.name, new Date(), new Date()]);
    cache_().remove(key);
    return { ok: true, token: token, name: p.name };
  } finally { lock.releaseLock(); }
}

/** Aggregated ratings: calibrated, trimmed, familiarity-weighted means. Cached 2 min. */
function results_() {
  const cached = cache_().get('results');
  if (cached) return JSON.parse(cached);
  const win = props_().getProperty('window'), open = props_().getProperty('open') === '1';
  const windows = rows_(sheet_('Windows')).map(function (r) { return r[0]; });
  const prevWin = windows[windows.indexOf(win) - 1] || null;
  const all = rows_(sheet_('Votes'));
  const seeds = seeds_();
  const out = { ok: true, window: win, open: open, minVotes: MIN_VOTES, priorWeight: PRIOR_WEIGHT, requireCode: !!String(settings_().claim_code || '').trim(), photos: photos_(), players: aggregate_(all.filter(function (r) { return r[0] === win; }), seeds), prev: prevWin ? aggregate_(all.filter(function (r) { return r[0] === prevWin; }), seeds) : {}, prevWindow: prevWin };
  out.ballots = Object.keys(all.filter(function (r) { return r[0] === win; }).reduce(function (m, r) { m[r[1]] = 1; return m; }, {})).length;
  cache_().put('results', JSON.stringify(out), 120);
  return out;
}

function aggregate_(rows, seeds) {
  seeds = seeds || {};
  // voter calibration: how generous is each ballot compared with everyone?
  const perBallot = {}; let gSum = 0, gN = 0;
  rows.forEach(function (r) { const b = r[1]; perBallot[b] = perBallot[b] || { s: 0, n: 0 }; STATS.forEach(function (_, i) { const v = Number(r[4 + i]); perBallot[b].s += v; perBallot[b].n++; gSum += v; gN++; }); });
  const gMean = gN ? gSum / gN : 0;
  const offset = {}; Object.keys(perBallot).forEach(function (b) { offset[b] = perBallot[b].s / perBallot[b].n - gMean; });

  const byTarget = {};
  rows.forEach(function (r) { (byTarget[r[2]] = byTarget[r[2]] || []).push(r); });
  const res = {};
  const names = {}; Object.keys(byTarget).forEach(function (t) { names[t] = 1; }); Object.keys(seeds).forEach(function (t) { names[t] = 1; });
  Object.keys(names).forEach(function (t) {
    const vs = byTarget[t] || [], n = vs.length, base = seeds[t] || null;
    const entry = { n: n, traits: {} };
    vs.forEach(function (r) { String(r[4 + STATS.length + 2] || '').split(',').filter(Boolean).forEach(function (t) { entry.traits[t] = (entry.traits[t] || 0) + 1; }); });
    if (base) entry.base = base;
    if (base || n >= MIN_VOTES) {
      entry.stats = {};
      STATS.forEach(function (k, i) {
        let items = vs.map(function (r) { return { v: Number(r[4 + i]) - 0.5 * offset[r[1]], w: Number(r[3]) || 1 }; }).sort(function (a, b) { return a.v - b.v; });
        if (n >= 5) { const cut = Math.floor(n * 0.2); items = items.slice(cut, items.length - cut); }
        let sum = items.reduce(function (a, x) { return a + x.v * x.w; }, 0), ws = items.reduce(function (a, x) { return a + x.w; }, 0);
        if (base) { sum += PRIOR_WEIGHT * base[k]; ws += PRIOR_WEIGHT; }   // the baseline is a strong prior; votes trim it
        entry.stats[k] = Math.max(RMIN, Math.min(RMAX, Math.round(sum / ws)));
      });
    }
    res[t] = entry;
  });
  return res;
}

/** A player's own view of themselves. Visible only to the admin (Settings → admin_name), as a starting point to adjust seeds. */
function self_(body) {
  const name = voterName_(body.t);
  if (!name) return { ok: false, error: 'This link isn’t recognised.' };
  const stats = STATS.map(function (k) { const v = Math.round(Number((body.stats || {})[k])); if (isNaN(v)) throw new Error('Missing ' + k); return Math.max(RMIN, Math.min(RMAX, v)); });
  const traits = cleanTraits_(body.traits), note = String(body.note || '').replace(/\s+/g, ' ').trim().slice(0, 280);
  const win = props_().getProperty('window');
  const lock = LockService.getScriptLock(); lock.waitLock(10000);
  try {
    const sh = sheet_('SelfRatings'), data = sh.getDataRange().getValues();
    let rowIdx = -1;
    for (let i = 1; i < data.length; i++) if (data[i][0] === win && String(data[i][1]).toLowerCase() === name.toLowerCase()) { rowIdx = i + 1; break; }
    const row = [win, name].concat(stats, [traits.join(','), note, new Date()]);
    if (rowIdx > 0) sh.getRange(rowIdx, 1, 1, row.length).setValues([row]); else sh.appendRow(row);
  } finally { lock.releaseLock(); }
  return { ok: true };
}

function admin_(token) {
  const name = voterName_(token);
  if (!name || !isAdmin_(name)) return { ok: false, error: 'Admin only.' };
  const win = props_().getProperty('window');
  const selfRatings = rows_(sheet_('SelfRatings')).filter(function (r) { return r[0] === win; }).map(function (r) { const o = rowToRating_(r, 2); o.name = String(r[1]); return o; });
  return { ok: true, window: win, selfRatings: selfRatings, seeds: seeds_(), traits: TRAITS };
}

function setSeed_(body) {
  const name = voterName_(body.t);
  if (!name || !isAdmin_(name)) return { ok: false, error: 'Admin only.' };
  const target = String(body.name || '').trim();
  const stats = STATS.map(function (k) { const v = Math.round(Number((body.stats || {})[k])); if (isNaN(v)) throw new Error('Missing ' + k); return Math.max(RMIN, Math.min(RMAX, v)); });
  const sh = sheet_('Seeds'), data = sh.getDataRange().getValues();
  let rowIdx = -1;
  for (let i = 1; i < data.length; i++) if (String(data[i][0]).toLowerCase() === target.toLowerCase()) { rowIdx = i + 1; break; }
  if (rowIdx > 0) sh.getRange(rowIdx, 2, 1, STATS.length).setValues([stats]); else sh.appendRow([target].concat(stats));
  return { ok: true };
}

/** A player's own photo. Saved to a Drive folder ("W9-10 Photos") owned by the script owner, shared view-only by link; the file id goes in the Photos tab. Admin can pass body.name to set or remove someone else's. */
function photo_(body) {
  const name = voterName_(body.t);
  if (!name) return { ok: false, error: 'This link isn’t recognised.' };
  const target = (body.name && isAdmin_(name)) ? String(body.name).trim() : name;
  const sh = sheet_('Photos'); if (sh.getLastRow() === 0) sh.appendRow(['name', 'file_id', 'updated']);
  const data = sh.getDataRange().getValues(); let rowIdx = -1, oldId = '';
  for (let i = 1; i < data.length; i++) if (String(data[i][0]).toLowerCase() === target.toLowerCase()) { rowIdx = i + 1; oldId = String(data[i][1] || ''); break; }
  if (body.remove) {
    if (rowIdx > 0) sh.deleteRow(rowIdx);
    if (oldId) try { DriveApp.getFileById(oldId).setTrashed(true); } catch (_) {}
    cache_().remove('results'); return { ok: true };
  }
  const m = String(body.data || '').match(/^data:(image\/(jpeg|png|webp));base64,(.+)$/);
  if (!m) return { ok: false, error: 'Send a JPEG, PNG or WebP.' };
  const bytes = Utilities.base64Decode(m[3]);
  if (bytes.length > 1500000) return { ok: false, error: 'That photo is too large.' };
  const file = photoFolder_().createFile(Utilities.newBlob(bytes, m[1], target.replace(/[^\w ]+/g, '') + '.' + (m[2] === 'jpeg' ? 'jpg' : m[2])));
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  const row = [target, file.getId(), new Date()];
  if (rowIdx > 0) sh.getRange(rowIdx, 1, 1, 3).setValues([row]); else sh.appendRow(row);
  if (oldId) try { DriveApp.getFileById(oldId).setTrashed(true); } catch (_) {}
  cache_().remove('results');
  return { ok: true, id: file.getId() };
}
function photoFolder_() {
  const p = props_(), id = p.getProperty('photo_folder');
  if (id) try { return DriveApp.getFolderById(id); } catch (_) {}
  const f = DriveApp.createFolder('W9-10 Photos'); p.setProperty('photo_folder', f.getId()); return f;
}
function photos_() { const o = {}, sh = ss_().getSheetByName('Photos'); if (!sh) return o; rows_(sh).forEach(function (r) { if (r[1]) o[String(r[0])] = String(r[1]); }); return o; }

function isAdmin_(name) { return String(name || '').toLowerCase() === String(settings_().admin_name || 'Raji').toLowerCase(); }
function cleanTraits_(arr) { const seen = {}; return (Array.isArray(arr) ? arr : []).map(function (t) { return String(t).toUpperCase(); }).filter(function (t) { if (TRAITS.indexOf(t) < 0 || seen[t]) return false; seen[t] = 1; return true; }).slice(0, MAX_TRAITS); }
function rowToRating_(r, off) { const st = {}; STATS.forEach(function (k, i) { st[k] = Number(r[off + i]); }); return { stats: st, traits: String(r[off + STATS.length] || '').split(',').filter(Boolean), note: String(r[off + STATS.length + 1] || ''), updated: r[off + STATS.length + 2] ? String(r[off + STATS.length + 2]) : '' }; }

/* ───────── helpers ───────── */

function ss_() { return SpreadsheetApp.getActive(); }
function sheet_(n) { return ss_().getSheetByName(n) || ss_().insertSheet(n); }
function rows_(sh) { const d = sh.getDataRange().getValues(); return d.slice(1).filter(function (r) { return String(r[0]) !== ''; }); }
function props_() { return PropertiesService.getScriptProperties(); }
function cache_() { return CacheService.getScriptCache(); }
function json_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
function sha_(s) { return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, s, Utilities.Charset.UTF_8).map(function (b) { return ('0' + (b & 255).toString(16)).slice(-2); }).join(''); }
function id1_(t) { return sha_('voter|' + t); }
function id2_(t) { return sha_('ballot|' + t); }
function newToken_() { return Utilities.getUuid().replace(/-/g, '') + Utilities.getUuid().replace(/-/g, '').slice(0, 8); }
function settings_() {
  const sh = sheet_('Settings');
  if (sh.getLastRow() === 0) { sh.appendRow(['key', 'value']); sh.appendRow(['claim_code', '']); }   // leave blank for no code
  const o = {}; rows_(sh).forEach(function (r) { o[String(r[0])] = String(r[1]); });
  if (!o.admin_name) { sh.appendRow(['admin_name', 'Raji']); o.admin_name = 'Raji'; }
  return o;
}
function defaultWindow_() { const d = new Date(); return d.getFullYear() + '-Q' + (Math.floor(d.getMonth() / 3) + 1); }
function shuffle_(a) { for (let i = a.length - 1; i > 0; i--) { const j = Math.floor(Math.random() * (i + 1)); const t = a[i]; a[i] = a[j]; a[j] = t; } return a; }

function voterName_(token) {
  if (!token || String(token).length < 16) return null;
  const id = id1_(String(token));
  const hit = rows_(sheet_('Voters')).filter(function (r) { return r[0] === id; })[0];
  return hit ? String(hit[1]) : null;
}

function seeds_() {
  const out = {};
  rows_(sheet_('Seeds')).forEach(function (r) { const st = {}; STATS.forEach(function (k, i) { st[k] = Number(r[1 + i]) || 65; }); out[String(r[0])] = st; });
  return out;
}

function shared_(a, b) { let s = 0; for (let i = 0; i < a.att.length; i++) s += a.att[i] & (b.att[i] || 0); return s; }

/** Players + attendance from the public attendance sheet (cached 10 min). */
function players_() {
  const c = cache_().get('players');
  if (c) return JSON.parse(c);
  const rows = Utilities.parseCsv(UrlFetchApp.fetch(MAIN_CSV).getContentText());
  let hdr = -1;
  for (let i = 0; i < rows.length; i++) if (rows[i][0] && rows[i][0].toLowerCase().indexOf('games played') >= 0) { hdr = i; break; }
  if (hdr < 0) throw new Error('header not found');
  const NON = ['balance', 'owing from previous sessions', 'images', 'amount paid', 'preferred foot', 'position', 'postion'];
  const cols = [];
  for (let i = SESSION_START_COL; i < rows[hdr].length; i++) { const h = (rows[hdr][i] || '').trim(); if (h && NON.indexOf(h.toLowerCase()) < 0) cols.push(i); }
  const seen = {}, out = [];
  rows.slice(hdr + 1).forEach(function (r) {
    const name = (r[1] || '').trim(), count = parseInt((r[0] || '').replace(/,/g, ''), 10);
    if (!name || isNaN(count) || count <= 0) return;
    const n = name.toLowerCase();
    if (SKIP.indexOf(n) >= 0 || /^\d/.test(n) || n.length < 2) return;
    const att = cols.map(function (i) { const v = parseFloat((r[i] || '').replace(/,/g, '')); return (!isNaN(v) && v > 0) ? 1 : 0; });
    const key = n + '|' + count + '|' + att.join('');
    if (seen[key]) return; seen[key] = 1;
    out.push({ name: name, count: count, position: (r[2] || '').trim(), att: att });
  });
  cache_().put('players', JSON.stringify(out), 600);
  return out;
}
