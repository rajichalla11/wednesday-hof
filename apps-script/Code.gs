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
/* Sheet rows that are the same person: alias (lower case) → the name used everywhere. Keep in step with ALIASES in index.html. */
const ALIASES_ = { 'rj': 'Rahul', 'sahil m.': 'Sahil M' };
const SKIP = ['', 'player', 'games played', 'amount paid', 'balance', 'owing from previous sessions', '+1s', 'krish p', 'images', 'postion', 'position', 'preferred foot'];

/* ───────── admin: run these from the editor ───────── */

function setup() {
  const v = sheet_('Voters');  if (v.getLastRow() === 0) v.appendRow(['voter_id', 'name', 'created', 'claimed', 'status']);
  if (v.getLastColumn() < 4) v.getRange(1, 4).setValue('claimed');
  if (v.getLastColumn() < 5) v.getRange(1, 5).setValue('status');   // pending | approved (blank on an old claimed row = approved)
  if (v.getLastColumn() < 6) v.getRange(1, 6).setValue('rated (rough)');   // 1–4 / 5–9 / 10+ teammates rated (never exact; which ballot is theirs stays unknown)
  settings_();   // creates the Settings tab with a group code if missing
  const b = sheet_('Votes');   if (b.getLastRow() === 0) b.appendRow(['window', 'ballot_id', 'target', 'w'].concat(STATS, ['comment', 'updated', 'traits']));
  if (b.getLastColumn() < 4 + STATS.length + 3) b.getRange(1, 4 + STATS.length + 3).setValue('traits');
  const sr = sheet_('SelfRatings'); if (sr.getLastRow() === 0) sr.appendRow(['window', 'name'].concat(STATS, ['traits', 'note', 'updated']));
  const w = sheet_('Windows'); if (w.getLastRow() === 0) w.appendRow(['window', 'opened', 'closed']);
  const ph = sheet_('Photos'); if (ph.getLastRow() === 0) ph.appendRow(['name', 'file_id', 'updated']);
  photoFolder_();   // creates the Drive folder for player photos (and asks for the Drive permission the first time)
  const s = sheet_('Seeds'); if (s.getLastRow() === 0) s.appendRow(['name'].concat(STATS));   // only names the admin sets get a row; everyone else is decided by votes
  const vi = sheet_('Visits'); if (vi.getLastRow() === 0) vi.appendRow(['ts', 'device', 'name', 'mobile', 'page']);
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
  bust_('Voters');
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
  bust_('Windows');
}
function closeWindow() {
  props_().setProperty('open', '0');
  const w = sheet_('Windows'), data = w.getDataRange().getValues();
  for (let i = data.length - 1; i >= 1; i--) if (data[i][0] === props_().getProperty('window')) { w.getRange(i + 1, 3).setValue(new Date()); break; }
  bust_('Windows');
}

/** Run from the editor after editing Settings, Seeds or Voters by hand, so the change shows at once instead of within 5 minutes. */
function clearCaches() { ['Voters', 'Votes', 'SelfRatings', 'Windows', 'Seeds', 'Settings', 'Photos', 'Visits', 'Checkins', 'MatchVotes', 'Games', 'Signups', 'Taps', 'Strikes', 'TopUps'].forEach(bust_); cache_().remove('players'); }

/** One-off: strips what could link a ballot to a voter from data written before Version 20. Voters keep only a rough
 *  "rated" band (no last-vote time), ballot and check-in times become the day only, and ballot rows are shuffled so
 *  their order says nothing about who voted when. Safe to run again. */
function anonymise() {
  const lock = LockService.getScriptLock(); lock.waitLock(30000);
  try {
    const v = sheet_('Voters'), vd = v.getDataRange().getValues();
    if (vd.length > 1 && vd[0].length >= 6) {
      v.getRange(1, 6, 1, 2).setValues([['rated (rough)', '']]);
      v.getRange(2, 6, vd.length - 1, 2).setValues(vd.slice(1).map(function (r) { const x = r[5]; return ["'" + (typeof x === 'number' ? band_(x) : String(x || '')), '']; }));
    }
    shuffleTab_('Votes', 4 + STATS.length + 1);     // updated column (0-based)
    shuffleTab_('MatchVotes', 5);
    const c = sheet_('Checkins'), cd = c.getDataRange().getValues();
    if (cd.length > 1 && cd[0].length >= 4) c.getRange(2, 4, cd.length - 1, 1).setValues(cd.slice(1).map(function (r) { return [dayOf_(r[3])]; }));
  } finally { lock.releaseLock(); }
  clearCaches();
}
function shuffleTab_(name, tsCol) {
  const sh = sheet_(name), data = sh.getDataRange().getValues();
  if (data.length < 3) return;
  const rows = shuffle_(data.slice(1).map(function (r) { if (tsCol < r.length) r[tsCol] = dayOf_(r[tsCol]); r[0] = r[0] instanceof Date ? r[0] : "'" + r[0]; return r; }));
  sh.getRange(2, 1, rows.length, rows[0].length).setValues(rows);
}
/** Rough count shown to the admin instead of the exact number of teammates a voter has rated. */
function band_(n) { n = Number(n) || 0; return n <= 0 ? '' : n < 5 ? '1–4' : n < 10 ? '5–9' : '10+'; }
/** The day only, as text (no time of day), for anything written next to a ballot or a name. */
function dayOf_(d) { if (!d) return ''; const x = d instanceof Date ? d : new Date(d); return isNaN(x.getTime()) ? String(d).slice(0, 10) : "'" + Utilities.formatDate(x, TZ, 'yyyy-MM-dd'); }
function today_() { return dayOf_(new Date()); }
/** Inserts a row at a random position below the header, so row order can't be matched against site opens. */
function insertRandom_(sh, row) {
  const last = sh.getLastRow();
  if (last < 2) { sh.appendRow(row); return; }
  const at = 2 + Math.floor(Math.random() * last);   // 2 … last+1
  if (at > last) sh.appendRow(row); else { sh.insertRowBefore(at); sh.getRange(at, 1, 1, row.length).setValues([row]); }
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
    if (body.action === 'setFlag') return json_(setFlag_(body));
    if (body.action === 'freeClaim') return json_(freeClaim_(body));
    if (body.action === 'tap') return json_(tap_(body));
    if (body.action === 'signup') return json_(signup_(body));
    if (body.action === 'withdraw') return json_(withdraw_(body));
    if (body.action === 'clearStrikes') return json_(clearStrikes_(body));
    if (body.action === 'addFixture') return json_(addFixture_(body));
    if (body.action === 'removeFixture') return json_(removeFixture_(body));
    if (body.action === 'moveStandby') return json_(moveStandby_(body));
    if (body.action === 'importList') return json_(importList_(body));
    if (body.action === 'editGame') return json_(editGame_(body));
    if (body.action === 'unrate') return json_(unrate_(body));
    if (body.action === 'photo') return json_(photo_(body));
    if (body.action === 'approve') return json_(approve_(body));
    if (body.action === 'unclaim') return json_(unclaim_(body));
    if (body.action === 'clearSeed') return json_(clearSeed_(body));
    if (body.action === 'hit') return json_(hit_(body));
    if (body.action === 'setGame') return json_(setGame_(body));
    if (body.action === 'checkin') return json_(checkin_(body));
    if (body.action === 'matchRate') return json_(matchRate_(body));
    if (body.action === 'setMotm') return json_(setMotm_(body));
    return json_({ ok: false, error: 'unknown action' });
  } catch (err) { return json_({ ok: false, error: String(err) }); }
}

function me_(token) {
  const row = voterRow_(token);
  if (!row) return { ok: false, error: 'This link isn’t recognised.' };
  if (row.status === 'pending') return { ok: true, pending: true, name: row.name };
  const name = row.name;
  const win = props_().getProperty('window'), open = props_().getProperty('open') === '1';
  const mine = {}, feedback = [];
  const b = id2_(token), b3 = id3_(token);
  tab_('Votes').forEach(function (r) {
    if (r[0] !== win) return;
    if (r[1] === b || r[1] === b3) { const st = {}; STATS.forEach(function (k, i) { st[k] = Number(r[4 + i]); }); mine[r[2]] = { stats: st, comment: r[4 + STATS.length] || '', traits: String(r[4 + STATS.length + 2] || '').split(',').filter(Boolean), anon: r[1] === b }; }   // anon: made before names were kept; changing it saves it with the voter's name
    if (r[2] === name && r[4 + STATS.length]) feedback.push(String(r[4 + STATS.length]));
  });
  voterStats_(token);   // keeps the admin's "rated" count current
  const selfRow = tab_('SelfRatings').filter(function (r) { return r[0] === win && String(r[1]).toLowerCase() === name.toLowerCase(); }).pop();
  const self = selfRow ? rowToRating_(selfRow, 2) : null;
  const seeds = seeds_(), hide = hiddenNames_();
  if (!isAdmin_(name)) hide.forEach(function (h) { Object.keys(seeds).forEach(function (n) { if (n.toLowerCase() === h) delete seeds[n]; }); });   // a hidden player's starting point is not sent to voters
  const g = gameInfo_(), gm = g.date ? { date: g.date, time: g.time, gate: g.gate, teams: g.teams, score: g.score, list: g.list, standby: g.standby, size: g.size, in: !!squad_(g.date)[name.toLowerCase()], mine: myMatch_(g.date, token) } : null;
  const lg = lastGame_(), last = lg ? { date: lg.date, mine: myMatch_(lg.date, token) } : null;
  return { ok: true, name: name, window: win, open: open, seeds: seeds, mine: mine, feedback: shuffle_(feedback), minShared: MIN_SHARED, game: gm, last: last,
           traits: TRAITS, maxTraits: MAX_TRAITS, self: self, isAdmin: isAdmin_(name), voteAll: voteAll_(name), topups: topupsFor_(name), potTopups: POTS_[name.toLowerCase()] ? topupsFor_(POTS_[name.toLowerCase()]) : null };
}

/* ───────── top-ups ─────────
   The sheet keeps only each player's running "paid in" total (column F). Every time the sheet is read (at most every 10
   minutes), totals are compared with the last ones seen; a rise is logged in the TopUps tab (ts, name, amount, total).
   The first read only takes the snapshot, so nothing is invented for money paid before tracking began. */
function topupWatch_(players) {
  const lock = LockService.getScriptLock(); if (!lock.tryLock(3000)) return;
  try {
    const pr = props_(), prev = JSON.parse(pr.getProperty('paid_snap') || 'null'), now = {}, rows = [], ts = "'" + new Date().toISOString();
    players.forEach(function (p) { now[p.name] = p.paid || 0; if (prev && now[p.name] > (prev[p.name] || 0) + 0.5) rows.push([ts, p.name, now[p.name] - (prev[p.name] || 0), now[p.name]]); });
    if (rows.length) { const sh = sheet_('TopUps'); if (sh.getLastRow() === 0) sh.appendRow(['ts', 'name', 'amount', 'total']); sh.getRange(sh.getLastRow() + 1, 1, rows.length, 4).setValues(rows); bust_('TopUps'); }
    pr.setProperty('paid_snap', JSON.stringify(now));
    if (!prev) pr.setProperty('paid_since', new Date().toISOString());
  } finally { lock.releaseLock(); }
}
/* Brothers who share one pot: the second name's games come out of the first name's balance. Keep in step with SHARED_POTS in index.html. */
const POTS_ = { 'dhruv': 'Rahul', 'tushar': 'Tanuj', 'meeral': 'Ajul', 'saurabh': 'Sid', 'uday': 'Ashish' };
function topupsFor_(name) {
  const n = String(name || '').toLowerCase(), list = ss_().getSheetByName('TopUps') ? tab_('TopUps').filter(function (r) { return String(r[1]).toLowerCase() === n; }).map(function (r) { return { ts: String(r[0]), amount: Number(r[2]) || 0 }; }) : [];
  return { since: props_().getProperty('paid_since') || '', list: list.slice(-10).reverse() };
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
  const shared = shared_(me, tp), anyone = voteAll_(name);
  if (shared < MIN_SHARED && !anyone) return { ok: false, error: 'You haven’t played enough sessions with ' + tp.name + '.' };
  const stats = STATS.map(function (k) { const v = Math.round(Number((body.stats || {})[k])); if (isNaN(v)) throw new Error('Missing ' + k); return Math.max(RMIN, Math.min(RMAX, v)); });
  const comment = String(body.comment || '').replace(/\s+/g, ' ').trim().slice(0, 280);
  const traits = cleanTraits_(body.traits);
  const w = shared >= 20 ? 1 : shared >= 8 ? 0.8 : 0.6;   // familiarity, bucketed so it can't identify the voter
  const win = props_().getProperty('window'), b = id3_(body.t), old = id2_(body.t);

  const lock = LockService.getScriptLock(); lock.waitLock(10000);
  try {
    const sh = sheet_('Votes'), data = sh.getDataRange().getValues();
    if (data[0].length < 4 + STATS.length + 4 || data[0][4 + STATS.length + 3] !== 'voter') sh.getRange(1, 4 + STATS.length + 4).setValue('voter');
    let rowIdx = -1;
    for (let i = 1; i < data.length; i++) if (data[i][0] === win && data[i][1] === b && data[i][2] === tp.name) { rowIdx = i + 1; break; }
    if (rowIdx < 0) for (let i = 1; i < data.length; i++) if (data[i][0] === win && data[i][1] === old && data[i][2] === tp.name) { rowIdx = i + 1; break; }   // an earlier anonymous rating of the same player is replaced by the named one (the site warns the voter first)
    const row = [win, b, tp.name, w].concat(stats, [comment, today_(), traits.join(','), name]);   // name: admin only, never in public results   // the day only: an exact time could be matched against who opened the site
    if (rowIdx > 0) sh.getRange(rowIdx, 1, 1, row.length).setValues([row]); else insertRandom_(sh, row);
  } finally { lock.releaseLock(); }
  bust_('Votes');   // the voter's rough count is refreshed on their next open (me_), not at the moment they vote
  return { ok: true };
}

function unrate_(body) {
  const name = voterName_(body.t);
  if (!name) return { ok: false, error: 'This link isn’t recognised.' };
  const win = props_().getProperty('window'), b = id3_(body.t), old = id2_(body.t), target = String(body.target || '');
  const sh = sheet_('Votes'), data = sh.getDataRange().getValues();
  for (let i = data.length - 1; i >= 1; i--) if (data[i][0] === win && (data[i][1] === b || data[i][1] === old) && data[i][2] === target) sh.deleteRow(i + 1);
  bust_('Votes');
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
      if (data[i][3]) return { ok: false, error: 'This name has already been claimed. If that was not you, ask in the group for it to be reset.' };
      voters.deleteRow(i + 1);
    }
    const token = newToken_();
    const status = String(settings_().require_approval || '').trim() === '1' ? 'pending' : 'approved';   // Settings → require_approval = 1 to make the admin approve each claim
    voters.appendRow([id1_(token), p.name, new Date(), new Date(), status]);
    cache_().remove(key); bust_('Voters');
    return { ok: true, token: token, name: p.name, pending: status === 'pending' };
  } finally { lock.releaseLock(); }
}

/** Aggregated ratings: calibrated, trimmed, familiarity-weighted means. Cached 2 min. */
function results_() {
  const cached = cache_().get('results');
  if (cached) return JSON.parse(cached);
  const win = props_().getProperty('window'), open = props_().getProperty('open') === '1';
  const windows = tab_('Windows').map(function (r) { return r[0]; });
  const prevWin = windows[windows.indexOf(win) - 1] || null;
  const all = tab_('Votes');
  const seeds = seeds_();
  const hide = hiddenNames_();
  const strip = function (agg) { Object.keys(agg).forEach(function (n) { delete agg[n].vote; delete agg[n].vw; if (hide.indexOf(n.toLowerCase()) >= 0) agg[n] = { n: agg[n].n, traits: agg[n].traits, hidden: true }; }); return agg; };
  const out = { ok: true, window: win, open: open, minVotes: MIN_VOTES, priorWeight: PRIOR_WEIGHT, requireCode: !!String(settings_().claim_code || '').trim(), photos: photos_(), claimed: claimedNames_(), players: strip(aggregate_(all.filter(function (r) { return r[0] === win; }), seeds)), prev: prevWin ? strip(aggregate_(all.filter(function (r) { return r[0] === prevWin; }), seeds)) : {}, prevWindow: prevWin };
  out.game = gamePublic_(); out.motm = motmTally_(); out.games = gamesPublic_();
  const lg = lastGame_(), ng = nextGame_(); out.last = lg ? gamePublic_(lg) : null; out.next = ng ? gamePublic_(ng) : null; out.fixtures = fixturesPublic_(); out.strikes = strikesPublic_();
  out.ballots = tab_('Voters').filter(function (r) { return r[3] && String(r[5] || '') && String(r[5]) !== '0'; }).length;   // people who have voted (a voter can hold an older anonymous ballot and a newer named one, so ballot ids would count them twice)
  cache_().put('results', JSON.stringify(out), 120);
  return out;
}

function aggregate_(rows, seeds) {
  seeds = seeds || {}; const fixed = fixedNames_();
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
    if (n) { entry.vote = {}; entry.vw = {}; }   // votes alone (calibrated, trimmed, weighted) and their total weight; admin only, stripped from public results
    if (base || n >= MIN_VOTES || n) {
      entry.stats = {};
      STATS.forEach(function (k, i) {
        let items = vs.map(function (r) { return { v: Number(r[4 + i]) - 0.5 * offset[r[1]], w: Number(r[3]) || 1 }; }).sort(function (a, b) { return a.v - b.v; });
        if (n >= 5) { const cut = Math.floor(n * 0.2); items = items.slice(cut, items.length - cut); }
        let sum = items.reduce(function (a, x) { return a + x.v * x.w; }, 0), ws = items.reduce(function (a, x) { return a + x.w; }, 0);
        if (n && ws) { entry.vote[k] = Math.round(10 * sum / ws) / 10; entry.vw[k] = Math.round(ws * 100) / 100; }   // one decimal so the site's live preview matches the card exactly
        if (base) { sum += PRIOR_WEIGHT * base[k]; ws += PRIOR_WEIGHT; }   // the baseline is a strong prior; votes trim it
        entry.stats[k] = Math.max(RMIN, Math.min(RMAX, Math.round(sum / ws)));
      });
      if (!base && n < MIN_VOTES) delete entry.stats;   // too few votes to show on a card yet
      if (base && fixed.indexOf(String(t).toLowerCase()) >= 0) { entry.stats = Object.assign({}, base); entry.fixed = true; }   // the admin's numbers, exactly; votes don't move them
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
  bust_('SelfRatings');
  return { ok: true };
}

function admin_(token) {
  const name = voterName_(token);
  if (!name || !isAdmin_(name)) return { ok: false, error: 'Admin only.' };
  const win = props_().getProperty('window');
  const selfRatings = tab_('SelfRatings').filter(function (r) { return r[0] === win; }).map(function (r) { const o = rowToRating_(r, 2); o.name = String(r[1]); return o; });
  const pending = tab_('Voters').filter(function (r) { return String(r[4]) === 'pending'; }).map(function (r) { return { name: String(r[1]), when: r[3] ? Utilities.formatDate(new Date(r[3]), Session.getScriptTimeZone(), 'd MMM, HH:mm') : '' }; });
  const claims = tab_('Voters').filter(function (r) { return r[3]; }).map(function (r) { return { name: String(r[1]), claimed: String(r[3]), status: String(r[4] || 'approved'), rated: typeof r[5] === 'number' ? band_(r[5]) : String(r[5] || ''), anyone: voteAll_(String(r[1])) }; });
  const agg = aggregate_(tab_('Votes').filter(function (r) { return r[0] === win; }), seeds_());   // unstripped, so the admin sees what votes are doing (including hidden players)
  const named = tab_('Votes').filter(function (r) { return r[0] === win && r[4 + STATS.length + 3]; }).map(function (r) { const o = rowToRating_(r, 4); return { voter: String(r[4 + STATS.length + 3]), target: String(r[2]), stats: o.stats, comment: String(r[4 + STATS.length] || ''), traits: String(r[4 + STATS.length + 2] || '').split(',').filter(Boolean) }; });
  const anon = {}; tab_('Votes').forEach(function (r) { if (r[0] === win && !r[4 + STATS.length + 3]) anon[r[2]] = (anon[r[2]] || 0) + 1; });   // earlier anonymous ballots: counts only
  const matchNamed = ss_().getSheetByName('MatchVotes') ? tab_('MatchVotes').filter(function (r) { return r[6]; }).map(function (r) { return { date: String(r[0]).slice(0, 10), voter: String(r[6]), target: String(r[2]), rating: Number(r[3]), motm: !!Number(r[4]) }; }) : [];
  const now = Date.now(), motmGames = ss_().getSheetByName('Games') ? tab_('Games').map(function (r) { return String(r[0]).slice(0, 10); }).filter(function (d) { return /^\d{4}-\d{2}-\d{2}$/.test(d) && kickoff_({ date: d, time: '00:00' }).getTime() <= now; })
    .sort().reverse().slice(0, 6).map(function (d) { const r = tab_('Games').filter(function (x) { return String(x[0]).slice(0, 10) === d; })[0], split = function (v) { return String(v || '').split(',').map(function (x) { return x.trim(); }).filter(Boolean); };
      return { date: d, players: split(r[1]).concat(split(r[2]), split(r[6])), score: r[3] === '' || r[4] === '' ? null : { black: Number(r[3]), white: Number(r[4]) }, votes: motmVotes_(d), auto: motmAuto_(d), override: motmOverride_(d) }; }) : [];
  return { ok: true, window: win, selfRatings: selfRatings, seeds: seeds_(), traits: TRAITS, pending: pending, activity: activity_(), claims: claims, agg: agg, named: named, anon: anon, matchNamed: matchNamed, motmGames: motmGames };
}

/** Admin approves (ok: true) or rejects (ok: false, row deleted so the name can be claimed again) a pending claim. */
function approve_(body) {
  const admin = voterName_(body.t);
  if (!admin || !isAdmin_(admin)) return { ok: false, error: 'Admin only.' };
  const target = String(body.name || '').trim().toLowerCase();
  const sh = sheet_('Voters'), data = sh.getDataRange().getValues();
  for (let i = data.length - 1; i >= 1; i--) {
    if (String(data[i][1]).toLowerCase() !== target || String(data[i][4]) !== 'pending') continue;
    if (body.ok) sh.getRange(i + 1, 5).setValue('approved'); else sh.deleteRow(i + 1);
    bust_('Voters');
    return { ok: true };
  }
  return { ok: false, error: 'No pending claim for that name.' };
}
/** "Not me": the voter deletes their own claim so the name can be claimed by the right person. Their ballot rows stay but are orphaned. */
function unclaim_(body) {
  const row = voterRow_(body.t);
  if (!row) return { ok: false, error: 'This link isn’t recognised.' };
  const id = id1_(String(body.t)), sh = sheet_('Voters'), data = sh.getDataRange().getValues();
  for (let i = data.length - 1; i >= 1; i--) if (data[i][0] === id) sh.deleteRow(i + 1);
  bust_('Voters');
  return { ok: true };
}
function claimedNames_() { return tab_('Voters').filter(function (r) { return r[3]; }).map(function (r) { return String(r[1]); }); }

/** Admin: remove a player's starting point so their rating comes from votes alone. */
function clearSeed_(body) {
  const name = voterName_(body.t);
  if (!name || !isAdmin_(name)) return { ok: false, error: 'Admin only.' };
  const target = String(body.name || '').trim().toLowerCase();
  const sh = sheet_('Seeds'), data = sh.getDataRange().getValues();
  for (let i = data.length - 1; i >= 1; i--) if (String(data[i][0]).toLowerCase() === target) sh.deleteRow(i + 1);
  bust_('Seeds');
  return { ok: true };
}

/** Anonymous open counter: one row per device per visit (the site sends one per browser session; the server allows one per device per minute). */
function hit_(body) {
  const dev = String(body.d || '').replace(/[^\w-]/g, '').slice(0, 40);
  if (!dev) return { ok: false };
  const key = 'hit|' + dev; if (cache_().get(key)) return { ok: true, dup: true }; cache_().put(key, '1', 60);
  sheet_('Visits').appendRow([new Date(), dev, String(body.name || '').slice(0, 60), body.m ? 1 : 0, String(body.p || '').slice(0, 20)]);
  return { ok: true };
}
function activity_() {
  const rows = tab_('Visits'), now = Date.now(), day = 86400000;
  const sum = function (since) { const devs = {}; let opens = 0, mobile = 0; rows.forEach(function (r) { const ts = new Date(r[0]).getTime(); if (ts < since) return; opens++; devs[r[1]] = 1; if (Number(r[3])) mobile++; }); return { opens: opens, devices: Object.keys(devs).length, mobile: opens ? Math.round(100 * mobile / opens) : 0 }; };
  const byName = {};
  rows.forEach(function (r) { const n = String(r[2] || ''); if (!n) return; const ts = new Date(r[0]).getTime(); byName[n] = byName[n] || { name: n, opens: 0, last: 0 }; byName[n].opens++; if (ts > byName[n].last) byName[n].last = ts; });
  const names = Object.keys(byName).map(function (n) { return { name: n, opens: byName[n].opens, last: new Date(byName[n].last).toISOString() }; }).sort(function (a, b) { return b.last < a.last ? -1 : 1; });
  const tabs = {}; if (ss_().getSheetByName('Taps')) tab_('Taps').forEach(function (r) { if (new Date(r[0]).getTime() < now - 7 * day) return; const k = String(r[2] || ''); tabs[k] = tabs[k] || {}; tabs[k][r[1]] = 1; });
  return { today: sum(now - day), week: sum(now - 7 * day), total: sum(0), names: names, tabs: Object.keys(tabs).map(function (k) { return { tab: k, people: Object.keys(tabs[k]).length }; }).sort(function (a, b) { return b.people - a.people; }) };
}
/** Which tabs people open: one row per device per tab per visit (no names). */
function tap_(body) {
  const dev = String(body.d || '').replace(/[^\w-]/g, '').slice(0, 40), tab = String(body.p || '').replace(/[^\w-]/g, '').slice(0, 20);
  if (!dev || !tab) return { ok: false };
  const key = 'tap|' + dev + '|' + tab; if (cache_().get(key)) return { ok: true, dup: true }; cache_().put(key, '1', 600);
  const sh = sheet_('Taps'); if (sh.getLastRow() === 0) sh.appendRow(['ts', 'device', 'tab']);
  sh.appendRow([new Date(), dev, tab]);
  return { ok: true };
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
  bust_('Seeds');
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
    bust_('Photos'); return { ok: true };
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
  bust_('Photos');
  return { ok: true, id: file.getId() };
}
function photoFolder_() {
  const p = props_(), id = p.getProperty('photo_folder');
  if (id) try { return DriveApp.getFolderById(id); } catch (_) {}
  const f = DriveApp.createFolder('W9-10 Photos'); p.setProperty('photo_folder', f.getId()); return f;
}
function photos_() { const o = {}; if (!ss_().getSheetByName('Photos')) return o; tab_('Photos').forEach(function (r) { if (r[1]) o[String(r[0])] = String(r[1]); }); return o; }

/** Players whose rating is shown as "To be determined": Settings → hide_stats, comma-separated names. If the row is missing, the admin's own rating is hidden; add the row with an empty value to show everyone. */
function hiddenNames_() {
  const s = settings_();
  const raw = Object.prototype.hasOwnProperty.call(s, 'hide_stats') ? s.hide_stats : (s.admin_name || 'Raji');
  return String(raw || '').split(',').map(function (x) { return x.trim().toLowerCase(); }).filter(Boolean);
}
/** Players whose card shows their starting point exactly, whatever the votes say: Settings → fixed_stats, comma-separated (none by default). Their traits still come from votes. */
function fixedNames_() { return String(settings_().fixed_stats || '').split(',').map(function (x) { return x.trim().toLowerCase(); }).filter(Boolean); }
/** Admin: free a claimed name so the right player can claim it again (deletes their Voters row; their ballot stays and still counts). */
function freeClaim_(body) {
  const name = voterName_(body.t);
  if (!name || !isAdmin_(name)) return { ok: false, error: 'Admin only.' };
  const target = String(body.name || '').trim().toLowerCase(); if (!target) return { ok: false, error: 'Which name?' };
  if (target === name.toLowerCase()) return { ok: false, error: 'That is your own card.' };
  const sh = sheet_('Voters'), data = sh.getDataRange().getValues(); let n = 0;
  for (let i = data.length - 1; i >= 1; i--) if (String(data[i][1]).trim().toLowerCase() === target) { sh.deleteRow(i + 1); n++; }
  bust_('Voters');
  return n ? { ok: true, freed: n } : { ok: false, error: 'Nobody has claimed that name.' };
}
/** Admin: set one of the display settings from the site (hide_stats, fixed_stats). */
function setFlag_(body) {
  const name = voterName_(body.t);
  if (!name || !isAdmin_(name)) return { ok: false, error: 'Admin only.' };
  const key = String(body.key || ''); if (['hide_stats', 'fixed_stats'].indexOf(key) < 0) return { ok: false, error: 'Unknown setting.' };
  setSetting_(key, "'" + String(body.value || '').slice(0, 300));
  return { ok: true, value: settings_()[key] };
}
/** Players allowed to rate anyone, whatever nights they shared: Settings → vote_all, comma-separated. If the row is missing it defaults to Nirav, whose games mostly predate the tracked sessions. */
function voteAll_(name) {
  const s = settings_();
  const raw = Object.prototype.hasOwnProperty.call(s, 'vote_all') ? s.vote_all : 'Nirav';
  return String(raw || '').split(',').map(function (x) { return x.trim().toLowerCase(); }).filter(Boolean).indexOf(String(name || '').toLowerCase()) >= 0;
}
/* ───────── game day ─────────
   Settings: game_date (yyyy-mm-dd), game_time (HH:mm, Nairobi), gate_code. The gate code only ever goes to approved
   voters through me_, never into the public results. Checkins: date, name, in|out, when. MatchVotes: date, ballot_id,
   target, rating 1–10, motm 0/1, day, voter. Match ratings open at kick-off and close at 11:00 the next morning, when the
   man of the match is settled. Settings always hold the game being set up; the last game played is read back from the
   Games tab, so entering next week's game doesn't cut short the voting on last night's. */
const TZ = 'Africa/Nairobi';
/** game_teams holds { black, white, list, standby }: teams once they're made, or just the list of who's in (plus standby) before that. */
function gameInfo_() {
  const s = settings_(); let raw = null, score = null;
  try { raw = s.game_teams ? JSON.parse(s.game_teams) : null; } catch (_) {}
  const m = String(s.game_score || '').match(/^(\d+)-(\d+)$/); if (m) score = { black: Number(m[1]), white: Number(m[2]) };
  const teams = raw && ((raw.black || []).length || (raw.white || []).length) ? { black: raw.black || [], white: raw.white || [] } : null;
  return { date: String(s.game_date || '').trim().slice(0, 10), time: String(s.game_time || '21:00').trim().slice(0, 5) || '21:00', gate: String(s.gate_code || '').trim(), teams: teams, score: score,
           list: (raw && raw.list) || [], standby: (raw && raw.standby) || [], size: Number(s.game_size) || 6 };
}
const cleanTeam_ = function (a) { return (Array.isArray(a) ? a : []).map(function (x) { return String(x).replace(/\s+/g, ' ').trim().slice(0, 40); }).filter(Boolean).slice(0, 30); };
/** One row per game in the Games tab: date, black names, white names, black goals, white goals. Rewritten whenever the admin saves that game. */
function saveGameRow_(date, teams, score, time, list, size) {
  if (!date || !teams) return;
  const sh = sheet_('Games'); if (sh.getLastRow() === 0) sh.appendRow(['date', 'black', 'white', 'black_goals', 'white_goals', 'time', 'list', 'size']);
  if (sh.getRange(1, 8).getValue() !== 'size') sh.getRange(1, 6, 1, 3).setValues([['time', 'list', 'size']]);
  const row = ["'" + date, teams.black.join(', '), teams.white.join(', '), score ? score.black : '', score ? score.white : '', "'" + (time || '21:00'), (list || []).join(', '), size || 6];
  const data = sh.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) if (String(data[i][0]).slice(0, 10) === date) { sh.getRange(i + 1, 1, 1, 8).setValues([row]); bust_('Games'); return; }
  sh.appendRow(row); bust_('Games');
}
/** Every game in the Games tab: teams or list, score, format, and the man of the match once voting has closed (or Raji picked one). */
function gamesPublic_() {
  if (!ss_().getSheetByName('Games')) return [];
  const now = Date.now(), raters = {}, sums = {};
  if (ss_().getSheetByName('MatchVotes')) tab_('MatchVotes').forEach(function (r) { const d = String(r[0]).slice(0, 10), k = String(r[2]); (raters[d] = raters[d] || {})[r[1]] = 1;
    const g = (sums[d] = sums[d] || {}); g[k] = g[k] || { s: 0, n: 0 }; g[k].s += Number(r[3]); g[k].n++; });
  const avgs = function (d) { const out = {}, g = sums[d] || {}; Object.keys(g).forEach(function (k) { if (g[k].n >= 2) out[k] = { avg: Math.round(10 * g[k].s / g[k].n) / 10, n: g[k].n }; }); return out; };   // the form guide: each player's average from that night (2+ votes)
  return tab_('Games').map(function (r) { const split = function (v) { return String(v || '').split(',').map(function (x) { return x.trim(); }).filter(Boolean); };
    const d = String(r[0]).slice(0, 10), settled = now >= closes_(d) || !!motmOverride_(d);
    return { date: d, black: split(r[1]), white: split(r[2]), b: r[3] === '' ? null : Number(r[3]), w: r[4] === '' ? null : Number(r[4]),
             time: String(r[5] || '21:00').slice(0, 5) || '21:00', list: split(r[6]), size: Number(r[7]) || 6, motm: settled ? motmFor_(d) : [], raters: Object.keys(raters[d] || {}).length, ratings: avgs(d) }; })
    .filter(function (g) { return /^\d{4}-\d{2}-\d{2}$/.test(g.date); });
}
function kickoff_(g) { return new Date(g.date + 'T' + (g.time || '21:00') + ':00+03:00'); }
/** Match voting closes at 11:00 Nairobi time the morning after the game. */
function closes_(date) { return new Date(date + 'T11:00:00+03:00').getTime() + 86400000; }
/** The most recent game that has kicked off: the one in Settings if it has, otherwise the latest row of the Games tab. */
/** Every game: the Settings one (game being set up, with gate code) plus each row of the Games tab, with sign-ups applied. */
function allGames_() {
  const cur = gameInfo_(), out = [];
  if (cur.date) out.push(cur);
  if (ss_().getSheetByName('Games')) tab_('Games').forEach(function (r) {
    const d = String(r[0]).slice(0, 10); if (!/^\d{4}-\d{2}-\d{2}$/.test(d) || d === cur.date) return;
    const split = function (v) { return String(v || '').split(',').map(function (x) { return x.trim(); }).filter(Boolean); };
    const b = split(r[1]), w = split(r[2]);
    out.push({ date: d, time: String(r[5] || '21:00').slice(0, 5) || '21:00', gate: '', teams: b.length || w.length ? { black: b, white: w } : null, score: r[3] === '' || r[4] === '' ? null : { black: Number(r[3]), white: Number(r[4]) }, list: split(r[6]), standby: [], size: Number(r[7]) || 6 });
  });
  const su = signupsByDate_();
  out.forEach(function (g) { const list = su[g.date]; if (!list || !list.length || (g.list && g.list.length)) return;   // a lineup the admin entered for the game wins over sign-ups
    const slots = 2 * (g.size || 6);   // sign-ups on the site are the list: first come, first in; the rest are standby
    g.list = list.slice(0, slots).map(function (x) { return x.name; }); g.standby = list.slice(slots).map(function (x) { return x.name; }); });
  return out;
}
function lastGame_() {
  const now = Date.now();
  return allGames_().filter(function (g) { return kickoff_(g).getTime() <= now; }).sort(function (a, b) { return kickoff_(b).getTime() - kickoff_(a).getTime(); })[0] || null;
}
/** The soonest game still to be played (Settings or a fixture added for sign-ups). */
function nextGame_() {
  const now = Date.now();
  return allGames_().filter(function (g) { return kickoff_(g).getTime() > now; }).sort(function (a, b) { return kickoff_(a).getTime() - kickoff_(b).getTime(); })[0] || null;
}

/* ───────── sign-ups ─────────
   Signups tab: date, ts, name, kind (player|guest), by. Order of ts is the order of the list; the first 2×size are in,
   the rest standby, so dropping out promotes the next standby automatically. Players sign themselves up; only the admin
   adds another player or a guest (shown as "via" whoever brought them). Only you, whoever added you, or the admin can remove you. */
function signupsByDate_() {
  const out = {}; if (!ss_().getSheetByName('Signups')) return out;
  tab_('Signups').forEach(function (r) { const d = String(r[0]).slice(0, 10); if (!/^\d{4}-\d{2}-\d{2}$/.test(d)) return; (out[d] = out[d] || []).push({ name: String(r[2]), kind: String(r[3] || 'player'), by: String(r[4] || ''), ts: String(r[1]) }); });
  Object.keys(out).forEach(function (d) { out[d].sort(function (a, b) { return a.ts < b.ts ? -1 : a.ts > b.ts ? 1 : 0; }); });
  return out;
}
function fixturesPublic_() {
  const now = Date.now(), su = signupsByDate_();
  return allGames_().filter(function (g) { return kickoff_(g).getTime() + 3600000 > now; }).sort(function (a, b) { return kickoff_(a).getTime() - kickoff_(b).getTime(); })
    .map(function (g) { return { date: g.date, time: g.time, size: g.size || 6, slots: 2 * (g.size || 6), teams: !!g.teams, signups: (su[g.date] || []).map(function (x) { return { name: x.name, kind: x.kind, by: x.by }; }) }; });
}
function signupSheet_() { const sh = sheet_('Signups'); if (sh.getLastRow() === 0) sh.appendRow(['date', 'ts', 'name', 'kind', 'by']); return sh; }
function signup_(body) {
  const me = voterName_(body.t);
  if (!me) return { ok: false, error: 'Find your card first.' };
  const date = String(body.date || ''), g = allGames_().filter(function (x) { return x.date === date; })[0];
  if (!g) return { ok: false, error: 'That game is not open.' };
  if (kickoff_(g).getTime() + 3600000 < Date.now()) return { ok: false, error: 'That game has been played.' };
  let name = me, kind = 'player';
  // Players sign up themselves only; adding someone else or a guest (+1) is for the admin
  if ((body.guest || (body.name && String(body.name).toLowerCase() !== me.toLowerCase())) && !isAdmin_(me)) return { ok: false, error: 'You can only sign yourself up. Ask the admin to add anyone else.' };
  if (body.guest) { name = String(body.guest).replace(/\s+/g, ' ').trim().slice(0, 30); kind = 'guest'; if (!name) return { ok: false, error: 'Type the guest’s name.' }; }
  else if (body.name) { const p = players_().filter(function (x) { return x.name.toLowerCase() === String(body.name).toLowerCase(); })[0]; if (!p) return { ok: false, error: 'Unknown player.' }; name = p.name; }
  const lock = LockService.getScriptLock(); lock.waitLock(10000);
  try {
    const sh = signupSheet_(), data = sh.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) if (String(data[i][0]).slice(0, 10) === date && String(data[i][2]).toLowerCase() === name.toLowerCase()) return { ok: false, error: name + ' is already on the list.' };
    sh.appendRow(["'" + date, "'" + new Date().toISOString(), name, kind, name === me && kind === 'player' ? '' : me]);
  } finally { lock.releaseLock(); }
  bust_('Signups');
  return { ok: true, fixtures: fixturesPublic_() };
}
function withdraw_(body) {
  const me = voterName_(body.t);
  if (!me) return { ok: false, error: 'Find your card first.' };
  const date = String(body.date || ''), name = String(body.name || me).toLowerCase(), admin = isAdmin_(me);
  const lock = LockService.getScriptLock(); lock.waitLock(10000); let done = false, strike = null;
  try {
    // Pulling out of a spot (not standby) within 36 hours of kick-off is a strike, unless the admin takes them off
    const g = allGames_().filter(function (x) { return x.date === date; })[0], su = signupsByDate_()[date] || [];
    const idx = su.map(function (x) { return x.name.toLowerCase(); }).indexOf(name), row = su[idx];
    if (g && row && row.kind === 'player' && !admin && idx < 2 * (g.size || 6)) { const left = kickoff_(g).getTime() - Date.now(); if (left < 36 * 3600000 && left > -3600000) strike = row.name; }
    const sh = signupSheet_(), data = sh.getDataRange().getValues();
    for (let i = data.length - 1; i >= 1; i--) {
      if (String(data[i][0]).slice(0, 10) !== date || String(data[i][2]).toLowerCase() !== name) continue;
      if (!(admin || name === me.toLowerCase() || String(data[i][4]).toLowerCase() === me.toLowerCase())) return { ok: false, error: 'Only they, whoever added them, or the admin can take them off.' };
      sh.deleteRow(i + 1); done = true;
    }
    if (done && strike) { const st = strikeSheet_(); st.appendRow(["'" + date, "'" + new Date().toISOString(), strike, me, '']); }
  } finally { lock.releaseLock(); }
  bust_('Signups'); if (done && strike) bust_('Strikes');
  return done ? { ok: true, fixtures: fixturesPublic_(), strikes: strikesPublic_(), strike: strike || '' } : { ok: false, error: 'Not on that list.' };
}
/* ───────── strikes ─────────
   Strikes tab: date (game), ts, name, by (who took them off), cleared. One for each late drop-out (inside 36 hours of
   kick-off, from a spot rather than standby). They only show; nothing is blocked. They stay until the admin clears them. */
function strikeSheet_() { const sh = sheet_('Strikes'); if (sh.getLastRow() === 0) sh.appendRow(['date', 'ts', 'name', 'by', 'cleared']); return sh; }
function strikesPublic_() {
  const out = {}; if (!ss_().getSheetByName('Strikes')) return out;
  tab_('Strikes').forEach(function (r) { if (!r[2] || r[4]) return; const n = String(r[2]); (out[n] = out[n] || []).push(String(r[0]).slice(0, 10)); });
  return out;
}
/** Admin: clear someone's strikes (all of them) */
function clearStrikes_(body) {
  const me = voterName_(body.t); if (!me || !isAdmin_(me)) return { ok: false, error: 'Admin only.' };
  const name = String(body.name || '').toLowerCase(); if (!name || !ss_().getSheetByName('Strikes')) return { ok: false, error: 'No strikes.' };
  const sh = strikeSheet_(), data = sh.getDataRange().getValues(), stamp = "'" + new Date().toISOString();
  for (let i = 1; i < data.length; i++) if (String(data[i][2]).toLowerCase() === name && !data[i][4]) sh.getRange(i + 1, 5).setValue(stamp);
  bust_('Strikes');
  return { ok: true, strikes: strikesPublic_() };
}
/** Admin: open a game for sign-ups (a Games row with no lineup yet), or remove one that has no sign-ups. */
function addFixture_(body) {
  const me = voterName_(body.t); if (!me || !isAdmin_(me)) return { ok: false, error: 'Admin only.' };
  const date = /^\d{4}-\d{2}-\d{2}$/.test(String(body.date || '')) ? String(body.date) : ''; if (!date) return { ok: false, error: 'Pick a date.' };
  const time = /^\d{1,2}:\d{2}$/.test(String(body.time || '')) ? ('0' + body.time).slice(-5) : '21:00', size = Math.max(4, Math.min(11, parseInt(body.size, 10) || 6));
  const sh = sheet_('Games'); if (sh.getLastRow() === 0) sh.appendRow(['date', 'black', 'white', 'black_goals', 'white_goals', 'time', 'list', 'size', 'motm']);
  const data = sh.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) if (String(data[i][0]).slice(0, 10) === date) { sh.getRange(i + 1, 6, 1, 1).setValue("'" + time); sh.getRange(i + 1, 8).setValue(size); bust_('Games'); return { ok: true, fixtures: fixturesPublic_() }; }
  sh.appendRow(["'" + date, '', '', '', '', "'" + time, '', size, '']); bust_('Games');
  return { ok: true, fixtures: fixturesPublic_() };
}
function removeFixture_(body) {
  const me = voterName_(body.t); if (!me || !isAdmin_(me)) return { ok: false, error: 'Admin only.' };
  const date = String(body.date || ''), su = signupsByDate_()[date] || [];
  if (su.length && !body.force) return { ok: false, error: su.length + ' signed up. Remove them first or confirm.' };
  if (su.length) { const s2 = signupSheet_(), d2 = s2.getDataRange().getValues(); for (let i = d2.length - 1; i >= 1; i--) if (String(d2[i][0]).slice(0, 10) === date) s2.deleteRow(i + 1); bust_('Signups'); }
  const sh = sheet_('Games'), data = sh.getDataRange().getValues();
  for (let i = data.length - 1; i >= 1; i--) if (String(data[i][0]).slice(0, 10) === date && !String(data[i][1]) && !String(data[i][2]) && data[i][3] === '') sh.deleteRow(i + 1);
  bust_('Games');
  return { ok: true, fixtures: fixturesPublic_() };
}
/** Admin: fix any game, past or upcoming: kick-off, format, teams or who played, and the score. */
function editGame_(body) {
  const me = voterName_(body.t); if (!me || !isAdmin_(me)) return { ok: false, error: 'Admin only.' };
  const date = /^\d{4}-\d{2}-\d{2}$/.test(String(body.date || '')) ? String(body.date) : ''; if (!date) return { ok: false, error: 'Which game?' };
  const time = /^\d{1,2}:\d{2}$/.test(String(body.time || '')) ? ('0' + body.time).slice(-5) : '21:00', size = Math.max(4, Math.min(11, parseInt(body.size, 10) || 6));
  const teams = { black: cleanTeam_((body.teams || {}).black), white: cleanTeam_((body.teams || {}).white) }, list = cleanTeam_(body.list);
  const b = parseInt(body.scoreBlack, 10), w = parseInt(body.scoreWhite, 10), score = !isNaN(b) && !isNaN(w) ? { black: b, white: w } : null;
  const sh = sheet_('Games'); if (sh.getLastRow() === 0) sh.appendRow(['date', 'black', 'white', 'black_goals', 'white_goals', 'time', 'list', 'size', 'motm']);
  const row = [teams.black.join(', '), teams.white.join(', '), score ? score.black : '', score ? score.white : '', "'" + time, list.join(', '), size];
  const data = sh.getDataRange().getValues(); let at = -1;
  for (let i = 1; i < data.length; i++) if (String(data[i][0]).slice(0, 10) === date) { at = i + 1; break; }
  if (at < 0) { sh.appendRow(["'" + date].concat(row)); } else sh.getRange(at, 2, 1, row.length).setValues([row]);
  bust_('Games');
  const cur = gameInfo_();
  if (cur.date === date) {   // keep the game-day settings in step when it's the game being set up
    setSetting_('game_time', "'" + time); setSetting_('game_size', size);
    setSetting_('game_teams', teams.black.length || teams.white.length || list.length ? JSON.stringify({ black: teams.black, white: teams.white, list: list, standby: cur.standby || [] }) : '');
    setSetting_('game_score', score ? "'" + score.black + '-' + score.white : '');
  }
  return { ok: true, games: gamesPublic_() };
}
/** Admin: load a list that was taken on WhatsApp, in order: entries [{ name, kind, by }]. Names already on the list are skipped. */
function importList_(body) {
  const me = voterName_(body.t); if (!me || !isAdmin_(me)) return { ok: false, error: 'Admin only.' };
  const date = String(body.date || ''); if (!allGames_().some(function (g) { return g.date === date; })) return { ok: false, error: 'Open the game first.' };
  const entries = Array.isArray(body.entries) ? body.entries.slice(0, 40) : [];
  const lock = LockService.getScriptLock(); lock.waitLock(10000); let n = 0;
  try {
    const sh = signupSheet_(), data = sh.getDataRange().getValues(), have = {};
    for (let i = 1; i < data.length; i++) if (String(data[i][0]).slice(0, 10) === date) have[String(data[i][2]).toLowerCase()] = 1;
    const base = Date.now();
    entries.forEach(function (e, k) { const name = String(e.name || '').replace(/\s+/g, ' ').trim().slice(0, 30); if (!name || have[name.toLowerCase()]) return; have[name.toLowerCase()] = 1;
      sh.appendRow(["'" + date, "'" + new Date(base + k).toISOString(), name, e.kind === 'guest' ? 'guest' : 'player', String(e.by || '').slice(0, 30)]); n++; });
  } finally { lock.releaseLock(); }
  bust_('Signups');
  return { ok: true, added: n, fixtures: fixturesPublic_() };
}
/** Admin: move the standby of one game to the end of another game's list, in their order. */
function moveStandby_(body) {
  const me = voterName_(body.t); if (!me || !isAdmin_(me)) return { ok: false, error: 'Admin only.' };
  const from = String(body.from || ''), to = String(body.to || ''), games = fixturesPublic_(), gf = games.filter(function (x) { return x.date === from; })[0], gt = games.filter(function (x) { return x.date === to; })[0];
  if (!gf || !gt) return { ok: false, error: 'Pick two open games.' };
  const moving = gf.signups.slice(gf.slots), already = {}; gt.signups.forEach(function (x) { already[x.name.toLowerCase()] = 1; });
  const lock = LockService.getScriptLock(); lock.waitLock(10000);
  try {
    const sh = signupSheet_(), data = sh.getDataRange().getValues(), names = {}; moving.forEach(function (x) { names[x.name.toLowerCase()] = 1; });
    for (let i = data.length - 1; i >= 1; i--) if (String(data[i][0]).slice(0, 10) === from && names[String(data[i][2]).toLowerCase()]) sh.deleteRow(i + 1);
    const base = Date.now();
    moving.forEach(function (x, k) { if (!already[x.name.toLowerCase()]) sh.appendRow(["'" + to, "'" + new Date(base + k).toISOString(), x.name, x.kind, x.by]); });
  } finally { lock.releaseLock(); }
  bust_('Signups');
  return { ok: true, moved: moving.length, fixtures: fixturesPublic_() };
}
/** Man of the match for one night: most stars; a tie goes to the better match rating; still tied, they share it. */
/** Man of the match for one night: Raji's pick if he set one (Games tab, column 9), otherwise the vote. */
function motmFor_(date) { return motmOverride_(date) || motmAuto_(date); }
function motmOverride_(date) {
  if (!ss_().getSheetByName('Games')) return null;
  const row = tab_('Games').filter(function (r) { return String(r[0]).slice(0, 10) === date; })[0], v = row ? String(row[8] || '').trim() : '';
  return v ? v.split(',').map(function (x) { return x.trim(); }).filter(Boolean) : null;
}
function motmVotes_(date) {
  const by = {};
  tab_('MatchVotes').forEach(function (r) { if (String(r[0]).slice(0, 10) !== date) return; const k = String(r[2]); by[k] = by[k] || { sum: 0, n: 0, stars: 0 }; by[k].sum += Number(r[3]); by[k].n++; by[k].stars += Number(r[4]) ? 1 : 0; });
  const out = {}; Object.keys(by).forEach(function (k) { out[k] = { stars: by[k].stars, n: by[k].n, avg: Math.round(10 * by[k].sum / by[k].n) / 10 }; });
  return out;
}
/** Admin: set (names) or clear (empty) the man of the match for a game, overriding the vote. */
function setMotm_(body) {
  const name = voterName_(body.t);
  if (!name || !isAdmin_(name)) return { ok: false, error: 'Admin only.' };
  const date = /^\d{4}-\d{2}-\d{2}$/.test(String(body.date || '')) ? String(body.date) : '';
  if (!date) return { ok: false, error: 'Which game?' };
  const names = cleanTeam_(body.names);
  const sh = sheet_('Games'); if (sh.getLastRow() === 0) sh.appendRow(['date', 'black', 'white', 'black_goals', 'white_goals', 'time', 'list', 'size', 'motm']);
  if (sh.getRange(1, 9).getValue() !== 'motm') sh.getRange(1, 9).setValue('motm');
  const data = sh.getDataRange().getValues(); let at = -1;
  for (let i = 1; i < data.length; i++) if (String(data[i][0]).slice(0, 10) === date) { at = i + 1; break; }
  if (at < 0) { sh.appendRow(["'" + date, '', '', '', '', '', '', '', '']); at = sh.getLastRow(); }
  sh.getRange(at, 9).setValue(names.join(', '));
  bust_('Games');
  return { ok: true, motm: motmFor_(date), auto: motmAuto_(date) };
}
/** By the vote: most stars; a tie goes to the better match rating; still tied, they share it. */
function motmAuto_(date) {
  const by = {};
  tab_('MatchVotes').forEach(function (r) { if (String(r[0]).slice(0, 10) !== date) return; const k = String(r[2]); by[k] = by[k] || { sum: 0, n: 0, stars: 0 }; by[k].sum += Number(r[3]); by[k].n++; by[k].stars += Number(r[4]) ? 1 : 0; });
  const ks = Object.keys(by).filter(function (k) { return by[k].stars > 0; }); if (!ks.length) return [];
  const top = Math.max.apply(null, ks.map(function (k) { return by[k].stars; })), tied = ks.filter(function (k) { return by[k].stars === top; });
  const avg = function (k) { return Math.round(10 * by[k].sum / by[k].n) / 10; }, best = Math.max.apply(null, tied.map(avg));
  return tied.filter(function (k) { return avg(k) === best; });
}
function setSetting_(key, value) {
  const sh = sheet_('Settings'), data = sh.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) if (String(data[i][0]) === key) { sh.getRange(i + 1, 2).setValue(value); bust_('Settings'); return; }
  sh.appendRow([key, value]); bust_('Settings');
}
function setGame_(body) {
  const name = voterName_(body.t);
  if (!name || !isAdmin_(name)) return { ok: false, error: 'Admin only.' };
  const date = /^\d{4}-\d{2}-\d{2}$/.test(String(body.date || '')) ? String(body.date) : '';
  const time = /^\d{1,2}:\d{2}$/.test(String(body.time || '')) ? ('0' + body.time).slice(-5) : '21:00';
  setSetting_('game_date', date ? "'" + date : '');
  setSetting_('game_time', "'" + time);
  setSetting_('gate_code', "'" + String(body.gate || '').trim().slice(0, 24));
  const teams = body.teams ? { black: cleanTeam_(body.teams.black), white: cleanTeam_(body.teams.white) } : { black: [], white: [] };
  const list = cleanTeam_(body.list), standby = cleanTeam_(body.standby), size = Math.max(4, Math.min(11, parseInt(body.size, 10) || 6));
  const any = teams.black.length || teams.white.length || list.length || standby.length;
  setSetting_('game_teams', any ? JSON.stringify({ black: teams.black, white: teams.white, list: list, standby: standby }) : '');
  setSetting_('game_size', size);
  const b = parseInt(body.scoreBlack, 10), w = parseInt(body.scoreWhite, 10), score = !isNaN(b) && !isNaN(w) ? { black: b, white: w } : null;
  setSetting_('game_score', score ? "'" + b + '-' + w : '');
  if (date && (teams.black.length || teams.white.length || list.length)) saveGameRow_(date, teams, score, time, list, size);
  return { ok: true, game: gameInfo_() };
}
function squad_(date) {
  const out = {};
  tab_('Checkins').forEach(function (r) { if (String(r[0]).slice(0, 10) !== date) return; const k = String(r[1]).toLowerCase(); if (String(r[2]) === 'out') delete out[k]; else out[k] = String(r[1]); });
  return out;
}
function setCheckin_(date, who, status) {
  const sh = sheet_('Checkins'); if (sh.getLastRow() === 0) sh.appendRow(['date', 'name', 'status', 'updated']);
  const data = sh.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) if (String(data[i][0]).slice(0, 10) === date && String(data[i][1]).toLowerCase() === who.toLowerCase()) { sh.getRange(i + 1, 3, 1, 2).setValues([[status, today_()]]); return; }
  sh.appendRow(["'" + date, who, status, today_()]);
}
function checkin_(body) {
  const name = voterName_(body.t);
  if (!name) return { ok: false, error: 'This link isn’t recognised.' };
  const g = gameInfo_(); if (!g.date) return { ok: false, error: 'No game set.' };
  let who = name;
  if (body.name) { const p = players_().filter(function (x) { return x.name.toLowerCase() === String(body.name).toLowerCase(); })[0]; if (!p) return { ok: false, error: 'Unknown player.' }; who = p.name; }
  const lock = LockService.getScriptLock(); lock.waitLock(10000);
  try { setCheckin_(g.date, who, body.out ? 'out' : 'in'); } finally { lock.releaseLock(); }
  bust_('Checkins');
  return { ok: true, squad: Object.keys(squad_(g.date)).map(function (k) { return squad_(g.date)[k]; }) };
}
function matchRate_(body) {
  const name = voterName_(body.t);
  if (!name) return { ok: false, error: 'This link isn’t recognised.' };
  const g = lastGame_();
  if (!g) return { ok: false, error: nextGame_() ? 'Match ratings open at kick-off.' : 'No game set.' };
  if (Date.now() >= closes_(g.date)) return { ok: false, error: 'Voting for this game closed at 11am.' };
  const names = {}; players_().forEach(function (p) { names[p.name.toLowerCase()] = p.name; });
  const ratings = body.ratings || {}, motm = String(body.motm || '').toLowerCase(), b = id3_(body.t), rows = [];
  Object.keys(ratings).forEach(function (k) {
    const who = names[k.toLowerCase()], v = Math.round(Number(ratings[k]) * 2) / 2;
    if (!who || who.toLowerCase() === name.toLowerCase() || isNaN(v) || v < 1 || v > 10) return;
    rows.push([g.date, b, who, v, who.toLowerCase() === motm ? 1 : 0, today_(), name]);
  });
  const lock = LockService.getScriptLock(); lock.waitLock(10000);
  try {
    const sh = sheet_('MatchVotes'); if (sh.getLastRow() === 0) sh.appendRow(['date', 'ballot_id', 'target', 'rating', 'motm', 'updated', 'voter']);
    if (sh.getRange(1, 7).getValue() !== 'voter') sh.getRange(1, 7).setValue('voter');
    const data = sh.getDataRange().getValues();
    for (let i = data.length - 1; i >= 1; i--) if (String(data[i][0]).slice(0, 10) === g.date && data[i][1] === b) sh.deleteRow(i + 1);   // a resubmission replaces the old one
    rows.forEach(function (r) { insertRandom_(sh, ["'" + r[0]].concat(r.slice(1))); });
    const sq = squad_(g.date);
    if (!sq[name.toLowerCase()]) setCheckin_(g.date, name, 'in');                 // rating tonight means you played tonight
    rows.forEach(function (r) { if (!sq[r[2].toLowerCase()]) setCheckin_(g.date, r[2], 'in'); });
  } finally { lock.releaseLock(); }
  bust_('MatchVotes'); bust_('Checkins');
  return { ok: true, saved: rows.length };
}
function myMatch_(date, token) {
  const b = id3_(String(token)), out = { ratings: {}, motm: '' };
  tab_('MatchVotes').forEach(function (r) { if (String(r[0]).slice(0, 10) !== date || r[1] !== b) return; out.ratings[String(r[2])] = Number(r[3]); if (Number(r[4])) out.motm = String(r[2]); });
  return out;
}
function gamePublic_(g) {
  g = g || gameInfo_(); if (!g || !g.date) return null;
  const sq = squad_(g.date), by = {}, raters = {};
  tab_('MatchVotes').forEach(function (r) {
    if (String(r[0]).slice(0, 10) !== g.date) return; raters[r[1]] = 1;
    const k = String(r[2]); by[k] = by[k] || { sum: 0, n: 0, motm: 0 }; by[k].sum += Number(r[3]); by[k].n++; by[k].motm += Number(r[4]) ? 1 : 0;
  });
  const ratings = {}; Object.keys(by).forEach(function (k) { ratings[k] = { avg: Math.round(10 * by[k].sum / by[k].n) / 10, n: by[k].n, motm: by[k].motm }; });
  (g.teams ? g.teams.black.concat(g.teams.white) : []).concat(g.list || []).forEach(function (n) { if (!sq[n.toLowerCase()]) sq[n.toLowerCase()] = n; });
  const closes = closes_(g.date), final = Date.now() >= closes;
  return { date: g.date, time: g.time, squad: Object.keys(sq).map(function (k) { return sq[k]; }), teams: g.teams, score: g.score, ratings: ratings, raters: Object.keys(raters).length, hasGate: !!g.gate,
           closes: new Date(closes).toISOString(), final: final, motm: motmFor_(g.date), motmPicked: !!motmOverride_(g.date), list: g.list || [], standby: g.standby || [], size: g.size || 6 };
}
/** Man of the match winners across every game: most MOTM votes on the night (at least 2), ties share it. */
function motmTally_() {
  if (!ss_().getSheetByName('MatchVotes')) return {};
  const dates = {}, out = {}, now = Date.now();
  tab_('MatchVotes').forEach(function (r) { dates[String(r[0]).slice(0, 10)] = 1; });
  if (ss_().getSheetByName('Games')) tab_('Games').forEach(function (r) { if (String(r[8] || '').trim()) dates[String(r[0]).slice(0, 10)] = 1; });   // a night Raji picked counts even without votes
  Object.keys(dates).forEach(function (d) { if (now < closes_(d)) return; motmFor_(d).forEach(function (k) { out[k] = (out[k] || 0) + 1; }); });   // a night counts once its voting has closed
  return out;
}

function isAdmin_(name) { return String(name || '').toLowerCase() === String(settings_().admin_name || 'Raji').toLowerCase(); }
function cleanTraits_(arr) { const seen = {}; return (Array.isArray(arr) ? arr : []).map(function (t) { return String(t).toUpperCase(); }).filter(function (t) { if (TRAITS.indexOf(t) < 0 || seen[t]) return false; seen[t] = 1; return true; }).slice(0, MAX_TRAITS); }
function rowToRating_(r, off) { const st = {}; STATS.forEach(function (k, i) { st[k] = Number(r[off + i]); }); return { stats: st, traits: String(r[off + STATS.length] || '').split(',').filter(Boolean), note: String(r[off + STATS.length + 1] || ''), updated: r[off + STATS.length + 2] ? String(r[off + STATS.length + 2]) : '' }; }

/* ───────── helpers ───────── */

/* Tab reads are the slow part of every call (300–600 ms each), so each tab is cached as JSON for 5 minutes, in 90 KB chunks, and busted by the write that changes it. Dates come back as ISO strings. */
const ROWS_TTL = 300;
function tab_(name) {
  const c = cache_(), meta = c.get('rows|' + name);
  if (meta) {
    const n = Number(meta), keys = []; for (let i = 0; i < n; i++) keys.push('rows|' + name + '|' + i);
    const parts = c.getAll(keys); let str = '';
    for (let i = 0; i < n; i++) { const k = 'rows|' + name + '|' + i; if (parts[k] == null) { str = null; break; } str += parts[k]; }
    if (str !== null) return JSON.parse(str);
  }
  const rows = rows_(sheet_(name)).map(function (r) { return r.map(function (v) { return v instanceof Date ? v.toISOString() : v; }); });
  const str = JSON.stringify(rows), size = 90000, n = Math.max(1, Math.ceil(str.length / size)), obj = {};
  for (let i = 0; i < n; i++) obj['rows|' + name + '|' + i] = str.slice(i * size, (i + 1) * size);
  obj['rows|' + name] = String(n);
  try { c.putAll(obj, ROWS_TTL); } catch (_) {}
  return rows;
}
function bust_(name) {
  SpreadsheetApp.flush();   // Sheets batches writes until the script ends; without this, a request landing in between could re-cache the old rows for 5 minutes
  const c = cache_(), meta = c.get('rows|' + name);
  if (meta) { const keys = ['rows|' + name]; for (let i = 0; i < Number(meta); i++) keys.push('rows|' + name + '|' + i); c.removeAll(keys); }
  c.remove('results');
}

function ss_() { return SpreadsheetApp.getActive(); }
function sheet_(n) { return ss_().getSheetByName(n) || ss_().insertSheet(n); }
function rows_(sh) { const d = sh.getDataRange().getValues(); return d.slice(1).filter(function (r) { return String(r[0]) !== ''; }); }
function props_() { return PropertiesService.getScriptProperties(); }
function cache_() { return CacheService.getScriptCache(); }
function json_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
function sha_(s) { return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, s, Utilities.Charset.UTF_8).map(function (b) { return ('0' + (b & 255).toString(16)).slice(-2); }).join(''); }
function id1_(t) { return sha_('voter|' + t); }
function id2_(t) { return sha_('ballot|' + t); }
/** Ballot id from 14 Oct 2026 (Version 21) on: votes are saved with the voter's name, which only the admin sees. A different
 *  hash from id2_, so a named vote can't be tied back to the same person's earlier anonymous ballot. */
function id3_(t) { return sha_('named|' + t); }
function newToken_() { return Utilities.getUuid().replace(/-/g, '') + Utilities.getUuid().replace(/-/g, '').slice(0, 8); }
function settings_() {
  const sh = sheet_('Settings');
  if (sh.getLastRow() === 0) { sh.appendRow(['key', 'value']); sh.appendRow(['claim_code', '']); }   // leave blank for no code
  const o = {}; tab_('Settings').forEach(function (r) { o[String(r[0])] = String(r[1]); });
  if (!o.admin_name) { sh.appendRow(['admin_name', 'Raji']); o.admin_name = 'Raji'; bust_('Settings'); }
  return o;
}
function defaultWindow_() { const d = new Date(); return d.getFullYear() + '-Q' + (Math.floor(d.getMonth() / 3) + 1); }
function shuffle_(a) { for (let i = a.length - 1; i > 0; i--) { const j = Math.floor(Math.random() * (i + 1)); const t = a[i]; a[i] = a[j]; a[j] = t; } return a; }

/** The Voters row for a token: { name, status }. An old claimed row with no status counts as approved. */
function voterRow_(token) {
  if (!token || String(token).length < 16) return null;
  const id = id1_(String(token));
  const hit = tab_('Voters').filter(function (r) { return r[0] === id; })[0];
  if (!hit) return null;
  return { name: String(hit[1]), status: String(hit[4] || '') === 'pending' ? 'pending' : 'approved' };
}
/** Writes a rough band (1–4, 5–9, 10+) of how many teammates this token's ballot has rated onto the voter's own row; never the exact number or a time, so a ballot can't be matched to a voter by watching counts change. */
function voterStats_(token) {
  if (!token) return;
  const win = props_().getProperty('window'), b = id2_(String(token)), b3 = id3_(String(token)), id = id1_(String(token));
  const n = tab_('Votes').filter(function (r) { return r[0] === win && (r[1] === b || r[1] === b3); }).length;
  const row = tab_('Voters').filter(function (r) { return r[0] === id; })[0];
  if (!row || String(row[5] || '') === band_(n)) return;
  const sh = sheet_('Voters'), data = sh.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) if (data[i][0] === id) { sh.getRange(i + 1, 6, 1, 2).setValues([["'" + band_(n), '']]); break; }
  bust_('Voters');
}
/** Approved voter name for a token, or null. Everything except me_ goes through this, so a pending claim can't act. */
function voterName_(token) { const r = voterRow_(token); return r && r.status === 'approved' ? r.name : null; }

function seeds_() {
  const out = {};
  tab_('Seeds').forEach(function (r) { const st = {}; STATS.forEach(function (k, i) { st[k] = Number(r[1 + i]) || 65; }); out[String(r[0])] = st; });
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
    const raw = (r[1] || '').trim(), count = parseInt((r[0] || '').replace(/,/g, ''), 10);
    if (!raw || isNaN(count) || count <= 0) return;
    const name = ALIASES_[raw.toLowerCase()] || raw, n = name.toLowerCase();   // rows that are the same person fold into one (attendance OR'd, counts added)
    if (SKIP.indexOf(n) >= 0 || /^\d/.test(n) || n.length < 2) return;
    const att = cols.map(function (i) { const v = parseFloat((r[i] || '').replace(/,/g, '')); return (!isNaN(v) && v > 0) ? 1 : 0; });
    const paid = parseFloat((r[5] || '').replace(/,/g, '')) || 0;
    if (seen[n]) { const m = seen[n]; m.att = m.att.map(function (v, i) { return v | att[i]; }); m.count += count; m.paid += paid; return; }   // same name twice = one person
    const p = { name: name, count: count, position: (r[2] || '').trim(), att: att, paid: paid }; seen[n] = p; out.push(p);
  });
  try { topupWatch_(out); } catch (e) { console.error(e); }
  cache_().put('players', JSON.stringify(out), 600);
  return out;
}
