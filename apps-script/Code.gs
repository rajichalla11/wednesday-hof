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
const ALIASES_ = { 'rj': 'Rahul', 'sahil m.': 'Sahil M', 'veer': 'Vali' };
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
function clearCaches() { ['Voters', 'Votes', 'SelfRatings', 'Windows', 'Seeds', 'Settings', 'Photos', 'Visits', 'Checkins', 'MatchVotes', 'Games', 'Signups', 'Taps', 'Strikes', 'TopUps', 'FootClaims', 'FootVotes', 'Profiles', 'Push', 'PushLog'].forEach(bust_); cache_().remove('players'); }

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
    if (body.action === 'setProfile') return json_(setProfile_(body));
    if (body.action === 'footVote') return json_(footVote_(body));
    if (body.action === 'setAttrs') return json_(setAttrs_(body));
    if (body.action === 'pushSub') return json_(pushSub_(body));
    if (body.action === 'pushUnsub') return json_(pushUnsub_(body));
    if (body.action === 'pushPrefs') return json_(pushPrefs_(body));
    if (body.action === 'pushConfig') return json_(pushConfig_(body));
    if (body.action === 'pushTest') return json_(pushTest_(body));
    if (body.action === 'pushBroadcast') return json_(pushBroadcast_(body));
    if (body.action === 'pushStrikes') return json_(pushStrikesNow_(body));
    if (body.action === 'linkCode') return json_(linkCode_(body));
    if (body.action === 'linkRedeem') return json_(linkRedeem_(body));
    if (body.action === 'renameSignup') return json_(renameSignup_(body));
    if (body.action === 'sheetAdopt') return json_(sheetAdopt_(body));
    if (body.action === 'adminPlayer') return json_(adminPlayer_(body));
    if (body.action === 'setGate') return json_(setGate_(body));
    if (body.action === 'addStrike') return json_(addStrike_(body));
    if (body.action === 'ackStrikes') return json_(ackStrikes_(body));
    if (body.action === 'addTopup') return json_(addTopup_(body));
    if (body.action === 'addMoney') return json_(addMoney_(body));
    if (body.action === 'removeTopup') return json_(removeTopup_(body));
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
           traits: TRAITS, maxTraits: MAX_TRAITS, self: self, isAdmin: isAdmin_(name), voteAll: voteAll_(name), topups: topupsFor_(name), footVotes: myFootVotes_(name), push: myPush_(name), strikes: myStrikes_(name), gates: fixturesPublic_().reduce(function (o, f) { const gc = gateFor_(f.date); if (gc) o[f.date] = gc; return o; }, {}), potTopups: POTS_[name.toLowerCase()] ? topupsFor_(POTS_[name.toLowerCase()]) : null };
}

/* ───────── top-ups ─────────
   The sheet keeps only each player's running "paid in" total (column F). Every time the sheet is read (at most every 10
   minutes), totals are compared with the last ones seen; a rise is logged in the TopUps tab (ts, name, amount, total).
   The first read only takes the snapshot, so nothing is invented for money paid before tracking began. */
function topupWatch_(players) {
  const lock = LockService.getScriptLock(); if (!lock.tryLock(500)) return;
  let told = [];
  try {
    const pr = props_(), prev = JSON.parse(pr.getProperty('paid_snap') || 'null'), now = {}, rows = [], ts = "'" + new Date().toISOString();
    const skip = JSON.parse(pr.getProperty('paid_skip') || '{}');   // money the site already added and logged (Add money), waiting for the CSV to catch up
    players.forEach(function (p) { now[p.name] = p.paid || 0; if (!prev) return;
      let inc = now[p.name] - (prev[p.name] || 0); if (inc <= 0.5) return;
      const sk = Math.min(inc, skip[p.name] || 0); inc -= sk; if (sk) { skip[p.name] -= sk; if (skip[p.name] < 0.5) delete skip[p.name]; }
      if (inc > 0.5) rows.push([ts, p.name, inc, now[p.name], p]); });
    if (rows.length) { const sh = sheet_('TopUps'); if (sh.getLastRow() === 0) sh.appendRow(['ts', 'name', 'amount', 'total']); sh.getRange(sh.getLastRow() + 1, 1, rows.length, 4).setValues(rows.map(function (r) { return r.slice(0, 4); })); bust_('TopUps'); }
    pr.setProperty('paid_snap', JSON.stringify(now)); pr.setProperty('paid_skip', JSON.stringify(skip));
    told = rows;
    if (!prev) pr.setProperty('paid_since', new Date().toISOString());
  } finally { lock.releaseLock(); }
  told.forEach(function (r) { pushTopup_(r[1], r[2], r[4].bal, players); });
}
/* Brothers who share one pot: the second name's games come out of the first name's balance. Keep in step with SHARED_POTS in index.html. */
const POTS_ = { 'dhruv': 'Rahul', 'tushar': 'Tanuj', 'meeral': 'Ajul', 'saurabh': 'Sid', 'uday': 'Ashish' };
function topupsFor_(name) {
  const n = String(name || '').toLowerCase(), list = ss_().getSheetByName('TopUps') ? tab_('TopUps').filter(function (r) { return String(r[1]).toLowerCase() === n; }).map(function (r) { return { ts: String(r[0]), amount: Number(r[2]) || 0 }; }) : [];
  list.sort(function (a, b) { return a.ts < b.ts ? 1 : a.ts > b.ts ? -1 : 0; });
  return { since: props_().getProperty('paid_since') || '', list: list.slice(0, 10) };
}
/** Tell whoever shares this pot that money went in. */
function pushTopup_(holder, amount, bal, list) {
  const ps = list || players_();
  const names = [holder].concat(Object.keys(POTS_).filter(function (k) { return POTS_[k].toLowerCase() === String(holder).toLowerCase(); }).map(function (k) { return ps.filter(function (p) { return p.name.toLowerCase() === k; }).map(function (p) { return p.name; })[0] || ''; }).filter(Boolean));
  const b = Number(bal), now = isNaN(b) || bal === null || bal === '' ? '' : ' Balance now ' + (b < 0 ? '−' : '') + 'KES ' + Math.abs(Math.round(b)).toLocaleString('en') + '.';
  pushSafe_(names, 'topup', 'KES ' + Math.round(amount).toLocaleString('en') + ' added to your balance', 'Thanks for topping up.' + now, '/#home');
}
/** Admin: add money to a player's balance from the site. Goes into their "Paid in" cell on the attendance sheet (the shared
 *  pot's row for brothers), is logged as a top-up, and they get a notification. With dry, says what it would change. */
function addMoney_(body) {
  const me = voterName_(body.t); if (!me || !isAdmin_(me)) return { ok: false, error: 'Admin only.' };
  const amount = Math.round(Number(body.amount) || 0); if (amount <= 0 || amount > 200000) return { ok: false, error: 'Amount?' };
  const ps = players_(), p = ps.filter(function (x) { return x.name.toLowerCase() === String(body.name || '').toLowerCase(); })[0]; if (!p) return { ok: false, error: 'Unknown player.' };
  const holder = POTS_[p.name.toLowerCase()] || p.name;
  const lock = LockService.getScriptLock(); if (!lock.tryLock(20000)) return { ok: false, error: 'Busy, try again.' };
  try {
    const wide = SpreadsheetApp.openById(MAIN_SHEET_ID).getSheetByName('Wednesday Group'); if (!wide) return { ok: false, error: 'Sheet tab not found.' };
    const names = wide.getRange(4, 2, wide.getLastRow() - 3, 1).getValues().map(function (r) { const n = String(r[0] || '').trim().toLowerCase(); return ALIASES_[n] ? ALIASES_[n].toLowerCase() : n; });
    const rows = []; names.forEach(function (n, i) { if (n === holder.toLowerCase()) rows.push(i + 4); });
    if (rows.length !== 1) return { ok: false, error: rows.length ? holder + ' is on the sheet more than once. Add it there by hand.' : holder + ' is not on the sheet.' };
    const row = rows[0], F = wide.getRange(row, 6), G = wide.getRange(row, 7), fF = F.getFormula(), vF = F.getValue(), fG = G.getFormula();
    const before = Number(String(vF).replace(/,/g, '')) || 0;
    if (body.dry) return { ok: true, dry: true, holder: holder, row: row, paidFormula: fF, paidValue: vF, balanceFormula: fG, balanceValue: G.getValue() };
    if (!fG) return { ok: false, error: "The balance cell for " + holder + " isn't a formula, so it wouldn't update. Add it on the sheet by hand." };
    if (fF) F.setFormula(fF + '+' + amount); else if (vF === '' || !isNaN(Number(String(vF).replace(/,/g, '')))) F.setValue(before + amount); else return { ok: false, error: 'The paid-in cell for ' + holder + ' has text in it. Add it on the sheet by hand.' };
    SpreadsheetApp.flush();
    const bal = G.getValue();
    const pr = props_(), skip = JSON.parse(pr.getProperty('paid_skip') || '{}'); skip[holder] = (skip[holder] || 0) + amount; pr.setProperty('paid_skip', JSON.stringify(skip));
    const sh = sheet_('TopUps'); if (sh.getLastRow() === 0) sh.appendRow(['ts', 'name', 'amount', 'total']);
    sh.appendRow(["'" + new Date().toISOString(), holder, amount, before + amount]); bust_('TopUps');
    cache_().remove('players');
    var out = { ok: true, holder: holder, paid: before + amount, balance: bal, list: topupsAll_() };
  } finally { lock.releaseLock(); }
  pushTopup_(holder, amount, out.balance, ps);
  return out;
}
/** Admin: log a top-up by hand (one that was paid before tracking started, or missed). Doesn't touch the attendance sheet. */
function addTopup_(body) {
  const me = voterName_(body.t); if (!me || !isAdmin_(me)) return { ok: false, error: 'Admin only.' };
  const p = players_().filter(function (x) { return x.name.toLowerCase() === String(body.name || '').toLowerCase(); })[0];
  const amount = Math.round(Number(body.amount) || 0), date = /^\d{4}-\d{2}-\d{2}$/.test(String(body.date || '')) ? String(body.date) : '';
  if (!p) return { ok: false, error: 'Unknown player.' }; if (amount <= 0) return { ok: false, error: 'Amount?' }; if (!date) return { ok: false, error: 'Date?' };
  const sh = sheet_('TopUps'); if (sh.getLastRow() === 0) sh.appendRow(['ts', 'name', 'amount', 'total']);
  sh.appendRow(["'" + new Date(date + 'T12:00:00+03:00').toISOString(), p.name, amount, 'added by hand']); bust_('TopUps');
  return { ok: true, list: topupsAll_() };
}
function removeTopup_(body) {
  const me = voterName_(body.t); if (!me || !isAdmin_(me)) return { ok: false, error: 'Admin only.' };
  const sh = sheet_('TopUps'), data = sh.getDataRange().getValues();
  for (let i = data.length - 1; i >= 1; i--) if (String(data[i][0]) === String(body.ts) && String(data[i][1]) === String(body.name)) { sh.deleteRow(i + 1); break; }
  bust_('TopUps'); return { ok: true, list: topupsAll_() };
}
function topupsAll_() {
  if (!ss_().getSheetByName('TopUps')) return [];
  return tab_('TopUps').map(function (r) { return { ts: String(r[0]), name: String(r[1]), amount: Number(r[2]) || 0, hand: String(r[3]) === 'added by hand' }; })
    .sort(function (a, b) { return a.ts < b.ts ? 1 : a.ts > b.ts ? -1 : 0; }).slice(0, 60);
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
    if (status === 'pending') pushSafe_([adminName_()], 'claims', 'New claim: ' + p.name, 'Approve it in Admin → Players.', '/#admin');
    return { ok: true, token: token, name: p.name, pending: status === 'pending' };
  } finally { lock.releaseLock(); }
}

/** Aggregated ratings: calibrated, trimmed, familiarity-weighted means. Cached 2 min. */
/* Results are cached for 5 minutes, with a stale copy kept for 6 hours. When the fresh copy times out, one request rebuilds it
   (about 4 s) while everyone else gets the stale copy straight away, instead of all of them rebuilding at once and queueing.
   A write (bust_) clears both, so changes always show on the next read. */
function results_() {
  const c = cache_(), cached = c.get('results');
  if (cached) return JSON.parse(cached);
  const stale = c.get('results_stale');
  if (stale && c.get('results_building')) return JSON.parse(stale);
  if (stale) c.put('results_building', '1', 30);
  try { return buildResults_(); } finally { c.remove('results_building'); }
}
function buildResults_() {
  const win = props_().getProperty('window'), open = props_().getProperty('open') === '1';
  const windows = tab_('Windows').map(function (r) { return r[0]; });
  const prevWin = windows[windows.indexOf(win) - 1] || null;
  const all = tab_('Votes');
  const seeds = seeds_();
  const hide = hiddenNames_();
  const strip = function (agg) { Object.keys(agg).forEach(function (n) { delete agg[n].vote; delete agg[n].vw; if (hide.indexOf(n.toLowerCase()) >= 0) agg[n] = { n: agg[n].n, traits: agg[n].traits, hidden: true }; }); return agg; };
  const out = { ok: true, window: win, open: open, minVotes: MIN_VOTES, priorWeight: PRIOR_WEIGHT, requireCode: !!String(settings_().claim_code || '').trim(), photos: photos_(), claimed: claimedNames_(), players: strip(aggregate_(all.filter(function (r) { return r[0] === win; }), seeds)), prev: prevWin ? strip(aggregate_(all.filter(function (r) { return r[0] === prevWin; }), seeds)) : {}, prevWindow: prevWin };
  out.game = gamePublic_(); out.motm = motmTally_(); out.games = gamesPublic_();
  const lg = lastGame_(), ng = nextGame_(); out.last = lg ? gamePublic_(lg) : null; out.next = ng ? gamePublic_(ng) : null; out.fixtures = fixturesPublic_(); out.strikes = strikesPublic_(); out.footReviews = footReviews_(); out.profiles = profiles_();
  out.ballots = tab_('Voters').filter(function (r) { return r[3] && String(r[5] || '') && String(r[5]) !== '0'; }).length;   // people who have voted (a voter can hold an older anonymous ballot and a newer named one, so ballot ids would count them twice)
  const js = JSON.stringify(out); cache_().put('results', js, 300); cache_().put('results_stale', js, 21600);
  pushTickThrottled_();
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
  return { ok: true, window: win, selfRatings: selfRatings, seeds: seeds_(), traits: TRAITS, pending: pending, activity: activity_(), claims: claims, agg: agg, named: named, anon: anon, matchNamed: matchNamed, motmGames: motmGames, topups: topupsAll_(),
    pushPeople: pushRows_().reduce(function (o, r) { const n = String(r[1]); o[n] = (o[n] || 0) + 1; return o; }, {}) };
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
  const recent = rows.slice(-120).reverse().map(function (r) { return { ts: new Date(r[0]).toISOString(), name: String(r[2] || ''), dev: String(r[1] || '').slice(-4), m: !!Number(r[3]), page: String(r[4] || '') }; });
  return { today: sum(now - day), week: sum(now - 7 * day), total: sum(0), names: names, recent: recent, tabs: Object.keys(tabs).map(function (k) { return { tab: k, people: Object.keys(tabs[k]).length }; }).sort(function (a, b) { return b.people - a.people; }) };
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
    .map(function (g) { return { date: g.date, time: g.time, size: g.size || 6, slots: 2 * (g.size || 6), teams: !!g.teams, hasGate: !!gateFor_(g.date), signups: (su[g.date] || []).map(function (x) { return { name: x.name, kind: x.kind, by: x.by }; }) }; });
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
  if (body.guest) { name = String(body.guest).replace(/\s+/g, ' ').trim().slice(0, 30); kind = 'guest'; if (!name) return { ok: false, error: 'Type the guest’s name.' };
    if (ALIASES_[name.toLowerCase()]) { name = ALIASES_[name.toLowerCase()]; kind = 'player'; } }   // a nickname of a player (Veer is Vali) is that player, not a guest
  else if (body.name) { const p = players_().filter(function (x) { return x.name.toLowerCase() === String(body.name).toLowerCase(); })[0]; if (!p) return { ok: false, error: 'Unknown player.' }; name = p.name; }
  const lock = LockService.getScriptLock(); lock.waitLock(10000);
  try {
    const sh = signupSheet_(), data = sh.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) if (String(data[i][0]).slice(0, 10) === date && String(data[i][2]).toLowerCase() === name.toLowerCase()) return { ok: false, error: name + ' is already on the list.' };
    sh.appendRow(["'" + date, "'" + new Date().toISOString(), name, kind, name === me && kind === 'player' ? '' : me]);
  } finally { lock.releaseLock(); }
  bust_('Signups'); syncSheetSafe_(date);
  return { ok: true, fixtures: fixturesPublic_() };
}
function withdraw_(body) {
  const me = voterName_(body.t);
  if (!me) return { ok: false, error: 'Find your card first.' };
  const date = String(body.date || ''), name = String(body.name || me).toLowerCase(), admin = isAdmin_(me), before = signupsByDate_()[date] || [];
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
  bust_('Signups'); if (done && strike) bust_('Strikes'); if (done) syncSheetSafe_(date);
  if (done) { try { notifyPromoted_(date, before); } catch (e) { console.error(e); }
    if (strike) { pushStrikes_(); pushSafe_([adminName_()], 'latedrop', 'Late drop-out: ' + strike, 'Pulled out of ' + dayLabel_(date) + ' inside 36 hours · strike ' + Math.min(3, (strikesPublic_()[strike] || []).length) + ' of 3.', '/#admin'); } }
  return done ? { ok: true, fixtures: fixturesPublic_(), strikes: strikesPublic_(), strike: strike || '' } : { ok: false, error: 'Not on that list.' };
}
/* ───────── fill the attendance sheet from sign-ups ─────────
   Once a game's list is full, the people in the spots are charged in the attendance sheet the same way the sheet's own
   "Submit This Weeks List" does it (FootballGameUpdater): a row each in Game Log (date, name, fee, note) and a fee in a new
   dated column of Wednesday Group (header row 3, data from row 4, dates from column I; balance G = F − SUM(H:ZZ)).
   Until kick-off it keeps both in step: someone joining from standby is added, someone leaving is taken off, except that a
   player who dropped out of a spot within 36 hours (a strike) stays charged. After kick-off the game is left alone.
   What was written is remembered in the script property sheet_<date>, so a column made by hand for the same date is never
   touched (no double charge), and only rows this code wrote are ever changed. */
const SHEET_FEE = 350, SHEET_NOTE = 'from site sign-ups';
function sheetLabel_(date) { return Utilities.formatDate(new Date(date + 'T12:00:00+03:00'), 'Africa/Nairobi', 'dd MMM yyyy'); }
function sheetCharges_(g) {
  const slots = 2 * (g.size || 6), su = signupsByDate_()[g.date] || [];
  const T = g.teams && ((g.teams.black || []).length || (g.teams.white || []).length), names = T ? (g.teams.black || []).concat(g.teams.white || []) : (g.list || []);
  const out = [], seen = {}, add = function (n, note) { const k = String(n).toLowerCase(); if (!n || seen[k]) return; seen[k] = 1; out.push({ name: n, note: note }); };
  names.forEach(function (n) { const x = su.filter(function (y) { return y.name.toLowerCase() === n.toLowerCase(); })[0]; add(n, x && x.kind === 'guest' ? 'guest, ' + SHEET_NOTE + (x.by ? ', via ' + x.by : '') : SHEET_NOTE); });
  if (ss_().getSheetByName('Strikes')) tab_('Strikes').forEach(function (r) { if (String(r[0]).slice(0, 10) === g.date && r[2] && !r[4]) add(String(r[2]), SHEET_NOTE + ', dropped out late'); });
  return { full: !!T || names.length >= slots, list: out };
}
function syncSheet_(date, dry) {
  const g = allGames_().filter(function (x) { return x.date === date; })[0]; if (!g) return { ok: false, error: 'no game' };
  if (kickoff_(g).getTime() <= Date.now()) return { ok: true, skipped: 'played' };
  const pr = props_(), key = 'sheet_' + date;
  const lock = dry ? null : LockService.getScriptLock(); if (lock && !lock.tryLock(20000)) return { ok: false, error: 'busy' };
  try {
    const prev = JSON.parse(pr.getProperty(key) || 'null');   // read under the lock: another sync may have just made the column
    const c = sheetCharges_(g); if (!prev && !c.full) return { ok: true, skipped: 'not full' };
    const ss = SpreadsheetApp.openById(MAIN_SHEET_ID), wide = ss.getSheetByName('Wednesday Group'), log = ss.getSheetByName('Game Log');
    if (!wide || !log) return { ok: false, error: 'sheet tabs not found' };
    const label = sheetLabel_(date), HR = 3, FR = 4, DC = 9;
    // the date column: ours if we made it, otherwise make one; a column for this date made by hand means hands off
    const lastCol = wide.getLastColumn(), heads = wide.getRange(HR, 1, 1, lastCol).getDisplayValues()[0];
    let col = -1; for (let i = DC - 1; i < heads.length; i++) if (String(heads[i]).trim() === label) col = i + 1;
    if (dry) { const rn = wide.getRange(4, 2, wide.getLastRow() - 3, 1).getValues().map(function (r) { return String(r[0] || '').trim().toLowerCase(); }); return { dry: true, label: label, existingCol: col, prev: prev, charge: c.list.map(function (x) { return x.name + (rn.indexOf(x.name.toLowerCase()) >= 0 ? '' : ' (NEW ROW?)') + ' · ' + x.note; }) }; }
    if (col > 0 && !prev) { pr.setProperty(key, JSON.stringify({ hands: true })); return { ok: true, skipped: 'column already there' }; }
    if (prev && prev.hands) return { ok: true, skipped: 'column made by hand' };
    if (col < 0) {
      let last = DC - 1; for (let i = DC - 1; i < heads.length; i++) if (String(heads[i]).trim() !== '') last = i + 1;
      col = last + 1; if (col > wide.getMaxColumns()) wide.insertColumnsAfter(wide.getMaxColumns(), col - wide.getMaxColumns());
      wide.getRange(HR, col).setNumberFormat('@').setValue(label);
      if (col > DC) wide.getRange(HR, col - 1).copyTo(wide.getRange(HR, col), SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
    }
    // rows by name: exact, then the site's alias, then a unique prefix; anyone not there gets a new row like the sheet's script makes
    let lastRow = wide.getLastRow(); const rowNames = wide.getRange(FR, 2, Math.max(lastRow - FR + 1, 1), 1).getValues().map(function (r) { return String(r[0] || '').trim(); });
    const findRow = function (n) {
      const l = n.toLowerCase(), alias = Object.keys(ALIASES_).filter(function (a) { return ALIASES_[a].toLowerCase() === l; });
      let i = rowNames.findIndex(function (x) { return x.toLowerCase() === l; }); if (i >= 0) return FR + i;
      i = rowNames.findIndex(function (x) { return alias.indexOf(x.toLowerCase()) >= 0; }); if (i >= 0) return FR + i;
      const pre = rowNames.map(function (x, k) { return x.toLowerCase().indexOf(l) === 0 ? k : -1; }).filter(function (k) { return k >= 0; }); return pre.length === 1 ? FR + pre[0] : -1;
    };
    const want = {}; c.list.forEach(function (x) { want[x.name.toLowerCase()] = x; });
    const had = (prev && prev.rows) || {};
    // take off anyone no longer charged (only cells this code filled)
    Object.keys(had).forEach(function (k) { if (!want[k]) { wide.getRange(had[k], col).clearContent(); } });
    const rows = {};
    c.list.forEach(function (x) {
      let r = findRow(x.name);
      if (r < 0) { r = ++lastRow; wide.getRange(r, 2).setValue(x.name); wide.getRange(r, 1).setFormula('=COUNT(H' + r + ':ZZ' + r + ')'); wide.getRange(r, 6).setValue(0); wide.getRange(r, 7).setFormula('=F' + r + '-SUM(H' + r + ':ZZ' + r + ')'); rowNames.push(x.name); }
      wide.getRange(r, col).setValue(SHEET_FEE); rows[x.name.toLowerCase()] = r;
    });
    // Game Log: replace this code's rows for the date
    const data = log.getDataRange().getDisplayValues();
    for (let i = data.length - 1; i >= 1; i--) if (String(data[i][0]).trim() === label && String(data[i][3]).indexOf(SHEET_NOTE) >= 0) log.deleteRow(i + 1);
    if (c.list.length) log.getRange(log.getLastRow() + 1, 1, c.list.length, 4).setValues(c.list.map(function (x) { return [label, x.name, SHEET_FEE, x.note]; }));
    pr.setProperty(key, JSON.stringify({ col: label, rows: rows, at: new Date().toISOString(), n: c.list.length }));
    return { ok: true, charged: c.list.length, label: label };
  } finally { if (lock) lock.releaseLock(); }
}
function syncSheetSafe_(date) { try { return syncSheet_(date); } catch (e) { console.error('syncSheet_ ' + date + ': ' + e); return { ok: false, error: String(e) }; } }
/** Editor only: what syncing would write, without writing anything. */
function syncSheetsDryRun() { const out = {}; fixturesPublic_().forEach(function (f) { out[f.date] = syncSheet_(f.date, true); }); Logger.log(JSON.stringify(out, null, 1)); }
/** Editor or admin: sync every open game now (e.g. right after deploying). */
function syncAllSheets() { const out = {}; fixturesPublic_().forEach(function (f) { out[f.date] = syncSheetSafe_(f.date); }); Logger.log(JSON.stringify(out)); return out; }
/** A player sets their own position and preferred foot. Written to the attendance sheet (Wednesday Group C/D on every row
 *  that is this player, and Players B/C), which is where the site reads them from. Claiming to be two-footed ("Both") goes to
 *  review instead: teammates agree or disagree when they rate the player, and after FOOT_NEED answers the majority decides. */
const FOOT_NEED = 5;
function setProfile_(body) {
  const me = voterName_(body.t); if (!me) return { ok: false, error: 'Find your card first.' };
  const POS = ['Attack', 'Midfield', 'Defense', 'Goalkeeper'], FOOT = ['Right', 'Left', 'Both'];
  const pos = POS.indexOf(body.position) >= 0 ? body.position : null, foot = FOOT.indexOf(body.foot) >= 0 ? body.foot : null;
  if (!pos && !foot) return { ok: false, error: 'Nothing to change.' };
  if (foot === 'Both') { openFootClaim_(me); if (pos) writeProfile_(me, pos, null); return { ok: true, review: true }; }
  if (foot) closeFootClaim_(me, 'withdrawn');
  const n = writeProfile_(me, pos, foot);
  if (!n) return { ok: false, error: 'Your row was not found in the sheet.' };
  return { ok: true, position: pos, foot: foot };
}
function writeProfile_(name, pos, foot) {
  const ss = SpreadsheetApp.openById(MAIN_SHEET_ID), l = name.toLowerCase(), alias = Object.keys(ALIASES_).filter(function (a) { return ALIASES_[a].toLowerCase() === l; });
  const isMe = function (v) { const x = String(v || '').trim().toLowerCase(); return x === l || alias.indexOf(x) >= 0; };
  let n = 0;
  const wide = ss.getSheetByName('Wednesday Group');
  if (wide) wide.getRange(4, 2, wide.getLastRow() - 3, 1).getValues().forEach(function (r, i) { if (!isMe(r[0])) return; if (pos) wide.getRange(4 + i, 3).setValue(pos); if (foot) wide.getRange(4 + i, 4).setValue(foot); n++; });
  const pl = ss.getSheetByName('Players');
  if (pl) pl.getRange(2, 1, pl.getLastRow() - 1, 1).getValues().forEach(function (r, i) { if (!isMe(r[0])) return; if (pos) pl.getRange(2 + i, 2).setValue(pos); if (foot) pl.getRange(2 + i, 3).setValue(foot); });
  cache_().remove('players'); bust_('Voters');
  return n;
}
/* Profiles tab: name, JSON of what players say about themselves (weak foot, skill moves, work rates, shirt number,
   nickname, favourite club), updated. Public: shown on the back of their card. */
function profiles_() {
  const out = {}; if (!ss_().getSheetByName('Profiles')) return out;
  tab_('Profiles').forEach(function (r) { if (!r[0]) return; try { out[String(r[0])] = JSON.parse(String(r[1] || '{}')); } catch (_) {} });
  return out;
}
function setAttrs_(body) {
  const me = voterName_(body.t); if (!me) return { ok: false, error: 'Find your card first.' };
  const a = body.attrs || {}, cur = profiles_()[me] || {}, clean = function (v, n) { return String(v == null ? '' : v).replace(/[<>]/g, '').replace(/\s+/g, ' ').trim().slice(0, n); };
  const star = function (v) { const n = parseInt(v, 10); return n >= 1 && n <= 5 ? n : null; }, wr = function (v) { return ['Low', 'Med', 'High'].indexOf(v) >= 0 ? v : null; };
  if ('weakFoot' in a) cur.weakFoot = star(a.weakFoot);
  if ('skills' in a) cur.skills = star(a.skills);
  if ('wrAtt' in a) cur.wrAtt = wr(a.wrAtt);
  if ('wrDef' in a) cur.wrDef = wr(a.wrDef);
  if ('number' in a) { const n = parseInt(a.number, 10); cur.number = n >= 1 && n <= 99 ? n : null; }
  if ('nickname' in a) cur.nickname = clean(a.nickname, 20) || null;
  if ('club' in a) cur.club = clean(a.club, 24) || null;
  Object.keys(cur).forEach(function (k) { if (cur[k] === null || cur[k] === '') delete cur[k]; });
  const sh = sheet_('Profiles'); if (sh.getLastRow() === 0) sh.appendRow(['name', 'data', 'updated']);
  const d = sh.getDataRange().getValues(); let at = -1;
  for (let i = 1; i < d.length; i++) if (String(d[i][0]).toLowerCase() === me.toLowerCase()) at = i + 1;
  const row = [me, JSON.stringify(cur), new Date()];
  if (at > 0) sh.getRange(at, 1, 1, 3).setValues([row]); else sh.appendRow(row);
  bust_('Profiles');
  return { ok: true, attrs: cur };
}

/* FootClaims: ts, name, status (review | approved | rejected | withdrawn), decided. FootVotes: claim ts, name, voter, agree (1/0), ts. */
function footClaimSheet_() { const sh = sheet_('FootClaims'); if (sh.getLastRow() === 0) sh.appendRow(['ts', 'name', 'status', 'decided']); return sh; }
function openFootClaim_(name) {
  closeFootClaim_(name, 'withdrawn');
  footClaimSheet_().appendRow(["'" + new Date().toISOString(), name, 'review', '']); bust_('FootClaims');
}
function closeFootClaim_(name, status) {
  if (!ss_().getSheetByName('FootClaims')) return;
  const sh = footClaimSheet_(), d = sh.getDataRange().getValues();
  for (let i = 1; i < d.length; i++) if (String(d[i][1]).toLowerCase() === name.toLowerCase() && d[i][2] === 'review') sh.getRange(i + 1, 3, 1, 2).setValues([[status, "'" + new Date().toISOString()]]);
  bust_('FootClaims');
}
function footReviews_() {
  const out = {}; if (!ss_().getSheetByName('FootClaims')) return out;
  tab_('FootClaims').forEach(function (r) { if (r[2] === 'review') out[String(r[1])] = { ts: String(r[0]), agree: 0, disagree: 0, need: FOOT_NEED }; });
  if (ss_().getSheetByName('FootVotes')) tab_('FootVotes').forEach(function (r) { const c = out[String(r[1])]; if (c && String(r[0]) === c.ts) { if (Number(r[3])) c.agree++; else c.disagree++; } });
  return out;
}
function myFootVotes_(me) {
  const open = footReviews_(), out = {}; if (!ss_().getSheetByName('FootVotes')) return out;
  tab_('FootVotes').forEach(function (r) { const c = open[String(r[1])]; if (c && String(r[0]) === c.ts && String(r[2]).toLowerCase() === me.toLowerCase()) out[String(r[1])] = !!Number(r[3]); });
  return out;
}
function footVote_(body) {
  const me = voterName_(body.t); if (!me) return { ok: false, error: 'Find your card first.' };
  const name = String(body.name || ''), agree = body.agree ? 1 : 0;
  if (name.toLowerCase() === me.toLowerCase()) return { ok: false, error: "You can't vote on your own." };
  const lock = LockService.getScriptLock(); lock.waitLock(10000);
  try {
    const c = footReviews_()[name]; if (!c) return { ok: false, error: 'That review is closed.' };
    const sh = sheet_('FootVotes'); if (sh.getLastRow() === 0) sh.appendRow(['claim', 'name', 'voter', 'agree', 'ts']);
    const d = sh.getDataRange().getValues(); let at = -1;
    for (let i = 1; i < d.length; i++) if (String(d[i][0]) === c.ts && String(d[i][1]) === name && String(d[i][2]).toLowerCase() === me.toLowerCase()) at = i + 1;
    if (at > 0) sh.getRange(at, 4, 1, 2).setValues([[agree, "'" + new Date().toISOString()]]); else sh.appendRow([c.ts, name, me, agree, "'" + new Date().toISOString()]);
    bust_('FootVotes');
    const t = footReviews_()[name], total = t.agree + t.disagree;
    let decided = '';
    if (total >= FOOT_NEED && t.agree !== t.disagree) {
      decided = t.agree > t.disagree ? 'approved' : 'rejected';
      closeFootClaim_(name, decided);
      if (decided === 'approved') writeProfile_(name, null, 'Both');
    }
    return { ok: true, review: decided ? null : footReviews_()[name], decided: decided, footReviews: footReviews_() };
  } finally { lock.releaseLock(); }
}

/* ───────── push notifications ─────────
   Phones that turned notifications on are kept in the Push tab (endpoint, name, keys, prefs). This backend decides who gets
   what; https://wednesday910.vercel.app/api/push (Vercel) does the sending, because Apps Script can't sign web push.
   The shared secret lives in the script property push_secret (set once by the admin through the pushConfig action).
   Kinds, each switchable per player: promoted (moved up from standby), gameday (2 h before kick-off, with the gate code),
   rate (after the game), newgame (sign-ups open), balance (dropped below zero); for the admin: claims, latedrop.
   Timed ones are checked by pushTick (every 15 min trigger, and opportunistically when results are rebuilt); PushLog
   makes sure each is sent once. */
const PUSH_URL = 'https://wednesday910.vercel.app/api/push';
const PUSH_KINDS = ['promoted', 'gameday', 'rate', 'newgame', 'balance', 'topup', 'claims', 'latedrop', 'copy'];
function pushSheet_() { const sh = sheet_('Push'); if (sh.getLastRow() === 0) sh.appendRow(['endpoint', 'name', 'p256dh', 'auth', 'prefs', 'created', 'ua']); return sh; }
function pushRows_() { return ss_().getSheetByName('Push') ? tab_('Push').filter(function (r) { return r[0] && r[1]; }) : []; }
function parsePrefs_(v) { try { const o = JSON.parse(String(v || '{}')); return o && typeof o === 'object' ? o : {}; } catch (_) { return {}; } }
function myPush_(me) {
  const rows = pushRows_().filter(function (r) { return String(r[1]).toLowerCase() === me.toLowerCase(); });
  return { devices: rows.length, prefs: rows.length ? parsePrefs_(rows[rows.length - 1][4]) : {}, endpoints: rows.map(function (r) { return String(r[0]).slice(-24); }) };
}
function pushSub_(body) {
  const me = voterName_(body.t); if (!me) return { ok: false, error: 'Find your card first.' };
  const sub = body.sub || {}, ep = String(sub.endpoint || ''), k = sub.keys || {};
  if (!/^https:\/\//.test(ep) || !k.p256dh || !k.auth) return { ok: false, error: 'Bad subscription.' };
  const lock = LockService.getScriptLock(); lock.waitLock(10000);
  try {
    const sh = pushSheet_(), d = sh.getDataRange().getValues(), mine = myPush_(me);
    const prefs = JSON.stringify(body.prefs && typeof body.prefs === 'object' ? body.prefs : mine.prefs);
    for (let i = d.length - 1; i >= 1; i--) if (String(d[i][0]) === ep) sh.deleteRow(i + 1);
    sh.appendRow([ep, me, String(k.p256dh), String(k.auth), prefs, new Date(), String(body.ua || '').slice(0, 120)]);
  } finally { lock.releaseLock(); }
  bust_('Push');
  return { ok: true, push: myPush_(me) };
}
function pushUnsub_(body) {
  const me = voterName_(body.t); if (!me) return { ok: false };
  const sh = pushSheet_(), d = sh.getDataRange().getValues();
  for (let i = d.length - 1; i >= 1; i--) if (String(d[i][0]) === String(body.endpoint || '') && String(d[i][1]).toLowerCase() === me.toLowerCase()) sh.deleteRow(i + 1);
  bust_('Push'); return { ok: true, push: myPush_(me) };
}
function pushPrefs_(body) {
  const me = voterName_(body.t); if (!me) return { ok: false };
  const p = {}; PUSH_KINDS.forEach(function (k) { if (body.prefs && k in body.prefs) p[k] = !!body.prefs[k]; });
  const sh = pushSheet_(), d = sh.getDataRange().getValues(), js = JSON.stringify(p);
  for (let i = 1; i < d.length; i++) if (String(d[i][1]).toLowerCase() === me.toLowerCase()) sh.getRange(i + 1, 5).setValue(js);
  bust_('Push'); return { ok: true, push: myPush_(me) };
}
function pushConfig_(body) {
  const me = voterName_(body.t); if (!me || !isAdmin_(me)) return { ok: false, error: 'Admin only.' };
  if (body.secret) props_().setProperty('push_secret', String(body.secret));
  return { ok: true, configured: !!props_().getProperty('push_secret') };
}
/** Send one notification to these players' phones (those who haven't switched this kind off), and the admin a copy. */
function pushTo_(names, kind, title, text, url) {
  const res = pushSend_(names, kind, title, text, url), admin = adminName_().toLowerCase();
  const others = res.people.filter(function (n) { return n.toLowerCase() !== admin; });
  if (others.length && kind !== 'test' && kind !== 'copy') {
    try { pushSend_([adminName_()], 'copy', 'Sent: ' + title, 'To ' + others.join(', ') + ' · ' + text, '/#admin'); } catch (e) { console.error('copy ' + e); }
  }
  return res.sent;
}
function pushSend_(names, kind, title, text, url) {
  const none = { sent: 0, people: [] }, secret = props_().getProperty('push_secret'); if (!secret || !names || !names.length) return none;
  const want = {}; names.forEach(function (n) { want[String(n).toLowerCase()] = 1; });
  const rows = pushRows_().filter(function (r) { return want[String(r[1]).toLowerCase()] && parsePrefs_(r[4])[kind] !== false; });
  if (!rows.length) return none;
  const people = []; rows.forEach(function (r) { const n = String(r[1]); if (people.indexOf(n) < 0) people.push(n); });
  const messages = rows.map(function (r) { return { sub: { endpoint: String(r[0]), keys: { p256dh: String(r[2]), auth: String(r[3]) } }, title: title, body: text, url: url || '/', tag: kind }; });
  let res = null;
  try { res = JSON.parse(UrlFetchApp.fetch(PUSH_URL, { method: 'post', contentType: 'application/json', payload: JSON.stringify({ secret: secret, messages: messages }), muteHttpExceptions: true }).getContentText()); } catch (e) { console.error('push ' + e); return none; }
  if (res && res.gone && res.gone.length) {   // phones that turned notifications off or uninstalled
    const sh = pushSheet_(), d = sh.getDataRange().getValues();
    for (let i = d.length - 1; i >= 1; i--) if (res.gone.indexOf(String(d[i][0])) >= 0) sh.deleteRow(i + 1);
    bust_('Push');
  }
  return { sent: res && res.sent || 0, people: res && res.sent ? people : [] };
}
function pushSafe_(names, kind, title, text, url) { try { return pushTo_(names, kind, title, text, url); } catch (e) { console.error('pushTo_ ' + e); return 0; } }
/** Run fn once for this key (PushLog), so timed notifications go out once. */
function pushOnce_(key, fn) {
  const sh = sheet_('PushLog'); if (sh.getLastRow() === 0) sh.appendRow(['key', 'ts']);
  if (tab_('PushLog').some(function (r) { return String(r[0]) === key; })) return false;
  sh.appendRow([key, new Date()]); bust_('PushLog'); fn(); return true;
}
const adminName_ = function () { return String(settings_().admin_name || 'Raji'); };
const dayLabel_ = function (date) { return Utilities.formatDate(new Date(date + 'T12:00:00+03:00'), 'Africa/Nairobi', 'EEEE d MMM'); };
const koLabel_ = function (g) { const h = parseInt((g.time || '21:00').slice(0, 2), 10), m = (g.time || '21:00').slice(3, 5); return ((h % 12) || 12) + (m !== '00' ? ':' + m : '') + (h >= 12 ? 'pm' : 'am'); };
function gameNames_(g) { return g.teams && ((g.teams.black || []).length || (g.teams.white || []).length) ? (g.teams.black || []).concat(g.teams.white || []) : (g.list || []); }
/** People who moved from standby into a spot, comparing the list before a change with now. */
function notifyPromoted_(date, before) {
  const g = allGames_().filter(function (x) { return x.date === date; })[0]; if (!g || kickoff_(g).getTime() < Date.now()) return;
  const slots = 2 * (g.size || 6), after = (signupsByDate_()[date] || []).map(function (x) { return x.name.toLowerCase(); });
  const up = before.filter(function (x, i) { return i >= slots && x.kind === 'player' && after.indexOf(x.name.toLowerCase()) >= 0 && after.indexOf(x.name.toLowerCase()) < slots; }).map(function (x) { return x.name; });
  if (up.length) pushSafe_(up, 'promoted', "You're in for " + dayLabel_(date), 'Someone dropped out, so you moved up from standby. Kick-off ' + koLabel_(g) + '.', '/#home');
}
/** Timed notifications: game day (2 h before), rate the game (after), balance below zero (daytime only). */
function pushTick() {
  if (!props_().getProperty('push_secret') || !pushRows_().length) return;
  const now = Date.now(), today = Utilities.formatDate(new Date(), 'Africa/Nairobi', 'yyyy-MM-dd'), hour = Number(Utilities.formatDate(new Date(), 'Africa/Nairobi', 'H'));
  allGames_().forEach(function (g) {
    const left = kickoff_(g).getTime() - now;
    if (g.date === today && left > 0 && left <= 2.25 * 3600000) pushOnce_('gameday|' + g.date, function () {
      const gate = gateFor_(g.date), tail = gate ? 'Gate code ' + gate + '. See you on the pitch.' : 'See you on the pitch.', T = g.teams || {};
      if ((T.black || []).length || (T.white || []).length) {   // teams are in: tell each side their colour
        if ((T.black || []).length) pushSafe_(T.black, 'gameday', "You're on Black tonight · " + koLabel_(g), 'Bring a black top. ' + tail, '/#home');
        if ((T.white || []).length) pushSafe_(T.white, 'gameday', "You're on White tonight · " + koLabel_(g), 'Bring a white top. ' + tail, '/#home');
      } else pushSafe_(gameNames_(g), 'gameday', 'Game tonight · ' + koLabel_(g), tail, '/#home');
    });
  });
  try { pushStrikes_(); } catch (e) { console.error('pushStrikes_ ' + e); }
  const L = lastGame_();
  if (L) { const ko = kickoff_(L).getTime(), closes = closes_(L.date); if (now >= ko + 3600000 && now < closes) pushOnce_('rate|' + L.date, function () {
    pushSafe_(gameNames_(L), 'rate', 'Rate the game', 'How did everyone play? Pick your man of the match. Voting closes at 11am.', '/#home'); }); }
  if (hour >= 9 && hour < 21) {
    const subs = {}; pushRows_().forEach(function (r) { subs[String(r[1]).toLowerCase()] = 1; });
    players_().forEach(function (p) { if (!subs[p.name.toLowerCase()] || !(p.bal < 0)) return;
      pushOnce_('bal|' + p.name + '|' + p.bal, function () { pushSafe_([p.name], 'balance', 'Your balance is −KES ' + Math.abs(p.bal).toLocaleString('en'), 'Time to top up the pot.', '/#home'); }); });
  }
}
function pushTickThrottled_() { const c = cache_(); if (c.get('pushtick')) return; c.put('pushtick', '1', 600); try { pushTick(); } catch (e) { console.error('pushTick ' + e); } }
/** Editor, once: run pushTick every 15 minutes. */
function setupPushTrigger() {
  ScriptApp.getProjectTriggers().forEach(function (t) { if (t.getHandlerFunction() === 'pushTick') ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('pushTick').timeBased().everyMinutes(15).create();
}
/** Admin: send a notification to everyone with notifications on (or just count them first with dry). */
function pushBroadcast_(body) {
  const me = voterName_(body.t); if (!me || !isAdmin_(me)) return { ok: false, error: 'Admin only.' };
  const rows = pushRows_(), people = []; rows.forEach(function (r) { const n = String(r[1]); if (people.indexOf(n) < 0) people.push(n); });
  if (body.dry) return { ok: true, phones: rows.length, people: people };
  const title = String(body.title || '').slice(0, 60).trim(), text = String(body.body || '').slice(0, 200).trim();
  if (!title) return { ok: false, error: 'Add a title.' };
  return { ok: true, sent: pushTo_(people, 'news', title, text, String(body.url || '/#home')), phones: rows.length, people: people };
}
/** Admin: send any strike notifications that haven't gone out yet (they also go out on the 15-minute tick). */
function pushStrikesNow_(body) {
  const me = voterName_(body.t); if (!me || !isAdmin_(me)) return { ok: false, error: 'Admin only.' };
  pushStrikes_(); return { ok: true };
}
function pushTest_(body) {
  const me = voterName_(body.t); if (!me) return { ok: false };
  return { ok: true, sent: pushTo_([me], 'test', 'Notifications are on', "You'll get a ping when it matters. You can change which ones on the Me tab.", '/#me') };
}

/* ───────── sign in on another device ─────────
   An installed iPhone app has its own storage, so it doesn't know the card picked in Safari. A signed-in device asks for a
   short code (valid 15 minutes); typing it on the new device hands over the same link, so it's the same card and ballot. */
const LINK_ABC = 'ABCDEFGHJKLMNPQRSTUVWXYZ23456789';
function linkCode_(body) {
  const me = voterName_(body.t); if (!me) return { ok: false, error: 'Find your card first.' };
  let code = ''; for (let i = 0; i < 6; i++) code += LINK_ABC.charAt(Math.floor(Math.random() * LINK_ABC.length));
  cache_().put('lc|' + code, String(body.t), 900);
  return { ok: true, code: code, minutes: 15 };
}
function linkRedeem_(body) {
  const c = cache_(), code = String(body.code || '').toUpperCase().replace(/[^A-Z0-9]/g, ''), fails = Number(c.get('lcfail') || 0);
  if (fails > 40) return { ok: false, error: 'Too many tries. Wait a few minutes and get a new code.' };
  const t = code.length === 6 ? c.get('lc|' + code) : null;
  if (!t) { c.put('lcfail', String(fails + 1), 900); return { ok: false, error: 'That code is wrong or has expired. Get a new one.' }; }
  c.remove('lc|' + code);
  const name = voterName_(t); if (!name) return { ok: false, error: 'That card is no longer linked. Pick it again.' };
  return { ok: true, token: t, name: name };
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
/** A player's own strikes (not cleared), with whether they've seen the pop-up about each (column F). */
function myStrikes_(name) {
  if (!ss_().getSheetByName('Strikes')) return [];
  return tab_('Strikes').filter(function (r) { return String(r[2]).toLowerCase() === name.toLowerCase() && !r[4]; })
    .map(function (r) { return { date: String(r[0]).slice(0, 10), ts: String(r[1]).replace(/^'/, ''), seen: !!r[5], fee: SHEET_FEE }; });
}
function ackStrikes_(body) {
  const me = voterName_(body.t); if (!me) return { ok: false };
  const sh = strikeSheet_(), d = sh.getDataRange().getValues(), stamp = "'" + new Date().toISOString();
  if (sh.getRange(1, 6).getValue() !== 'seen') sh.getRange(1, 6).setValue('seen');
  for (let i = 1; i < d.length; i++) if (String(d[i][2]).toLowerCase() === me.toLowerCase() && !d[i][4] && !d[i][5]) sh.getRange(i + 1, 6).setValue(stamp);
  bust_('Strikes'); return { ok: true };
}
/** Admin: give a strike by hand (e.g. the admin took someone off late). Counts like any other: shown, and charged for that game. */
function addStrike_(body) {
  const me = voterName_(body.t); if (!me || !isAdmin_(me)) return { ok: false, error: 'Admin only.' };
  const date = /^\d{4}-\d{2}-\d{2}$/.test(String(body.date || '')) ? String(body.date) : ''; if (!date) return { ok: false, error: 'Which game?' };
  const p = players_().filter(function (x) { return x.name.toLowerCase() === String(body.name || '').toLowerCase(); })[0]; if (!p) return { ok: false, error: 'Unknown player.' };
  strikeSheet_().appendRow(["'" + date, "'" + new Date().toISOString(), p.name, me + ' (by hand)', '']); bust_('Strikes');
  const sync = syncSheetSafe_(date); try { pushStrikes_(); } catch (e) { console.error(e); }
  return { ok: true, strikes: strikesPublic_(), sync: sync };
}
/** Tell each player about a strike they haven't seen yet, once per strike (PushLog), also from pushTick as a catch-up. */
function pushStrikes_() {
  if (!props_().getProperty('push_secret') || !ss_().getSheetByName('Strikes')) return;
  const all = strikesPublic_(), subs = {}; pushRows_().forEach(function (r) { subs[String(r[1]).toLowerCase()] = 1; });
  tab_('Strikes').forEach(function (r) {
    if (!r[2] || r[4] || r[5] || !subs[String(r[2]).toLowerCase()]) return;   // wait until they have notifications on
    const name = String(r[2]), date = String(r[0]).replace(/^'/, '').slice(0, 10), n = Math.min(3, (all[name] || []).length);
    pushOnce_('strike|' + name + '|' + date, function () {
      pushSafe_([name], 'strike', 'You got a strike', 'You came off ' + dayLabel_(date) + ' inside 36 hours of kick-off, so you\'ve been charged KES ' + SHEET_FEE + '. Strike ' + n + ' of 3.', '/#home');
    });
  });
}
/** Admin: clear someone's strikes (all of them) */
function clearStrikes_(body) {
  const me = voterName_(body.t); if (!me || !isAdmin_(me)) return { ok: false, error: 'Admin only.' };
  const name = String(body.name || '').toLowerCase(); if (!name || !ss_().getSheetByName('Strikes')) return { ok: false, error: 'No strikes.' };
  const sh = strikeSheet_(), data = sh.getDataRange().getValues(), stamp = "'" + new Date().toISOString();
  const dates = {};
  for (let i = 1; i < data.length; i++) if (String(data[i][2]).toLowerCase() === name && !data[i][4]) { sh.getRange(i + 1, 5).setValue(stamp); dates[String(data[i][0]).replace(/^'/, '').slice(0, 10)] = 1; }
  bust_('Strikes');
  Object.keys(dates).forEach(function (d) { syncSheetSafe_(d); });   // a cleared late drop is no longer charged for a game still to come
  return { ok: true, strikes: strikesPublic_() };
}
/** Admin: everything about one player, for the player inspector: claim, visits, profile, ratings given, match votes,
 *  sign-ups, strikes, notifications, top-ups. Votes are only those saved with the voter's name (14 Oct on). */
function adminPlayer_(body) {
  const me = voterName_(body.t); if (!me || !isAdmin_(me)) return { ok: false, error: 'Admin only.' };
  const name = String(body.name || ''), l = name.toLowerCase(), has = function (t) { return !!ss_().getSheetByName(t); };
  const iso = function (v) { const d = v instanceof Date ? v : new Date(String(v).replace(/^'/, '')); return isNaN(d) ? String(v || '') : d.toISOString(); };
  const vr = tab_('Voters').filter(function (r) { return String(r[1]).toLowerCase() === l; })[0];
  const visits = has('Visits') ? tab_('Visits').filter(function (r) { return String(r[2]).toLowerCase() === l; }).map(function (r) { return { ts: iso(r[0]), dev: String(r[1] || '').slice(-4), m: !!Number(r[3]), page: String(r[4] || '') }; }) : [];
  const votesAll = tab_('Votes');
  const given = votesAll.filter(function (r) { return String(r[13] || '').toLowerCase() === l; }).map(function (r) { return { win: String(r[0]), target: String(r[2]), stats: r.slice(4, 10).map(Number), day: r[11] instanceof Date ? Utilities.formatDate(r[11], 'Africa/Nairobi', 'yyyy-MM-dd') : String(r[11] || '').replace(/^'/, ''), traits: String(r[12] || ''), comment: String(r[10] || '') }; });
  const received = votesAll.filter(function (r) { return String(r[2]).toLowerCase() === l; }).length;
  const mv = {}; if (has('MatchVotes')) tab_('MatchVotes').forEach(function (r) { if (String(r[6] || '').toLowerCase() !== l) return; const d = String(r[0]).slice(0, 10); mv[d] = mv[d] || { date: d, rated: 0, motm: '', ts: iso(r[5]) }; mv[d].rated++; if (r[4]) mv[d].motm = String(r[2]); });
  const su = []; const all = signupsByDate_(); Object.keys(all).forEach(function (d) { const g = allGames_().filter(function (x) { return x.date === d; })[0], slots = 2 * ((g && g.size) || 6);
    all[d].forEach(function (x, i) { if (x.name.toLowerCase() === l || String(x.by).toLowerCase() === l) su.push({ date: d, ts: iso(x.ts), name: x.name, kind: x.kind, by: x.by, pos: i + 1, inn: i < slots }); }); });
  const strikes = has('Strikes') ? tab_('Strikes').filter(function (r) { return String(r[2]).toLowerCase() === l; }).map(function (r) { return { date: String(r[0]).slice(0, 10), ts: iso(r[1]), by: String(r[3] || ''), cleared: !!r[4] }; }) : [];
  const prof = has('Profiles') ? tab_('Profiles').filter(function (r) { return String(r[0]).toLowerCase() === l; })[0] : null;
  const foot = has('FootClaims') ? tab_('FootClaims').filter(function (r) { return String(r[1]).toLowerCase() === l; }).map(function (r) { return { ts: iso(r[0]), status: String(r[2]), decided: iso(r[3]) }; }) : [];
  const self = tab_('SelfRatings').filter(function (r) { return String(r[1]).toLowerCase() === l; }).map(function (r) { return { win: String(r[0]), updated: iso(r[r.length - 1]) }; });
  return { ok: true, name: name,
    claim: vr ? { created: iso(vr[2]), claimed: vr[3] ? iso(vr[3]) : '', status: String(vr[4] || 'approved'), rated: String(vr[5] || '') } : null,
    visits: visits.slice(-200).reverse(), visitCount: visits.length, devices: Object.keys(visits.reduce(function (o, v) { o[v.dev] = 1; return o; }, {})).length,
    given: given, received: received, matchVotes: Object.keys(mv).map(function (k) { return mv[k]; }).sort(function (a, b) { return a.date < b.date ? 1 : -1; }),
    signups: su.sort(function (a, b) { return a.ts < b.ts ? 1 : -1; }), strikes: strikes, push: myPush_(name),
    profile: prof ? { attrs: parsePrefs_(prof[1]), updated: iso(prof[2]) } : null, foot: foot, self: self,
    photo: !!photos_()[name], topups: topupsFor_(name).list };
}
/** Admin: take charge of a game's dated column again (e.g. after it was wrongly marked as made by hand): rebuild what was
 *  written from the column itself (every row holding the fee), then sync. */
function sheetAdopt_(body) {
  const me = voterName_(body.t); if (!me || !isAdmin_(me)) return { ok: false, error: 'Admin only.' };
  const date = String(body.date || ''), label = sheetLabel_(date);
  const wide = SpreadsheetApp.openById(MAIN_SHEET_ID).getSheetByName('Wednesday Group'), heads = wide.getRange(3, 1, 1, wide.getLastColumn()).getDisplayValues()[0];
  let col = -1; for (let i = 8; i < heads.length; i++) if (String(heads[i]).trim() === label) col = i + 1;
  if (col < 0) return { ok: false, error: 'No column for ' + label };
  const vals = wide.getRange(4, 1, wide.getLastRow() - 3, col).getValues(), rows = {};
  vals.forEach(function (r, i) { const v = Number(r[col - 1]); if (r[1] && v) rows[String(r[1]).trim().toLowerCase()] = 4 + i; });
  props_().setProperty('sheet_' + date, JSON.stringify({ col: label, rows: rows, at: new Date().toISOString(), adopted: true }));
  return { ok: true, adopted: Object.keys(rows).length, sync: syncSheetSafe_(date) };
}
/** Admin: a name on a game's list was really someone else (Veer is Vali): rename it in place (keeps the spot), drop that
 *  person's other entry for the same game, re-sync the attendance sheet, and delete the sheet row the old name got if empty. */
function renameSignup_(body) {
  const me = voterName_(body.t); if (!me || !isAdmin_(me)) return { ok: false, error: 'Admin only.' };
  const date = String(body.date || ''), from = String(body.from || '').trim(), to = String(body.to || '').trim();
  const p = players_().filter(function (x) { return x.name.toLowerCase() === to.toLowerCase(); })[0]; if (!p) return { ok: false, error: 'Unknown player ' + to };
  const lock = LockService.getScriptLock(); lock.waitLock(10000); let renamed = 0, dropped = 0;
  try {
    const sh = signupSheet_(), d = sh.getDataRange().getValues();
    let at = -1; for (let i = 1; i < d.length; i++) if (String(d[i][0]).slice(0, 10) === date && String(d[i][2]).toLowerCase() === from.toLowerCase()) at = i;
    if (at < 0) return { ok: false, error: from + ' is not on ' + date };
    sh.getRange(at + 1, 3, 1, 3).setValues([[p.name, 'player', '']]); renamed = 1;
    for (let i = d.length - 1; i >= 1; i--) if (i !== at && String(d[i][0]).slice(0, 10) === date && String(d[i][2]).toLowerCase() === p.name.toLowerCase()) { sh.deleteRow(i + 1); dropped++; }
  } finally { lock.releaseLock(); }
  bust_('Signups');
  const sync = syncSheetSafe_(date);
  let removedRow = false;
  try {   // the old name's own row in Wednesday Group, if it now holds nothing
    const wide = SpreadsheetApp.openById(MAIN_SHEET_ID).getSheetByName('Wednesday Group'), last = wide.getLastRow(), w = wide.getLastColumn();
    const vals = wide.getRange(4, 1, last - 3, w).getValues();
    for (let i = vals.length - 1; i >= 0; i--) if (String(vals[i][1]).trim().toLowerCase() === from.toLowerCase()) {
      const fees = vals[i].slice(7).some(function (v) { return v !== '' && v !== null && Number(v) !== 0; }), paid = Number(vals[i][5]) || 0;
      if (!fees && !paid) { const r = 4 + i; wide.deleteRow(r); removedRow = true;
        const pr = props_(), all = pr.getProperties();   // rows remembered by the sheet sync shift up by one below the deleted row
        Object.keys(all).forEach(function (k) { if (k.indexOf('sheet_') !== 0) return; const st = JSON.parse(all[k] || 'null'); if (!st || !st.rows) return;
          Object.keys(st.rows).forEach(function (n) { if (st.rows[n] > r) st.rows[n]--; else if (st.rows[n] === r) delete st.rows[n]; }); pr.setProperty(k, JSON.stringify(st)); }); }
    }
  } catch (e) { console.error(e); }
  return { ok: true, renamed: renamed, dropped: dropped, sync: sync, removedRow: removedRow, fixtures: fixturesPublic_() };
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
  pushSafe_(pushRows_().map(function (r) { return String(r[1]); }), 'newgame', 'New game: ' + dayLabel_(date), 'Sign-ups are open · ' + (2 * size) + ' spots, ' + koLabel_({ time: time }) + ' kick-off.', '/#in');
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
  syncSheetSafe_(date);
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
  bust_('Signups'); syncSheetSafe_(date);
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
  bust_('Signups'); syncSheetSafe_(from); syncSheetSafe_(to);
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
/** Gate code for one game: per-date setting gate_<date>, else the game-day settings' code if it's that game. */
function gateFor_(date) { const s = settings_(), g = String(s['gate_' + date] || '').trim(); if (g) return g; const cur = gameInfo_(); return cur.date === date ? cur.gate : ''; }
function setGate_(body) {
  const me = voterName_(body.t); if (!me || !isAdmin_(me)) return { ok: false, error: 'Admin only.' };
  const date = /^\d{4}-\d{2}-\d{2}$/.test(String(body.date || '')) ? String(body.date) : ''; if (!date) return { ok: false, error: 'Which game?' };
  const gate = String(body.gate || '').replace(/\s+/g, '').toUpperCase().slice(0, 20);
  setSetting_('gate_' + date, gate);
  if (gameInfo_().date === date) setSetting_('gate_code', gate);
  return { ok: true, gate: gate, fixtures: fixturesPublic_() };
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
  return { date: g.date, time: g.time, squad: Object.keys(sq).map(function (k) { return sq[k]; }), teams: g.teams, score: g.score, ratings: ratings, raters: Object.keys(raters).length, hasGate: !!gateFor_(g.date),
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
  c.removeAll(['results', 'results_stale']);
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
    const paid = parseFloat((r[5] || '').replace(/,/g, '')) || 0, balRaw = parseFloat((r[6] || '').replace(/,/g, '')), bal = isNaN(balRaw) ? null : balRaw;
    if (seen[n]) { const m = seen[n]; m.att = m.att.map(function (v, i) { return v | att[i]; }); m.count += count; m.paid += paid; if (bal !== null) m.bal = (m.bal || 0) + bal; return; }   // same name twice = one person
    const p = { name: name, count: count, position: (r[2] || '').trim(), att: att, paid: paid, bal: bal }; seen[n] = p; out.push(p);
  });
  try { topupWatch_(out); } catch (e) { console.error(e); }
  cache_().put('players', JSON.stringify(out), 600);
  return out;
}

/** Editor-only: time each part of results_() to find what's slow. */
function debugTiming() {
  const t = {}, T = function (k, f) { const s = Date.now(); try { f(); } catch (e) { t[k + '_err'] = String(e); } t[k] = Date.now() - s; };
  T('cacheGetResults', function () { cache_().get('results'); });
  T('cacheGetPlayers', function () { t.playersCached = !!cache_().get('players'); });
  T('players', function () { players_(); });
  T('votes', function () { tab_('Votes'); });
  T('photos', function () { photos_(); });
  T('claimed', function () { claimedNames_(); });
  T('seeds', function () { seeds_(); });
  T('gamePublic', function () { gamePublic_(); });
  T('motmTally', function () { motmTally_(); });
  T('gamesPublic', function () { gamesPublic_(); });
  T('lastNext', function () { lastGame_(); nextGame_(); });
  T('fixtures', function () { fixturesPublic_(); });
  T('strikes', function () { strikesPublic_(); });
  cache_().remove('results');
  T('resultsFull', function () { results_(); });
  T('resultsCached', function () { results_(); });
  T('playersAfter', function () { t.playersCachedAfter = !!cache_().get('players'); t.playersLen = (cache_().get('players') || '').length; });
  Logger.log(JSON.stringify(t));
}
