/**
 * Squats Challenge — Google Apps Script web app.
 * Data lives in this spreadsheet: "Players" (who joined, when) and "Log" (every entry).
 * A player's total is the sum of their Log rows, so you can fix mistakes by editing or deleting Log rows.
 */

const PLAYERS = 'Players';
const LOG = 'Log';

// Players already in the challenge on 3 Nov 2026 (name, start date, squats so far).
const SEED = [
  ['Manoj Gupta', '2026-10-21', 770],
  ['Manoj Kumar', '2026-10-21', 400],
  ['Akhil', '2026-10-21', 485],
  ['Nikhil', '2026-10-23', 385],
  ['Geeta', '2026-10-27', 300],
  ['Java', '2026-10-29', 460],
  ['Aman', '2026-11-02', 90],
  ['Bhavya', '2026-11-03', 0],
  ['Sugandha', '2026-11-03', 0],
];
const SEED_DATE = '2026-11-03';

function doGet(e) {
  // The hosted page asks for the leaderboard with ?api=board
  if (e && e.parameter && e.parameter.api === 'board') return json_({ data: getBoard() });
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Squats Challenge')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/** The hosted page sends changes here as {fn, args}. */
function doPost(e) {
  try {
    const body = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    const actions = { getBoard: getBoard, logSquats: logSquats, joinChallenge: joinChallenge, removeLast: removeLast };
    const fn = actions[body.fn];
    if (!fn) throw new Error('Unknown action.');
    return json_({ data: fn.apply(null, Array.isArray(body.args) ? body.args : []) });
  } catch (err) {
    return json_({ error: String((err && err.message) || err) });
  }
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

/** Returns the two sheets, creating and seeding them on first use. */
function sheets_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let p = ss.getSheetByName(PLAYERS);
  let l = ss.getSheetByName(LOG);
  if (p && l) return { ss, p, l };

  const lock = LockService.getScriptLock();
  const mine = !lock.hasLock();
  if (mine) lock.waitLock(20000);
  try {
    p = ss.getSheetByName(PLAYERS);
    l = ss.getSheetByName(LOG);
    if (!p) {
      p = ss.insertSheet(PLAYERS);
      p.getRange('A:B').setNumberFormat('@');
      p.appendRow(['Name', 'Start date (YYYY-MM-DD)']);
      p.setFrozenRows(1);
    }
    if (!l) {
      l = ss.insertSheet(LOG);
      l.getRange('B:C').setNumberFormat('@');
      l.getRange('E:E').setNumberFormat('@');
      l.appendRow(['Time', 'Name', 'For date', 'Squats', 'Type']);
      l.setFrozenRows(1);
      if (p.getLastRow() <= 1) {
        SEED.forEach(function (s) {
          p.appendRow([s[0], s[1]]);
          l.appendRow([new Date(), s[0], SEED_DATE, s[2], 'join']);
        });
      }
    }
    const blank = ss.getSheetByName('Sheet1');
    if (blank && blank.getLastRow() === 0 && ss.getSheets().length > 1) ss.deleteSheet(blank);
    SpreadsheetApp.flush();
  } finally {
    if (mine) lock.releaseLock();
  }
  return { ss, p, l };
}

function dateStr_(v, tz) {
  if (v instanceof Date) return Utilities.formatDate(v, tz, 'yyyy-MM-dd');
  return String(v == null ? '' : v).trim();
}

function validDate_(s) {
  return typeof s === 'string' && /^\d{4}-\d{2}-\d{2}$/.test(s);
}

function cleanName_(name) {
  return String(name == null ? '' : name)
    .replace(/\s+/g, ' ')
    .trim()
    .replace(/^[=+\-@]+/, '') // keep names from being read as spreadsheet formulas
    .trim()
    .slice(0, 40);
}

/** Everything the page needs: players with totals and their full history, plus a recent-activity feed. */
function getBoard() {
  const s = sheets_();
  const tz = s.ss.getSpreadsheetTimeZone();
  const pv = s.p.getLastRow() > 1 ? s.p.getRange(2, 1, s.p.getLastRow() - 1, 2).getValues() : [];
  const lv = s.l.getLastRow() > 1 ? s.l.getRange(2, 1, s.l.getLastRow() - 1, 5).getValues() : [];
  const byName = {};
  const players = [];
  pv.forEach(function (r) {
    const name = String(r[0]).trim();
    if (!name || byName[name.toLowerCase()]) return;
    const o = { name: name, startDate: dateStr_(r[1], tz), total: 0, history: [] };
    byName[name.toLowerCase()] = o;
    players.push(o);
  });
  const feed = [];
  lv.forEach(function (r) {
    const o = byName[String(r[1]).trim().toLowerCase()];
    if (!o) return;
    const d = Number(r[3]) || 0;
    o.total = Math.round((o.total + d) * 10) / 10;
    const e = {
      at: r[0] instanceof Date ? r[0].getTime() : 0,
      date: dateStr_(r[2], tz),
      delta: d,
      total: o.total,
      kind: String(r[4] || 'add'),
    };
    o.history.push(e);
    feed.push({ name: o.name, at: e.at, date: e.date, delta: e.delta, total: e.total, kind: e.kind });
  });
  feed.sort(function (a, b) { return b.at - a.at; });
  return { players: players, feed: feed.slice(0, 15) };
}

function findPlayer_(board, name) {
  const key = cleanName_(name).toLowerCase();
  return board.players.find(function (x) { return x.name.toLowerCase() === key; });
}

/** mode "add": add `amount` squats. mode "set": `amount` is the new running total. */
function logSquats(name, amount, date, mode) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const s = sheets_();
    const o = findPlayer_(getBoard(), name);
    if (!o) throw new Error('We could not find that player. Refresh the page and try again.');
    let n = Number(amount);
    if (!isFinite(n)) throw new Error('Type a number of squats.');
    n = Math.round(n * 10) / 10;
    if (mode !== 'set' && n <= 0) throw new Error('Type how many squats you did.');
    if (mode === 'set' && n < 0) throw new Error('Your total can not be below zero.');
    const delta = mode === 'set' ? Math.round((n - o.total) * 10) / 10 : n;
    if (Math.abs(delta) > 5000) throw new Error('That number looks too big. Check it and try again.');
    const d = validDate_(date) ? date : Utilities.formatDate(new Date(), s.ss.getSpreadsheetTimeZone(), 'yyyy-MM-dd');
    s.l.appendRow([new Date(), o.name, d, delta, mode === 'set' ? 'set' : 'add']);
    SpreadsheetApp.flush();
  } finally {
    lock.releaseLock();
  }
  return getBoard();
}

function joinChallenge(name, startDate, initial, today) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const s = sheets_();
    const clean = cleanName_(name);
    if (!clean) throw new Error('Type your name.');
    const board = getBoard();
    if (findPlayer_(board, clean)) throw new Error(clean + ' is already in the challenge. Pick your name from the list.');
    if (board.players.length >= 200) throw new Error('The challenge is full.');
    if (!validDate_(startDate)) throw new Error('Pick the day you started.');
    let n = Number(initial) || 0;
    n = Math.round(n * 10) / 10;
    if (n < 0 || n > 100000) throw new Error('Check your squats so far.');
    const tz = s.ss.getSpreadsheetTimeZone();
    s.p.appendRow([clean, startDate]);
    s.l.appendRow([new Date(), clean, validDate_(today) ? today : Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd'), n, 'join']);
    SpreadsheetApp.flush();
  } finally {
    lock.releaseLock();
  }
  return getBoard();
}

/** Removes the player's most recent entry (never their joining row). */
function removeLast(name) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const s = sheets_();
    const key = cleanName_(name).toLowerCase();
    const last = s.l.getLastRow();
    if (last > 1) {
      const rows = s.l.getRange(2, 1, last - 1, 5).getValues();
      for (let i = rows.length - 1; i >= 0; i--) {
        if (String(rows[i][1]).trim().toLowerCase() === key && String(rows[i][4]) !== 'join') {
          s.l.deleteRow(i + 2);
          break;
        }
      }
      SpreadsheetApp.flush();
    }
  } finally {
    lock.releaseLock();
  }
  return getBoard();
}
