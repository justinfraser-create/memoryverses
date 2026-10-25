/**
 * TVAC Values Verse Tracker — Google Sheet sync
 * ---------------------------------------------
 * Paste this whole file into Extensions → Apps Script in your tracker Google Sheet,
 * set SYNC_KEY below, then Deploy → New deployment → Web app
 * (Execute as: Me · Who has access: Anyone).
 *
 * The teacher page sends each tally here and reads the whole class back.
 * The Sheet keeps four tabs, all created automatically:
 *   Overview  – one row per student, one column per value (successes / attempts). Rebuilt on every change.
 *   Tallies   – one row per student per value (the record the page reads).
 *   Students  – the class list.
 *   Settings  – the term start date.
 * Edit names on the page, not in the Sheet: the page rewrites the Students tab.
 */

const SYNC_KEY = 'PASTE_YOUR_SYNC_KEY_HERE';

const TALLY_HEAD = ['Week', 'Value', 'Verse', 'Student ID', 'Student', 'Attempts', 'Successes', 'Value ID', 'Last updated'];
const STUDENT_HEAD = ['Student ID', 'Name'];
const SETTING_HEAD = ['Setting', 'Value'];

function doGet(e) {
  const p = (e && e.parameter) || {};
  if (!keyOk(p.key)) return out({ error: 'bad_key' });
  return out(readAll());
}

function doPost(e) {
  let body;
  try { body = JSON.parse(e.postData.contents); } catch (err) { return out({ error: 'bad_request' }); }
  if (!keyOk(body.key)) return out({ error: 'bad_key' });
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);                       // one change at a time, so two devices never overwrite each other
  try {
    apply(body);
    SpreadsheetApp.flush();
    return out(readAll());
  } finally {
    lock.releaseLock();
  }
}

function keyOk(k) {
  return SYNC_KEY && SYNC_KEY !== 'PASTE_YOUR_SYNC_KEY_HERE' && k === SYNC_KEY;
}

function out(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

/* ---------- sheets ---------- */
function tab(name, head) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let s = ss.getSheetByName(name);
  if (!s) {
    s = ss.insertSheet(name);
    s.getRange(1, 1, 1, head.length).setValues([head]).setFontWeight('bold');
    s.setFrozenRows(1);
  }
  return s;
}
function rows(sheet) {                         // all rows below the header, as displayed text
  const v = sheet.getDataRange().getDisplayValues();
  return v.slice(1).filter(r => r.some(c => c !== ''));
}
function writeAll(sheet, head, data) {         // replace everything below the header
  sheet.clearContents();
  const all = [head].concat(data);
  sheet.getRange(1, 1, all.length, head.length).setValues(all);
  sheet.getRange(1, 1, 1, head.length).setFontWeight('bold');
}

/* ---------- read ---------- */
function readAll() {
  const roster = rows(tab('Students', STUDENT_HEAD)).map(r => ({ id: r[0], name: r[1] }));
  const settings = {};
  rows(tab('Settings', SETTING_HEAD)).forEach(r => { if (r[0] === 'Term start (Monday of week 1)') settings.start = r[1]; });
  const weeks = {};
  rows(tab('Tallies', TALLY_HEAD)).forEach(r => {
    const vid = r[7], sid = r[3];
    if (!vid || !sid) return;
    (weeks[vid] = weeks[vid] || {})[sid] = { a: Number(r[5]) || 0, s: Number(r[6]) || 0 };
  });
  return { roster: roster, settings: settings, weeks: weeks };
}

/* ---------- write ---------- */
function apply(b) {
  const meta = Array.isArray(b.meta) ? b.meta : [];
  const metaById = {};
  meta.forEach(m => { metaById[m.id] = m; });

  if (b.op === 'roster' && Array.isArray(b.roster)) {
    const clean = b.roster.map(st => [String(st.id).slice(0, 20), String(st.name || '').slice(0, 60)]);
    writeAll(tab('Students', STUDENT_HEAD), STUDENT_HEAD, clean);
  }

  if (b.op === 'settings' && b.settings && /^\d{4}-\d{2}-\d{2}$/.test(b.settings.start || '')) {
    const s = tab('Settings', SETTING_HEAD);
    s.getRange('B:B').setNumberFormat('@');   // keep the date as text so Sheets doesn't reformat it
    writeAll(s, SETTING_HEAD, [['Term start (Monday of week 1)', b.settings.start]]);
  }

  // Names follow the Students tab into the Tallies tab.
  const names = {};
  rows(tab('Students', STUDENT_HEAD)).forEach(r => { names[r[0]] = r[1]; });

  const t = tab('Tallies', TALLY_HEAD);
  let data = rows(t);

  if (b.op === 'cells' && b.vid) {
    const vid = String(b.vid);
    if (b.reset) data = data.filter(r => r[7] !== vid);           // "Clear this week"
    const cells = b.cells || {};
    const stamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm');
    Object.keys(cells).forEach(sid => {
      const c = cells[sid] || {};
      const a = Math.max(0, Math.floor(Number(c.a) || 0)), s = Math.max(0, Math.floor(Number(c.s) || 0));
      const m = metaById[vid] || {};
      const row = data.find(r => r[7] === vid && r[3] === sid);
      const vals = [m.week || '', m.name || vid, m.ref || '', sid, names[sid] || '', a, s, vid, stamp];
      if (row) row.splice(0, row.length, ...vals); else data.push(vals);
    });
  }

  data.forEach(r => { if (names[r[3]] !== undefined) r[4] = names[r[3]]; });
  data.sort((x, y) => (Number(x[0]) || 99) - (Number(y[0]) || 99) || sidNum(x[3]) - sidNum(y[3]));
  writeAll(t, TALLY_HEAD, data);

  if (meta.length) buildOverview(meta, names, data);
}

function sidNum(id) { return Number(String(id).replace(/\D/g, '')) || 0; }

/* Overview: a printable class grid. Each cell is "successes / attempts". */
function buildOverview(meta, names, data) {
  const head = ['Student'].concat(meta.map(m => 'Wk ' + m.week + ' ' + m.name)).concat(['Term successes', 'Term attempts']);
  const by = {};
  data.forEach(r => { by[r[7] + '|' + r[3]] = r; });
  const ids = Object.keys(names);
  const body = ids.map(sid => {
    let ts = 0, ta = 0;
    const cells = meta.map(m => {
      const r = by[m.id + '|' + sid];
      if (!r) return '';
      const s = Number(r[6]) || 0, a = Number(r[5]) || 0;
      ts += s; ta += a;
      return (s || a) ? s + ' / ' + a : '';
    });
    return [names[sid] || sid].concat(cells).concat([ts, ta]);
  });
  const o = tab('Overview', head);
  writeAll(o, head, body.concat([[''], ['Each cell shows successes / attempts.']].map(r => r.concat(Array(head.length - r.length).fill('')))));
  o.setFrozenRows(1);
}
