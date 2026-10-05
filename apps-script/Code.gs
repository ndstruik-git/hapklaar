/**
 * Hapklaar backend (Google Apps Script, gebonden aan de Hapklaar-sheet).
 *
 * Eenmalig:
 *   1. Plak dit bestand in Extensies > Apps Script (vervang alles in Code.gs).
 *   2. Kies bovenin de functie `setup` en klik Uitvoeren. Geef toestemming.
 *      In het logboek staat daarna je sleutel.
 *   3. (Optioneel) Projectinstellingen > Scripteigenschappen > CLAUDE_API_KEY toevoegen.
 *   4. Implementeren > Nieuwe implementatie > Web-app.
 *      Uitvoeren als: ik. Toegang: iedereen. Kopieer de web-app-URL.
 *
 * Bij een wijziging in deze code: Implementeren > Implementaties beheren > potlood >
 * Versie: Nieuwe versie. De URL blijft dan hetzelfde.
 */

const TZ = 'Europe/Amsterdam';
const MODEL = 'claude-haiku-4-5-20251001';

const SCHEMA = {
  recepten: { tab: 'Recepten', key: 'id', cols: [['id', 's'], ['naam', 's'], ['tags', 'list'], ['ingredienten', 'list'], ['aangemaakt_op', 's']] },
  ingredienten: { tab: 'Ingredienten', key: 'id', cols: [['id', 's'], ['naam', 's'], ['categorie', 's'], ['categorie_bron', 's'], ['is_voorraad', 'b']] },
  planning: { tab: 'Planning', key: 'datum', sort: true, cols: [['datum', 's'], ['type', 's'], ['recept_id', 's'], ['tekst', 's']] },
  boodschappen: { tab: 'Boodschappen', key: 'id', sort: true, cols: [['id', 's'], ['week_start', 's'], ['ingredient_id', 's'], ['extra', 'b'], ['voorraad_actief', 'b'], ['afgevinkt', 'b'], ['afgevinkt_op', 's']] }
};
const CAT_TAB = 'Categorieen';
const DEFAULT_CATS = ['Groente & fruit', 'Brood & bakkerij', 'Zuivel & eieren', 'Kaas & vleeswaren', 'Vlees & vis', 'Vega', 'Pasta, rijst & wereldkeuken', 'Conserven & sauzen', 'Kruiden & specerijen', 'Ontbijt & beleg', 'Diepvries', 'Snacks', 'Dranken', 'Huishouden & verzorging', 'Overig'];

/* ---------------- HTTP ---------------- */

function doGet(e) {
  return handle_(function () {
    const p = (e && e.parameter) || {};
    if (!p.action) return { ok: true, app: 'hapklaar' };
    auth_(p.k);
    if (p.action === 'load') return load_();
    if (p.action === 'agenda') return agenda_(p.week);
    throw fail_('unknown_action');
  });
}

function doPost(e) {
  return handle_(function () {
    let body;
    try { body = JSON.parse((e && e.postData && e.postData.contents) || '{}'); } catch (err) { throw fail_('bad_json'); }
    auth_(body.k);
    if (body.action === 'save') return save_(body.ops || [], body.categorize || []);
    throw fail_('unknown_action');
  });
}

function handle_(fn) {
  let out;
  try { out = fn(); } catch (err) {
    out = { ok: false, error: err && err.code ? err.code : 'server_error', message: String(err && err.message || err) };
  }
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}
function fail_(code, message) { const e = new Error(message || code); e.code = code; return e; }

function auth_(k) {
  const key = PropertiesService.getScriptProperties().getProperty('APP_KEY');
  if (!key) throw fail_('not_configured', 'Draai eerst setup().');
  if (!k || String(k) !== key) throw fail_('bad_key');
}

/* ---------------- sheet helpers ---------------- */

function ss_() {
  const id = PropertiesService.getScriptProperties().getProperty('SHEET_ID');
  return id ? SpreadsheetApp.openById(id) : SpreadsheetApp.getActiveSpreadsheet();
}

function cell_(v, type) {
  if (type === 'b') return v === true || String(v).toUpperCase() === 'TRUE';
  if (v instanceof Date) return Utilities.formatDate(v, TZ, 'yyyy-MM-dd');
  const s = v === null || v === undefined ? '' : String(v);
  if (type === 'list') return s.split(',').map(function (x) { return x.trim(); }).filter(String);
  return s;
}
function out_(v, type) {
  if (type === 'b') return v === true;
  if (type === 'list') return Array.isArray(v) ? v.map(String).join(',') : String(v || '');
  return v === null || v === undefined ? '' : String(v);
}

/** Reads a tab into an ordered map id -> doc. Columns are found by header name, so reordering columns in the sheet is safe. */
function readTab_(ss, schema) {
  const sh = ss.getSheetByName(schema.tab);
  const map = {};
  if (!sh || sh.getLastRow() < 2) return map;
  const values = sh.getRange(1, 1, sh.getLastRow(), Math.max(sh.getLastColumn(), 1)).getValues();
  const head = values[0].map(function (h) { return String(h).trim(); });
  const idx = schema.cols.map(function (c) { return head.indexOf(c[0]); });
  for (let r = 1; r < values.length; r++) {
    const row = values[r];
    const doc = {};
    schema.cols.forEach(function (c, i) { doc[c[0]] = idx[i] >= 0 ? cell_(row[idx[i]], c[1]) : cell_('', c[1]); });
    const id = doc[schema.key];
    if (id) map[id] = doc;
  }
  return map;
}

function writeTab_(ss, schema, map) {
  let sh = ss.getSheetByName(schema.tab);
  const n = schema.cols.length;
  if (!sh) { sh = ss.insertSheet(schema.tab); textFormat_(sh, n); }
  let ids = Object.keys(map);
  if (schema.sort) ids.sort();
  const rows = [schema.cols.map(function (c) { return c[0]; })];
  ids.forEach(function (id) { const d = map[id]; rows.push(schema.cols.map(function (c) { return out_(d[c[0]], c[1]); })); });
  if (sh.getMaxRows() < rows.length) {
    const add = rows.length - sh.getMaxRows() + 50;
    sh.insertRowsAfter(sh.getMaxRows(), add);
    textFormat_(sh, n);
  }
  const last = sh.getLastRow();
  sh.getRange(1, 1, rows.length, n).setValues(rows);
  if (last > rows.length) sh.getRange(rows.length + 1, 1, last - rows.length, n).clearContent();
}

function textFormat_(sh, n) {
  sh.getRange(1, 1, sh.getMaxRows(), n).setNumberFormat('@');
}

function categories_(ss) {
  const sh = ss.getSheetByName(CAT_TAB);
  if (!sh || sh.getLastRow() < 2) return DEFAULT_CATS.slice();
  const v = sh.getRange(2, 1, sh.getLastRow() - 1, 2).getValues()
    .filter(function (r) { return String(r[0]).trim(); })
    .map(function (r, i) { return { naam: String(r[0]).trim(), volgorde: Number(r[1]) || (i + 1) }; })
    .sort(function (a, b) { return a.volgorde - b.volgorde; })
    .map(function (r) { return r.naam; });
  if (v.indexOf('Overig') < 0) v.push('Overig');
  return v;
}

/* ---------------- API ---------------- */

function load_() {
  const ss = ss_();
  const data = {};
  Object.keys(SCHEMA).forEach(function (col) { data[col] = readTab_(ss, SCHEMA[col]); });
  return { ok: true, data: data, categorieen: categories_(ss), now: new Date().toISOString() };
}

function save_(ops, categorizeIds) {
  if (!Array.isArray(ops) || ops.length > 500) throw fail_('bad_ops');
  const lock = LockService.getScriptLock();
  lock.waitLock(25000);
  try {
    const ss = ss_();
    const touched = {};
    ops.forEach(function (o) {
      const schema = SCHEMA[o && o.col];
      if (!schema) throw fail_('bad_collection', String(o && o.col));
      const id = String(o.id || '');
      if (!id || id.length > 200) throw fail_('bad_id');
      if (!touched[o.col]) touched[o.col] = readTab_(ss, schema);
      if (o.op === 'del') delete touched[o.col][id];
      else if (o.op === 'set') {
        const d = o.data || {};
        const doc = {};
        schema.cols.forEach(function (c) { doc[c[0]] = d[c[0]]; });
        doc[schema.key] = id;
        touched[o.col][id] = doc;
      } else throw fail_('bad_op');
    });
    Object.keys(touched).forEach(function (col) { writeTab_(ss, SCHEMA[col], touched[col]); });

    const res = { ok: true, applied: ops.length };
    if (Array.isArray(categorizeIds) && categorizeIds.length) {
      const r = categorize_(ss, categorizeIds.slice(0, 40));
      if (r.error) res.categorizeError = r.error;
      if (r.updated) res.updated = { ingredienten: r.updated };
    }
    return res;
  } finally {
    lock.releaseLock();
  }
}

/** Asks Claude for one aisle per new ingredient. Only touches rows still marked ai-wacht, never manual choices. */
function categorize_(ss, ids) {
  const apiKey = PropertiesService.getScriptProperties().getProperty('CLAUDE_API_KEY');
  if (!apiKey) return { error: 'no_api_key' };
  const map = readTab_(ss, SCHEMA.ingredienten);
  const todo = ids.map(String).filter(function (id) { return map[id] && map[id].categorie_bron === 'ai-wacht'; });
  if (!todo.length) return { updated: {} };
  const cats = categories_(ss);
  const prompt = 'Je sorteert boodschappen voor een Nederlandse supermarkt. Kies voor elk item precies één categorie uit deze lijst, letterlijk gespeld:\n' +
    cats.join('\n') +
    '\n\nAntwoord alleen met één JSON-object: elke sleutel is het item zoals gegeven, de waarde is de categorie. Voorbeeld: {"yoghurt":"Zuivel & eieren","wc-papier":"Huishouden & verzorging"}\n\nItems:\n' +
    todo.map(function (id) { return map[id].naam; }).join('\n');
  let parsed;
  try {
    const resp = UrlFetchApp.fetch('https://api.anthropic.com/v1/messages', {
      method: 'post',
      contentType: 'application/json',
      headers: { 'x-api-key': apiKey, 'anthropic-version': '2023-06-01' },
      payload: JSON.stringify({ model: MODEL, max_tokens: 1024, messages: [{ role: 'user', content: prompt }] }),
      muteHttpExceptions: true
    });
    if (resp.getResponseCode() !== 200) return { error: 'ai_http_' + resp.getResponseCode() };
    const body = JSON.parse(resp.getContentText());
    const text = (body.content || []).map(function (c) { return c.text || ''; }).join('');
    const m = text.match(/\{[\s\S]*\}/);
    parsed = m ? JSON.parse(m[0]) : null;
  } catch (err) {
    return { error: 'ai_failed' };
  }
  if (!parsed) return { error: 'ai_unparsable' };
  const updated = {};
  todo.forEach(function (id) {
    const c = parsed[map[id].naam];
    if (typeof c === 'string' && cats.indexOf(c) >= 0) {
      map[id].categorie = c;
      map[id].categorie_bron = 'ai';
      updated[id] = map[id];
    }
  });
  if (Object.keys(updated).length) writeTab_(ss, SCHEMA.ingredienten, map);
  return { updated: updated };
}

/** Niels's calendar for one week (Sunday to Saturday), read-only. Cached for 10 minutes. */
function agenda_(week) {
  if (!/^\d{4}-\d{2}-\d{2}$/.test(String(week || ''))) throw fail_('bad_week');
  const cache = CacheService.getScriptCache();
  const hit = cache.get('ag_' + week);
  if (hit) return { ok: true, events: JSON.parse(hit), cached: true };

  const days = [];
  for (let i = 0; i < 7; i++) days.push(addDays_(week, i));
  const inWeek = {};
  days.forEach(function (d) { inWeek[d] = true; });
  const start = Utilities.parseDate(week + ' 00:00', TZ, 'yyyy-MM-dd HH:mm');
  const end = Utilities.parseDate(addDays_(week, 7) + ' 00:00', TZ, 'yyyy-MM-dd HH:mm');
  const events = [];
  CalendarApp.getDefaultCalendar().getEvents(start, end).forEach(function (ev) {
    try { if (ev.getMyStatus() === CalendarApp.GuestStatus.NO) return; } catch (err) { /* not a guest */ }
    const titel = String(ev.getTitle() || '(geen titel)').slice(0, 80);
    if (ev.isAllDayEvent()) {
      let d = Utilities.formatDate(ev.getAllDayStartDate(), TZ, 'yyyy-MM-dd');
      const stop = Utilities.formatDate(ev.getAllDayEndDate(), TZ, 'yyyy-MM-dd');
      for (let g = 0; d < stop && g < 60; g++) {
        if (inWeek[d]) events.push({ datum: d, titel: titel, tijd: null });
        d = addDays_(d, 1);
      }
    } else {
      const s = ev.getStartTime();
      const d = Utilities.formatDate(s, TZ, 'yyyy-MM-dd');
      if (inWeek[d]) events.push({ datum: d, titel: titel, tijd: Utilities.formatDate(s, TZ, 'HH:mm') });
    }
  });
  events.sort(function (a, b) { return a.datum < b.datum ? -1 : a.datum > b.datum ? 1 : String(a.tijd || '') < String(b.tijd || '') ? -1 : 1; });
  try { cache.put('ag_' + week, JSON.stringify(events), 600); } catch (err) { /* too big to cache */ }
  return { ok: true, events: events };
}

function addDays_(ymd, n) {
  const p = ymd.split('-').map(Number);
  const d = new Date(Date.UTC(p[0], p[1] - 1, p[2] + n));
  return d.getUTCFullYear() + '-' + ('0' + (d.getUTCMonth() + 1)).slice(-2) + '-' + ('0' + d.getUTCDate()).slice(-2);
}

/* ---------------- one-time setup ---------------- */

function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const props = PropertiesService.getScriptProperties();
  props.setProperty('SHEET_ID', ss.getId());

  Object.keys(SCHEMA).forEach(function (col) {
    const schema = SCHEMA[col];
    const sh = ss.getSheetByName(schema.tab) || ss.insertSheet(schema.tab);
    const n = schema.cols.length;
    textFormat_(sh, n);
    if (sh.getLastRow() === 0) sh.getRange(1, 1, 1, n).setValues([schema.cols.map(function (c) { return c[0]; })]);
    sh.getRange(1, 1, 1, n).setFontWeight('bold');
    sh.setFrozenRows(1);
  });
  const cs = ss.getSheetByName(CAT_TAB) || ss.insertSheet(CAT_TAB);
  if (cs.getLastRow() < 2) {
    cs.getRange(1, 1, DEFAULT_CATS.length + 1, 2).setValues([['naam', 'volgorde']].concat(DEFAULT_CATS.map(function (c, i) { return [c, i + 1]; })));
  }
  cs.getRange(1, 1, 1, 2).setFontWeight('bold');
  cs.setFrozenRows(1);

  if (!props.getProperty('APP_KEY')) props.setProperty('APP_KEY', Utilities.getUuid().replace(/-/g, ''));

  // Touch the calendar once so the permission prompt includes it.
  CalendarApp.getDefaultCalendar().getName();

  Logger.log('Klaar. Je sleutel: ' + props.getProperty('APP_KEY'));
  Logger.log('Claude-sortering: ' + (props.getProperty('CLAUDE_API_KEY') ? 'aan' : 'uit (voeg CLAUDE_API_KEY toe in Scripteigenschappen)'));
  Logger.log('Volgende stap: Implementeren > Nieuwe implementatie > Web-app (uitvoeren als: ik, toegang: iedereen).');
}

/** Run this if the key ever leaks: old links stop working, then share the new one. */
function nieuweSleutel() {
  const props = PropertiesService.getScriptProperties();
  props.setProperty('APP_KEY', Utilities.getUuid().replace(/-/g, ''));
  Logger.log('Nieuwe sleutel: ' + props.getProperty('APP_KEY'));
}
