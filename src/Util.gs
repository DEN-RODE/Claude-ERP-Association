/**
 * Util.gs : fonctions communes (normalisation, dates, feuilles, log)
 */

var TZ = 'Europe/Paris';

/** minuscules, sans accents, apostrophes unifiées, espaces compactés */
function norm_(s) {
  if (s === null || s === undefined) return '';
  return String(s)
    .normalize('NFD').replace(/[\u0300-\u036f]/g, '')
    .replace(/[\u2019\u2018`]/g, "'")
    .replace(/[\u2013\u2014]/g, '-')
    .replace(/\u00a0/g, ' ')
    .toLowerCase()
    .replace(/\s+/g, ' ')
    .trim();
}

/** Date -> 'dd/MM/yyyy' ; string laissée telle quelle */
function fmtDate_(d) {
  if (!d) return '';
  if (Object.prototype.toString.call(d) === '[object Date]') {
    if (isNaN(d.getTime())) return '';
    return Utilities.formatDate(d, TZ, 'dd/MM/yyyy');
  }
  return String(d);
}

function fmtDateTime_(d) {
  if (!d) return '';
  if (Object.prototype.toString.call(d) === '[object Date]') {
    if (isNaN(d.getTime())) return '';
    return Utilities.formatDate(d, TZ, 'dd/MM/yyyy HH:mm');
  }
  return String(d);
}

function isDate_(d) {
  return Object.prototype.toString.call(d) === '[object Date]' && !isNaN(d.getTime());
}

/** parse 'dd/MM/yyyy', 'yyyy-MM-dd', 'M/d/yyyy' (export anglais) ou Date */
function parseDate_(v) {
  if (!v) return null;
  if (isDate_(v)) return v;
  var s = String(v).trim();
  var m = s.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (m) return new Date(+m[1], +m[2] - 1, +m[3]);
  m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/);
  if (m) return new Date(+m[3], +m[2] - 1, +m[1]);
  return null;
}

function ageAt_(naissance, ref) {
  var d = parseDate_(naissance);
  if (!d) return '';
  ref = ref || new Date();
  var a = ref.getFullYear() - d.getFullYear();
  var mo = ref.getMonth() - d.getMonth();
  if (mo < 0 || (mo === 0 && ref.getDate() < d.getDate())) a--;
  return (a >= 0 && a < 120) ? a : '';
}

/** Rend une valeur sérialisable pour google.script.run (pas de Date) */
function plain_(v) {
  if (isDate_(v)) return fmtDate_(v);
  if (v === null || v === undefined) return '';
  return v;
}

function num_(v) {
  if (typeof v === 'number') return v;
  if (v === null || v === undefined || v === '') return 0;
  var s = String(v).replace(/[€\s\u00a0]/g, '').replace(/\((.*)\)/, '-$1').replace(',', '.');
  var n = parseFloat(s);
  return isNaN(n) ? 0 : n;
}

function uid_() {
  return 'R' + Utilities.getUuid().replace(/-/g, '').slice(0, 9);
}

/**
 * Texte écrit dans un Sheet : on force le mode texte quand Sheets le transformerait
 * (0612345678 -> 612345678, =... -> formule, 2026-001 -> date).
 */
function safeCell_(v) {
  if (typeof v !== 'string') return v;
  var t = v.trim();
  if (t === '') return v;
  if (/^[=+\-@]/.test(t) || /^\d/.test(t)) return "'" + v;
  return v;
}

/** Classeur ERP (celui auquel le script est lié) */
function erpSS_() {
  var id = PropertiesService.getScriptProperties().getProperty('ERP_SS_ID');
  if (id) return SpreadsheetApp.openById(id);
  return SpreadsheetApp.getActiveSpreadsheet();
}

function erpSheet_(name) {
  var sh = erpSS_().getSheetByName(name);
  if (!sh) throw new Error('Onglet "' + name + '" introuvable dans le classeur ERP. Lance setup().');
  return sh;
}

/** Lit un onglet ERP simple (en-têtes ligne 1) en tableau d'objets {_row, ...} */
function readTable_(name) {
  var sh = erpSheet_(name);
  var values = sh.getDataRange().getValues();
  if (values.length < 1) return [];
  var h = values[0].map(String);
  var out = [];
  for (var i = 1; i < values.length; i++) {
    var r = values[i];
    if (r.join('') === '') continue;
    var o = { _row: i + 1 };
    for (var j = 0; j < h.length; j++) if (h[j]) o[h[j]] = plain_(r[j]);
    out.push(o);
  }
  return out;
}

function headersOf_(name) {
  var sh = erpSheet_(name);
  return sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(String);
}

/** upsert par colonne ID */
function upsertRow_(name, obj, idCol) {
  idCol = idCol || 'ID';
  var sh = erpSheet_(name);
  var h = headersOf_(name);
  // valeurs BRUTES (les dates restent des dates, les nombres des nombres)
  var values = sh.getDataRange().getValues();
  var iId = h.indexOf(idCol);
  var foundRow = -1;
  for (var i = 1; i < values.length; i++) if (String(values[i][iId]) === String(obj[idCol])) { foundRow = i; break; }
  var line = h.map(function (k, j) {
    if (obj.hasOwnProperty(k)) return safeCell_(obj[k]);
    return foundRow >= 0 ? safeCell_(values[foundRow][j]) : '';
  });
  if (foundRow >= 0) sh.getRange(foundRow + 1, 1, 1, h.length).setValues([line]);
  else sh.appendRow(line);
  return obj;
}

function deleteRowById_(name, id, idCol) {
  idCol = idCol || 'ID';
  var rows = readTable_(name);
  for (var i = rows.length - 1; i >= 0; i--) {
    if (String(rows[i][idCol]) === String(id)) {
      erpSheet_(name).deleteRow(rows[i]._row);
      return true;
    }
  }
  return false;
}

function logAction_(action, detail) {
  try {
    var sh = erpSS_().getSheetByName('Log');
    if (!sh) return;
    sh.appendRow([new Date(), currentUserLabel_(), safeCell_(String(action)), safeCell_(String(detail || '').slice(0, 500))]);
  } catch (e) { /* le log ne doit jamais bloquer */ }
}

function currentUserLabel_() {
  try { return Session.getActiveUser().getEmail() || 'ERP'; } catch (e) { return 'ERP'; }
}

function escapeRegex_(s) {
  return String(s).replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}
