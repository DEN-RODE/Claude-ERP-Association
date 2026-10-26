/**
 * harness.js : charge le bundle Apps Script dans un contexte Node et simule les
 * services Google dont il depend.
 *
 * Apps Script n'est pas testable localement : pas de `require`, pas de npm, et un
 * runtime qui n'existe que chez Google. On contourne en trois temps :
 *
 *   1. build.js produit un fichier unique (bundle/Code.gs) qui est du JavaScript
 *      ES5 valide, sans aucun import ;
 *   2. ce harness cree un contexte `vm` ou l'on injecte des doubles des services
 *      (SpreadsheetApp, DriveApp, GmailApp, CacheService, Utilities...) ;
 *   3. le bundle est evalue dans ce contexte, et toutes ses fonctions globales
 *      deviennent appelables depuis les tests.
 *
 * Le double de classeur est une vraie structure en memoire : onglets, lignes,
 * couleurs de fond, formules. Il suffit pour tester la logique metier, qui est la
 * seule chose que l'on veut verifier ici. Ce qui releve purement de l'API Google
 * (un PDF reellement genere, un mail reellement remis) ne se teste pas ainsi et
 * est marque comme tel dans les tests.
 */
const fs = require('fs');
const path = require('path');
const vm = require('vm');

const BUNDLE = path.join(__dirname, '..', 'bundle', 'Code.gs');

// ---------------------------------------------------------------- classeur
class FakeRange {
  constructor(sheet, row, col, nRows, nCols) {
    Object.assign(this, { sheet, row, col, nRows, nCols });
  }
  _cell(r, c) { return this.sheet._cell(this.row + r, this.col + c); }
  getValues() {
    const out = [];
    for (let r = 0; r < this.nRows; r++) {
      const ligne = [];
      for (let c = 0; c < this.nCols; c++) ligne.push(this._cell(r, c).v);
      out.push(ligne);
    }
    return out;
  }
  getValue() { return this._cell(0, 0).v; }
  setValues(vals) {
    vals.forEach((ligne, r) => ligne.forEach((v, c) => { this._cell(r, c).v = v; }));
    return this;
  }
  setValue(v) {
    for (let r = 0; r < this.nRows; r++) for (let c = 0; c < this.nCols; c++) this._cell(r, c).v = v;
    return this;
  }
  getBackgrounds() {
    const out = [];
    for (let r = 0; r < this.nRows; r++) {
      const ligne = [];
      for (let c = 0; c < this.nCols; c++) ligne.push(this._cell(r, c).bg || '#ffffff');
      out.push(ligne);
    }
    return out;
  }
  getBackground() { return this._cell(0, 0).bg || '#ffffff'; }
  setBackground(bg) {
    for (let r = 0; r < this.nRows; r++) for (let c = 0; c < this.nCols; c++) this._cell(r, c).bg = bg;
    return this;
  }
  setFormula(f) { this._cell(0, 0).v = f; return this; }
  getFormula() { const v = this._cell(0, 0).v; return typeof v === 'string' && v[0] === '=' ? v : ''; }
  setNumberFormat() { return this; }
  setFontWeight() { return this; }
  setFontColor() { return this; }
  setFontSize() { return this; }
  setFontStyle() { return this; }
  setHorizontalAlignment() { return this; }
  setVerticalAlignment() { return this; }
  setWrap() { return this; }
  merge() { return this; }
  clearContent() { return this.setValue(''); }
  getA1Notation() { return 'R' + this.row + 'C' + this.col; }
  getNumRows() { return this.nRows; }
  getNumColumns() { return this.nCols; }
}

class FakeSheet {
  constructor(name, rows) {
    this.name = name;
    this.cells = {};          // "r:c" -> { v, bg }
    this.hidden = false;
    (rows || []).forEach((ligne, r) => ligne.forEach((v, c) => { this._cell(r + 1, c + 1).v = v; }));
  }
  _cell(r, c) {
    const k = r + ':' + c;
    if (!this.cells[k]) this.cells[k] = { v: '', bg: '#ffffff' };
    return this.cells[k];
  }
  getName() { return this.name; }
  setName(n) { this.name = n; return this; }
  getLastRow() {
    let max = 0;
    for (const k of Object.keys(this.cells)) {
      const [r] = k.split(':').map(Number);
      if (this.cells[k].v !== '' && this.cells[k].v !== null && r > max) max = r;
    }
    return max;
  }
  getLastColumn() {
    let max = 0;
    for (const k of Object.keys(this.cells)) {
      const [, c] = k.split(':').map(Number);
      if (this.cells[k].v !== '' && this.cells[k].v !== null && c > max) max = c;
    }
    return max;
  }
  getMaxRows() { return Math.max(this.getLastRow(), 1000); }
  getMaxColumns() { return Math.max(this.getLastColumn(), 26); }
  getRange(a, b, c, d) {
    if (typeof a === 'string') {              // 'B:B', 'A1', 'A1:C3'
      const m = a.match(/^([A-Z]+)(\d*):?([A-Z]*)(\d*)$/);
      const col = (s) => s.split('').reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0);
      const c1 = col(m[1]), r1 = m[2] ? +m[2] : 1;
      const c2 = m[3] ? col(m[3]) : c1, r2 = m[4] ? +m[4] : (m[2] ? r1 : this.getMaxRows());
      return new FakeRange(this, r1, c1, r2 - r1 + 1, c2 - c1 + 1);
    }
    return new FakeRange(this, a, b, c === undefined ? 1 : c, d === undefined ? 1 : d);
  }
  appendRow(vals) {
    const r = this.getLastRow() + 1;
    vals.forEach((v, c) => { this._cell(r, c + 1).v = v; });
    return this;
  }
  deleteRow(r) {
    const last = this.getLastRow();
    for (let i = r; i < last; i++) {
      for (let c = 1; c <= this.getLastColumn(); c++) {
        const src = this._cell(i + 1, c);
        const dst = this._cell(i, c);
        dst.v = src.v; dst.bg = src.bg;
      }
    }
    for (let c = 1; c <= this.getLastColumn(); c++) { const x = this._cell(last, c); x.v = ''; x.bg = '#ffffff'; }
    return this;
  }
  insertRowBefore() { return this; }
  insertRowsAfter() { return this; }
  setFrozenRows() { return this; }
  setColumnWidth() { return this; }
  autoResizeColumn() { return this; }
  hideSheet() { this.hidden = true; return this; }
  getDataRange() { return this.getRange(1, 1, Math.max(this.getLastRow(), 1), Math.max(this.getLastColumn(), 1)); }
  getSheetId() { return this.name.length; }
  clear() { this.cells = {}; return this; }
}

class FakeSpreadsheet {
  constructor(id, onglets) {
    this.id = id;
    this.sheets = Object.keys(onglets || {}).map((n) => new FakeSheet(n, onglets[n]));
  }
  getId() { return this.id; }
  getName() { return 'Classeur ' + this.id; }
  getUrl() { return 'https://example.org/sheets/' + this.id; }
  getSheets() { return this.sheets; }
  getSheetByName(n) { return this.sheets.find((s) => s.name === n) || null; }
  insertSheet(n) { const s = new FakeSheet(n, []); this.sheets.push(s); return s; }
  deleteSheet(s) { this.sheets = this.sheets.filter((x) => x !== s); }
  getActiveSheet() { return this.sheets[0]; }
  getSpreadsheetTimeZone() { return 'Europe/Paris'; }
}

// ---------------------------------------------------------------- services
function services(etat) {
  const classeurs = etat.classeurs;

  const SpreadsheetApp = {
    openById: (id) => {
      if (!classeurs[id]) throw new Error('Classeur introuvable : ' + id);
      return classeurs[id];
    },
    getActiveSpreadsheet: () => classeurs[etat.erpId] || null,
    getUi: () => ({
      createMenu: () => { const m = { addItem: () => m, addSeparator: () => m, addToUi: () => m }; return m; },
      alert: (...a) => { etat.alertes.push(a); },
      ButtonSet: { OK: 'OK' }
    }),
    flush: () => { },
    newDataValidation: () => ({ requireValueInList: () => ({ build: () => ({}) }) })
  };

  const Utilities = {
    getUuid: () => {
      etat.uuid += 1;
      return ('00000000' + etat.uuid).slice(-8) + '-0000-0000-0000-000000000000';
    },
    formatDate: (d, tz, fmt) => {
      const p = (n, l) => ('0000' + n).slice(-(l || 2));
      return fmt
        .replace('yyyy', d.getFullYear())
        .replace('MM', p(d.getMonth() + 1))
        .replace('dd', p(d.getDate()))
        .replace('HH', p(d.getHours()))
        .replace('mm', p(d.getMinutes()))
        .replace('ss', p(d.getSeconds()));
    },
    sleep: () => { },
    base64Encode: (s) => Buffer.from(s).toString('base64'),
    base64Decode: (s) => Array.from(Buffer.from(s, 'base64')),
    newBlob: (bytes, type, name) => ({ bytes, type, name, getBytes: () => bytes, getName: () => name }),
    // vrai SHA-256 : le controle du code PIN compare des empreintes, un double qui
    // renvoie une valeur constante ferait passer n'importe quel code
    computeDigest: (algo, valeur) => {
      const nom = { SHA_256: 'sha256', SHA_1: 'sha1', MD5: 'md5' }[algo] || 'sha256';
      const buf = require('crypto').createHash(nom).update(String(valeur), 'utf8').digest();
      return Array.from(buf).map((b) => (b > 127 ? b - 256 : b)); // octets signes, comme Apps Script
    },
    DigestAlgorithm: { SHA_256: 'SHA_256', SHA_1: 'SHA_1', MD5: 'MD5' },
    Charset: { UTF_8: 'UTF_8' },
    parseCsv: (t) => t.split('\n').map((l) => l.split(','))
  };

  const PropertiesService = {
    getScriptProperties: () => ({
      getProperty: (k) => (k in etat.props ? etat.props[k] : null),
      setProperty: (k, v) => { etat.props[k] = String(v); },
      deleteProperty: (k) => { delete etat.props[k]; },
      getProperties: () => Object.assign({}, etat.props)
    })
  };

  const CacheService = {
    getScriptCache: () => ({
      get: (k) => (k in etat.cache ? etat.cache[k] : null),
      put: (k, v) => { etat.cache[k] = v; },
      remove: (k) => { delete etat.cache[k]; },
      removeAll: (ks) => ks.forEach((k) => delete etat.cache[k]),
      getAll: (ks) => { const o = {}; ks.forEach((k) => { if (k in etat.cache) o[k] = etat.cache[k]; }); return o; }
    })
  };

  const GmailApp = {
    createDraft: (to, sujet, corps, opts) => {
      const d = { id: 'd' + (etat.brouillons.length + 1), to, sujet, corps, opts };
      etat.brouillons.push(d);
      return {
        getId: () => d.id,
        getMessage: () => ({ getSubject: () => sujet, getTo: () => to, getPlainBody: () => corps, getId: () => d.id }),
        send: () => { etat.envoyes.push(d); },
        deleteDraft: () => { etat.brouillons = etat.brouillons.filter((x) => x !== d); }
      };
    },
    getDrafts: () => etat.brouillons.map((d) => ({
      getId: () => d.id,
      getMessage: () => ({ getSubject: () => d.sujet, getTo: () => d.to, getPlainBody: () => d.corps, getId: () => d.id })
    })),
    sendEmail: (to, sujet, corps, opts) => { etat.envoyes.push({ to, sujet, corps, opts }); },
    search: () => [],
    getUserLabelByName: () => null,
    createLabel: (n) => ({ getName: () => n, addToThread: () => { } })
  };

  const dossier = (id, nom) => ({
    getId: () => id,
    getName: () => nom,
    getUrl: () => 'https://example.org/drive/' + id,
    getFoldersByName: (n) => {
      const k = id + '/' + n;
      if (!etat.dossiers[k]) etat.dossiers[k] = dossier(k, n);
      return { hasNext: () => true, next: () => etat.dossiers[k] };
    },
    createFolder: (n) => { const k = id + '/' + n; etat.dossiers[k] = dossier(k, n); return etat.dossiers[k]; },
    getFiles: () => ({ hasNext: () => false, next: () => null }),
    getFilesByName: () => ({ hasNext: () => false, next: () => null }),
    createFile: (b) => ({ getId: () => 'f' + (etat.fichiers.push(b)), getUrl: () => 'https://example.org/file' })
  });

  const DriveApp = {
    getRootFolder: () => dossier('root', 'root'),
    getFolderById: (id) => dossier(id, 'dossier ' + id),
    getFileById: (id) => ({
      getId: () => id,
      getName: () => 'fichier ' + id,
      getUrl: () => 'https://example.org/file/' + id,
      getBlob: () => ({ getBytes: () => [], getContentType: () => 'image/png' }),
      getParents: () => ({ hasNext: () => true, next: () => dossier('parent', 'parent') }),
      getOwner: () => ({ getEmail: () => etat.utilisateur }),
      makeCopy: (nom, dst) => ({ getId: () => 'copie-' + id, getName: () => nom, getUrl: () => 'https://example.org/copie', moveTo: () => { }, setName: () => { } }),
      moveTo: () => { },
      setTrashed: () => { }
    }),
    createFile: (b) => ({ getId: () => 'f' + (etat.fichiers.push(b)), getUrl: () => 'https://example.org/file' })
  };

  const Session = {
    getActiveUser: () => ({ getEmail: () => etat.utilisateur }),
    getEffectiveUser: () => ({ getEmail: () => etat.utilisateur }),
    getTemporaryActiveUserKey: () => 'cle-' + etat.utilisateur,
    getScriptTimeZone: () => 'Europe/Paris'
  };

  const ScriptApp = {
    getProjectTriggers: () => etat.triggers,
    deleteTrigger: (t) => { etat.triggers = etat.triggers.filter((x) => x !== t); },
    newTrigger: (fn) => {
      const t = { fn, uid: 'u' + (etat.triggers.length + 1) };
      const b = {
        forSpreadsheet: () => b, onFormSubmit: () => b, timeBased: () => b,
        everyDays: () => b, everyMinutes: () => b, atHour: () => b, inTimezone: () => b, nearMinute: () => b,
        create: () => { etat.triggers.push({ getHandlerFunction: () => fn, getUniqueId: () => t.uid }); return t; }
      };
      return b;
    }
  };

  return {
    SpreadsheetApp, Utilities, PropertiesService, CacheService, GmailApp, DriveApp, Session, ScriptApp,
    DocumentApp: {
      openById: () => ({
        getBody: () => ({ replaceText: () => { }, getText: () => '', appendParagraph: () => ({ setHeading: () => { } }), appendTable: () => ({ getRow: () => ({ getCell: () => ({ setText: () => { } }) }) }) }),
        saveAndClose: () => { }, getId: () => 'doc', getName: () => 'doc'
      }),
      ParagraphHeading: { HEADING1: 1, HEADING2: 2, NORMAL: 0 }
    },
    CalendarApp: { getDefaultCalendar: () => ({ getEvents: () => [], createEvent: () => ({ getId: () => 'ev' }) }), getCalendarById: () => null },
    HtmlService: {
      createTemplateFromFile: () => ({ evaluate: () => ({ setTitle: () => ({ addMetaTag: () => ({ setXFrameOptionsMode: () => 'page' }) }) }) }),
      createHtmlOutputFromFile: () => ({ getContent: () => '' }),
      XFrameOptionsMode: { DEFAULT: 'DEFAULT' }
    },
    Logger: { log: (...a) => etat.logs.push(a) },
    MailApp: { sendEmail: (o) => etat.envoyes.push(o) },
    LockService: { getScriptLock: () => ({ tryLock: () => true, releaseLock: () => { } }) },
    FormApp: { openById: () => ({ getItems: () => [] }) },
    console
  };
}

/**
 * Prepare un ERP de test.
 *   onglets : { Config: [[...]], Cours: [[...]], ... } lignes brutes, en-tete inclus
 *   autres  : autres classeurs par identifiant (reponses de formulaire, modele eco)
 */
function charger({ onglets = {}, autres = {}, utilisateur = 'test@example.org' } = {}) {
  const erpId = 'ERP_TEST';
  const etat = {
    erpId,
    classeurs: {},
    props: { ERP_SS_ID: erpId },
    cache: {},
    brouillons: [],
    envoyes: [],
    dossiers: {},
    fichiers: [],
    triggers: [],
    logs: [],
    alertes: [],
    uuid: 0,
    utilisateur
  };
  etat.classeurs[erpId] = new FakeSpreadsheet(erpId, onglets);
  for (const id of Object.keys(autres)) etat.classeurs[id] = new FakeSpreadsheet(id, autres[id]);

  const contexte = vm.createContext(Object.assign({ globalThis: null }, services(etat)));
  contexte.globalThis = contexte;
  vm.runInContext(fs.readFileSync(BUNDLE, 'utf8'), contexte, { filename: 'Code.gs' });
  return { erp: contexte, etat, classeur: etat.classeurs[erpId] };
}

/** Onglet Config minimal : [['Cle','Valeur','Description'], ...] */
function config(paires) {
  return [['Cle', 'Valeur', 'Description']].concat(Object.keys(paires).map((k) => [k, paires[k], '']));
}

module.exports = { charger, config, FakeSheet, FakeSpreadsheet };
