/**
 * Presences.gs : appel depuis l'ERP -> fiches de présence (Fiches_presence / Stage_Fiches_presence).
 *
 * Les noms des fiches se remplissent déjà tout seuls (IMPORTRANGE depuis les Sheets réponses),
 * mais les cases ☐ sont du texte posé par position de ligne : si un élève arrive / part / change
 * de cours, les noms glissent et les croix ne sont plus en face du bon enfant.
 * L'ERP garde donc chaque présence dans son onglet "Presences" (clé = cours + séance + nom + prénom)
 * et réécrit la grille de la fiche à chaque appel, en face du bon nom.
 *
 * Marques écrites : ☑ présent, ✗ absent, E excusé, Es cours d'essai, ☐ rien (un autre texte tapé à la main est gardé).
 */

var MARQUES = { P: '☑', A: '✗', E: 'E', S: 'Es' };

/** Lit une case de la fiche : P / A / E / S (essai), '' si vide, 'T:texte' si autre chose (gardé tel quel) */
function lireMarque_(v) {
  var s = String(v === null || v === undefined ? '' : v).trim();
  if (!s || s === '☐') return '';
  if (/^[☑✓✔✅vV1]$/.test(s) || /^pr[eé]sent/i.test(s)) return 'P';
  if (/^[☒✗✘xX]$/.test(s) || /^abs/i.test(s)) return 'A';
  if (/^e$/i.test(s) || /^exc/i.test(s)) return 'E';
  if (/^es$/i.test(s) || /^essai/i.test(s)) return 'S';
  return 'T:' + s.slice(0, 40);
}

function ecrireMarque_(statut) {
  if (!statut) return '☐';
  if (MARQUES[statut]) return MARQUES[statut];
  if (statut.indexOf('T:') === 0) return safeCell_(statut.slice(2));
  return '☐';
}

/**
 * Trouve l'onglet de fiche d'un cours (B4 = clé) ou d'un stage (A2 contient les dates, M4 le groupe).
 * Rapide : l'onglet trouvé est mémorisé 6 h (sinon on relit l'en-tête de chaque onglet),
 * et tout l'onglet est lu en UNE fois (en-tête + noms + cases).
 */
var FICHE_LARG = 24;
var _ficheSS = {}, _ficheSheets = {}, _presMemo = null;
function ficheTab_(source, cle) {
  var c = getConfig_();
  var id = source === 'cours' ? c.ID_FICHES_COURS : c.ID_FICHES_STAGES;
  if (!id) throw new Error('ID de la fiche de présence manquant dans Config.');
  var g = source === 'cours' ? coursList_().filter(function (x) { return x.cle === cle; })[0] : stagesList_().filter(function (x) { return x.id === cle; })[0];
  if (!g) throw new Error('Cours / stage inconnu.');
  var ss = _ficheSS[id] || (_ficheSS[id] = SpreadsheetApp.openById(id));
  var sheets = _ficheSheets[id] || (_ficheSheets[id] = ss.getSheets());
  var cache = CacheService.getScriptCache();
  var ck = 'ficheTab_' + id + '_' + source + '|' + cle;
  var connu = cache.get(ck);
  if (connu) sheets = sheets.filter(function (x) { return String(x.getSheetId()) === connu; }).concat(sheets.filter(function (x) { return String(x.getSheetId()) !== connu; }));
  function correspond(head) {
    if (source === 'cours') return norm_(head[3][1]) === norm_(g.cle);
    var a2 = norm_(head[1][0]), m4 = norm_(head[3][12] || '');
    return (' ' + a2).indexOf(' ' + norm_(g.cle)) >= 0 && m4.indexOf(g.groupe) >= 0;
  }
  for (var i = 0; i < sheets.length; i++) {
    var sh = sheets[i];
    var larg = Math.min(sh.getMaxColumns(), FICHE_LARG);
    var head = sh.getRange(1, 1, 6, Math.min(larg, 20)).getDisplayValues();
    if (!correspond(head)) continue;
    cache.put(ck, String(sh.getSheetId()), 21600);
    var last = Math.max(sh.getLastRow(), 6);
    var disp = sh.getRange(1, 1, last, larg).getDisplayValues();
    // colonnes de séances : ligne 6, de E jusqu'à "Observations"
    var cols = [];
    for (var j = 4; j < disp[5].length; j++) {
      var lab = String(disp[5][j]).trim();
      if (!lab || /^observ/i.test(lab)) break;
      cols.push({ col: j + 1, label: lab, date: String(disp[4][j] || '').trim() });
    }
    // lignes élèves : à partir de 7, tant que la colonne A est un numéro
    var rows = [];
    for (var r = 6; r < disp.length; r++) {
      if (!/^\d+$/.test(String(disp[r][0]).trim())) break;
      rows.push({ row: r + 1, nom: String(disp[r][1]).trim(), prenom: String(disp[r][2]).trim() });
    }
    return { ss: ss, sheet: sh, onglet: sh.getName(), cols: cols, rows: rows, groupe: g, disp: disp };
  }
  throw new Error('Onglet de fiche de présence introuvable pour « ' + (g.libelle || g.nom) + ' ».');
}

/**
 * Dates des séances d'un cours annuel : S1 = semaine de DEBUT_SAISON_COURS (lundi 7 sept. 2026),
 * puis une par semaine, en sautant les jours listés dans VACANCES_COURS (vacances zone A, fériés).
 */
var JOURS_SEM = ['lundi', 'mardi', 'mercredi', 'jeudi', 'vendredi', 'samedi', 'dimanche'];
function periodesVacances_() {
  return String(getConfig_().VACANCES_COURS || '').split(/[;\n]+/).map(function (x) {
    var m = x.trim().split(/\s*>\s*/);
    var a = parseDate_(m[0]), b = parseDate_(m[1] || m[0]);
    return a && b ? [a.getTime(), b.getTime()] : null;
  }).filter(Boolean);
}
function datesSeances_(jour, n) {
  var debut = parseDate_(getConfig_().DEBUT_SAISON_COURS);
  var off = JOURS_SEM.indexOf(norm_(jour));
  if (!debut || off < 0) return [];
  var lundi = new Date(debut.getFullYear(), debut.getMonth(), debut.getDate() - ((debut.getDay() + 6) % 7));
  var vac = periodesVacances_();
  var out = [];
  for (var w = 0; out.length < n && w < 60; w++) {
    var d = new Date(lundi.getFullYear(), lundi.getMonth(), lundi.getDate() + w * 7 + off);
    var t = d.getTime();
    if (vac.some(function (v) { return t >= v[0] && t <= v[1]; })) continue;
    out.push(d);
  }
  return out;
}

function clePresence_(source, cle, seance, nom, prenom) {
  return [source, cle, seance, norm_(nom), norm_(prenom)].join('|');
}

/** Index des présences de l'ERP : clé -> {row, statut} */
function presSheet_() {
  var ss = erpSS_();
  var sh = ss.getSheetByName('Presences');
  if (!sh) { sh = ss.insertSheet('Presences'); sh.getRange(1, 1, 1, TABLES.Presences.length).setValues([TABLES.Presences]); }
  return sh;
}

function presencesIndex_() {
  if (_presMemo) return _presMemo;
  var sh = presSheet_();
  var v = sh.getDataRange().getValues();
  var idx = {};
  for (var i = 1; i < v.length; i++) if (v[i][0]) idx[String(v[i][0])] = { row: i + 1, statut: String(v[i][6] || ''), par: String(v[i][8] || ''), date: v[i][7] instanceof Date ? Utilities.formatDate(v[i][7], TZ, 'dd/MM/yyyy') : '' };
  _presMemo = idx;
  return idx;
}

/** Écrit un lot de présences dans l'onglet Presences (mise à jour ou ajout) */
function ecrirePresences_(items, source, cle, onglet) {
  if (!items.length) return;
  _presMemo = null;
  var sh = presSheet_();
  var idx = presencesIndex_();
  var ajout = [];
  var qui = currentUserLabel_();
  items.forEach(function (it) {
    var k = clePresence_(source, cle, it.seance, it.nom, it.prenom);
    var line = [k, source, safeCell_(String(cle)), safeCell_(onglet), safeCell_(it.seance), safeCell_(it.nom), safeCell_(it.statut), new Date(), it.par || qui, safeCell_(it.prenom)];
    if (idx[k]) sh.getRange(idx[k].row, 1, 1, line.length).setValues([line]);
    else ajout.push(line);
  });
  if (ajout.length) sh.getRange(sh.getLastRow() + 1, 1, ajout.length, ajout[0].length).setValues(ajout);
  _presMemo = null;
}

/**
 * Synchronise une fiche : reprend les cases cochées à la main dans le Sheet (si les noms n'ont pas bougé
 * depuis la dernière écriture, ou au tout premier usage), puis réécrit toute la grille depuis l'ERP.
 */
function syncFiche_(source, cle, dejaLue) {
  var f = dejaLue || ficheTab_(source, cle);
  if (!f.rows.length || !f.cols.length) return f;
  var props = PropertiesService.getScriptProperties();
  var etatKey = 'fiche_' + f.ss.getId() + '_' + f.sheet.getSheetId();
  var layout = f.rows.map(function (r) { return norm_(r.nom) + '|' + norm_(r.prenom); }).join(';');
  var etat = null;
  try { etat = JSON.parse(props.getProperty(etatKey) || 'null'); } catch (e) { etat = null; }
  var idx = presencesIndex_();
  var prefix = [source, cle].join('|') + '|';
  var dejaUtilise = Object.keys(idx).some(function (k) { return k.indexOf(prefix) === 0; });

  var c0 = f.cols[0].col, nc = f.cols.length;
  var grid = f.rows.map(function (r) { return f.disp[r.row - 1].slice(c0 - 1, c0 - 1 + nc); });

  // 1) reprise de ce qui a été coché / modifié / effacé directement dans le Sheet
  //    - tout 1er usage : toutes les cases remplies
  //    - ensuite : seulement les cases différentes de ce que l'ERP a écrit la dernière fois,
  //      et seulement si les noms n'ont pas bougé entre temps (sinon on ne sait pas à qui c'est)
  var memeNoms = etat && etat.layout === layout && etat.grid && etat.grid.length === grid.length;
  if ((!etat && !dejaUtilise) || memeNoms) {
    var aReprendre = [];
    f.rows.forEach(function (r, i) {
      if (!r.nom && !r.prenom) return;
      f.cols.forEach(function (col, j) {
        var vu = String(grid[i][j]);
        if (memeNoms && String((etat.grid[i] || [])[j]) === vu) return;
        var m = lireMarque_(vu);
        if (!memeNoms && !m) return;
        var k = clePresence_(source, cle, col.label, r.nom, r.prenom);
        var connu = idx[k] ? idx[k].statut : '';
        if (m !== connu) aReprendre.push({ seance: col.label, nom: r.nom, prenom: r.prenom, statut: m, par: 'fiche (à la main)' });
      });
    });
    ecrirePresences_(aReprendre, source, cle, f.onglet);
    if (aReprendre.length) idx = presencesIndex_();
  }

  // 2) réécriture de la grille en face des noms actuels
  var out = f.rows.map(function (r) {
    return f.cols.map(function (col) {
      if (!r.nom && !r.prenom) return '☐';
      var p = idx[clePresence_(source, cle, col.label, r.nom, r.prenom)];
      return ecrireMarque_(p ? p.statut : '');
    });
  });
  var affiche = out.map(function (line) { return line.map(function (v) { return String(v).replace(/^'/, ''); }); });
  var change = affiche.some(function (line, i) { return line.some(function (v, j) { return v !== String(grid[i][j]); }); });
  if (change) {
    f.sheet.getRange(f.rows[0].row, c0, f.rows.length, nc).setValues(out);
    // la copie en mémoire suit ce qui vient d'être écrit (évite de relire l'onglet)
    f.rows.forEach(function (r, i) { for (var j = 0; j < nc; j++) f.disp[r.row - 1][c0 - 1 + j] = affiche[i][j]; });
  }
  props.setProperty(etatKey, JSON.stringify({ layout: layout, grid: affiche }));
  f.idx = idx;
  return f;
}

/** Ouvre l'appel d'un cours / stage : séances, élèves de la fiche, marques déjà connues */
/** À la connexion : tous les appels des cours annuels + les stages de la semaine, en un seul appel serveur */
/**
 * Toutes les fiches d'appel d'un coup. p.jour limite au jour demandé (« Mercredi ») :
 * c'est ce que fait le préchargement après connexion, pour ne pas lire les 9 cours
 * alors qu'on n'en ouvrira qu'un ou deux dans la journée.
 */
function apiAppelsTous_(p) {
  var out = {};
  var jour = p && p.jour ? norm_(p.jour) : '';
  var lock = LockService.getScriptLock(); lock.waitLock(30000);
  try {
    coursList_().forEach(function (k) {
      if (jour && norm_(k.jour) !== jour) return;
      try { out['cours|' + k.cle] = appelData_({ source: 'cours', cle: k.cle }); } catch (e) { }
    });
    var now = Date.now();
    stagesList_().forEach(function (st) {
      var d = parseDate_(st.debut);
      if (!d || Math.abs(d.getTime() - now) > 8 * 86400000) return;
      try { out['stages|' + st.id] = appelData_({ source: 'stages', cle: st.id }); } catch (e) { }
    });
  } finally { lock.releaseLock(); }
  return out;
}

function apiAppel_(p) {
  var lock = LockService.getScriptLock(); lock.waitLock(20000);
  try { return appelData_(p); } finally { lock.releaseLock(); }
}

/**
 * Élèves à NE PAS afficher dans l'appel : statut Désinscrit dans le suivi ERP
 * + onglet "desistement" du Sheet réponses. (Ils restent dans la fiche de présence,
 * qui est alimentée par IMPORTRANGE, mais l'ERP ne les propose plus.)
 */
var _exclusAppel = {};
function exclusAppel_(source) {
  if (_exclusAppel[source]) return _exclusAppel[source];
  var set = {};
  try {
    loadEleves_(source).forEach(function (e) {
      if (e.statut === 'Désinscrit') set[norm_(e.nom) + '|' + norm_(e.prenom)] = 'Désinscrit';
    });
  } catch (e) { }
  try {
    desistements_(source).forEach(function (d) { set[norm_(d.nom) + '|' + norm_(d.prenom)] = 'Désistement'; });
  } catch (e) { }
  _exclusAppel[source] = set;
  return set;
}

function appelData_(p) {
  var f = syncFiche_(p.source, p.cle);
  var idx = f.idx || presencesIndex_();
  var hors = exclusAppel_(p.source);
  var masques = [];
  var lignes = f.rows.filter(function (r) {
    if (!r.nom && !r.prenom) return false;
    var m = hors[norm_(r.nom) + '|' + norm_(r.prenom)];
    if (m) { masques.push(String(r.prenom + ' ' + r.nom).trim() + ' (' + m + ')'); return false; }
    return true;
  });
  var eleves = lignes.map(function (r) {
    var marks = {};
    f.cols.forEach(function (col) {
      var x = idx[clePresence_(p.source, p.cle, col.label, r.nom, r.prenom)];
      if (x && x.statut) marks[col.label] = x.statut;
    });
    return { nom: r.nom, prenom: r.prenom, marks: marks };
  });
  // séance proposée : stage = jour du jour s'il y est, sinon 1re séance sans aucune marque
  var now = new Date();
  var today = Utilities.formatDate(now, TZ, 'dd/MM'), todayFull = Utilities.formatDate(now, TZ, 'dd/MM/yyyy');
  var defaut = '';
  // cours annuel : la séance de la semaine en cours (S1 = semaine du 7 septembre), dates affichées
  if (p.source === 'cours') {
    var dates = datesSeances_(f.groupe.jour, f.cols.length);
    if (dates.length) {
      var j0 = new Date(now.getFullYear(), now.getMonth(), now.getDate());
      var lundi = new Date(j0.getFullYear(), j0.getMonth(), j0.getDate() - ((j0.getDay() + 6) % 7)).getTime();
      f.cols.forEach(function (c, i) { if (dates[i] && !c.date) c.date = Utilities.formatDate(dates[i], TZ, 'dd/MM'); });
      var k = -1;
      for (var i = 0; i < dates.length; i++) if (dates[i].getTime() >= lundi) { k = i; break; }
      if (k < 0 && dates.length >= f.cols.length) k = f.cols.length - 1;
      if (k >= 0 && f.cols[k]) defaut = f.cols[k].label;
    }
  }
  // stage : la colonne du jour ; cours : la séance déjà commencée aujourd'hui dans l'ERP
  f.cols.forEach(function (c) { if (!defaut && c.date === today) defaut = c.label; });
  f.cols.forEach(function (c) {
    if (defaut) return;
    if (lignes.some(function (r) { var x = idx[clePresence_(p.source, p.cle, c.label, r.nom, r.prenom)]; return x && x.date === todayFull && x.par.indexOf('fiche') !== 0; })) defaut = c.label;
  });
  if (!defaut) f.cols.forEach(function (c) { if (!defaut && !eleves.some(function (e) { return e.marks[c.label]; })) defaut = c.label; });
  if (!defaut && f.cols.length) defaut = f.cols[f.cols.length - 1].label;
  return {
    onglet: f.onglet, seances: f.cols.map(function (c) { return { label: c.label, date: c.date }; }), defaut: defaut, eleves: eleves, masques: masques,
    lien: 'https://docs.google.com/spreadsheets/d/' + f.ss.getId() + '/edit#gid=' + f.sheet.getSheetId()
  };
}

/** Enregistre l'appel d'une séance : p = { source, cle, seance, marks: [{nom, prenom, statut: P|A|E|''}] } */
function apiSaveAppel_(p) {
  if (!p.seance) throw new Error('Séance manquante.');
  var lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    var f = ficheTab_(p.source, p.cle);
    if (!f.cols.some(function (c) { return c.label === p.seance; })) throw new Error('Séance « ' + p.seance + ' » absente de la fiche.');
    // on remet d'abord la fiche d'aplomb (reprise des cases cochées à la main)
    syncFiche_(p.source, p.cle, f);
    ecrirePresences_((p.marks || []).map(function (m) {
      var st = String(m.statut || '');
      return { seance: p.seance, nom: m.nom, prenom: m.prenom, statut: ['P', 'A', 'E', 'S'].indexOf(st) >= 0 || st.indexOf('T:') === 0 ? st : '' };
    }), p.source, p.cle, f.onglet);
    syncFiche_(p.source, p.cle, f);
    logAction_('appel', f.onglet + ' ' + p.seance);
    return true;
  } finally { lock.releaseLock(); }
}

/** Tous les matins : remet chaque fiche en face des bons noms (après ajouts / départs / changements de cours) */
function syncToutesFiches_() {
  var n = 0;
  coursList_().forEach(function (k) { try { syncFiche_('cours', k.cle); n++; } catch (e) { } });
  stagesList_().forEach(function (s) { try { syncFiche_('stages', s.id); n++; } catch (e) { } });
  return n;
}
