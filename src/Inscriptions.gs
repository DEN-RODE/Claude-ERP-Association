/**
 * Inscriptions.gs : lecture / écriture directe des Sheets réponses des Forms
 * (cours annuels + stages), suivi admin (paiement, statut) dans l'onglet Suivi.
 *
 * Les colonnes sont retrouvées par leur INTITULÉ, pas par leur position :
 * si le Form change d'ordre ou gagne une question, rien ne casse.
 */

var FIELD_PATTERNS = [
  // clé interne, regex sur l'en-tête normalisé (norm_())
  ['ts', /^(timestamp|horodateur|horodatage|date de soumission)/],
  ['email', /^(email address|adresse e-?mail|adresse email|e-?mail address)$/],
  ['emailContact', /^(e-?mail|email de contact|e-?mail de contact)$/],
  ['nom', /^nom( de l'enfant| du pratiquant| de famille)?$/],
  ['prenom', /^prenom/],
  ['naissance', /date de naissance/],
  ['age', /^age/],
  ['creneaux', /creneau\(x\)|creneaux souhaites/],
  ['stage', /stage\(s\) souhaite|stages? souhaite/],
  ['creneau', /^creneau$/],
  ['niveau', /^niveau/],
  ['essai', /cours d'essai/],
  ['responsable', /responsable legal/],
  ['lien', /^lien avec/],
  ['adresse', /adresse postale/],
  ['tel', /^telephone/],
  ['second', /second contact/],
  ['autorises', /personnes autorisees/],
  ['repartSeul', /repartir seul/],
  ['allergies', /^allergies/],
  ['sante', /sante/],
  ['vaccins', /vaccination/],
  ['urgence', /personne a prevenir/],
  ['image', /droit a l'image/],
  ['soins', /autorisation de soins/],
  ['assurance', /assurance/],
  ['adhesion', /adhesion/],
  ['reductions', /reduction/],
  ['modePaiement', /mode de paiement/],
  ['commentaire', /commentaire/],
  ['certification', /certification/]
];

var FIELD_LABELS = {
  nom: 'Nom', prenom: 'Prénom', naissance: 'Date de naissance', creneaux: 'Créneau(x)', stage: 'Stage',
  niveau: 'Niveau', essai: "Cours d'essai", responsable: 'Responsable légal', lien: 'Lien', adresse: 'Adresse',
  tel: 'Téléphone', email: 'Email', emailContact: 'Email de contact', second: 'Second contact',
  autorises: 'Personnes autorisées', repartSeul: 'Repart seul', allergies: 'Allergies / régime',
  sante: 'Santé / traitement / PAI', vaccins: 'Vaccins à jour', urgence: 'Personne à prévenir',
  image: "Droit à l'image", soins: 'Autorisation de soins', assurance: 'Assurance RC',
  adhesion: 'Adhésion', reductions: 'Réductions', modePaiement: 'Mode de paiement', commentaire: 'Commentaire',
  certification: 'Certification'
};

/** valeurs considérées comme "rien à signaler" dans les champs santé */
var RAS = ['', 'non', 'ras', 'r.a.s', 'r.a.s.', 'na', 'n/a', 'nan', 'aucun', 'aucune', 'rien', 'neant', '-', '/', 'no', 'none', 'pas de', 'x', '0', 'non merci', 'rien a signaler', 'aucune allergie', 'pas d\'allergie', 'pas d\'allergies', 'aucun probleme', 'non aucun', 'non aucune'];

function isRas_(v) {
  var n = norm_(v).replace(/[.!]+$/, '');
  return RAS.indexOf(n) >= 0;
}

/** Associe chaque clé interne à un index de colonne */
function mapHeaders_(headers) {
  var map = {};
  var used = {};
  var nh = headers.map(norm_);
  FIELD_PATTERNS.forEach(function (fp) {
    for (var j = 0; j < nh.length; j++) {
      if (used[j] || !nh[j]) continue;
      if (fp[1].test(nh[j])) { map[fp[0]] = j; used[j] = true; break; }
    }
  });
  return map;
}

/** Ouvre la feuille de réponses (source = 'cours' | 'stages') */
function responseSheet_(source) {
  var c = getConfig_();
  var id = source === 'cours' ? c.ID_COURS_REPONSES : c.ID_STAGES_REPONSES;
  var tab = source === 'cours' ? c.ONGLET_COURS_REPONSES : c.ONGLET_STAGES_REPONSES;
  if (!id) throw new Error('ID du Sheet réponses ' + source + ' manquant dans Config.');
  var ss = SpreadsheetApp.openById(id);
  var sh = tab ? ss.getSheetByName(tab) : ss.getSheets()[0];
  if (!sh) throw new Error('Onglet "' + tab + '" introuvable dans le Sheet réponses ' + source);
  return sh;
}

/** Ligne d'en-tête = 1re ligne (parmi les 5 premières) qui contient Timestamp/Horodateur */
function readResponses_(source) {
  var sh = responseSheet_(source);
  var range = sh.getDataRange();
  var values = range.getValues();
  var bgs = range.getBackgrounds();
  var hr = 0;
  for (var i = 0; i < Math.min(5, values.length); i++) {
    var m = mapHeaders_(values[i].map(String));
    if (m.ts !== undefined && m.nom !== undefined) { hr = i; break; }
  }
  var headers = (values[hr] || []).map(String);
  var map = mapHeaders_(headers);
  var rows = [];
  for (var r = hr + 1; r < values.length; r++) {
    var line = values[r];
    if (line.join('').trim() === '') continue;
    var o = { _row: r + 1 };
    Object.keys(map).forEach(function (k) { o[k] = line[map[k]]; });
    // code couleur de l'équipe : case Nom en VERT = payé (rouge = signalé)
    o._bgNom = map.nom !== undefined ? bgs[r][map.nom] : '';
    o._bgPrenom = map.prenom !== undefined ? bgs[r][map.prenom] : '';
    rows.push(o);
  }
  return { sheet: sh, headers: headers, map: map, headerRow: hr + 1, rows: rows };
}

/** #6aa84f, #93c47d, #b7e1cd... = vert ; #f8f9fa (bandes alternées) = neutre */
function couleur_(hex) {
  var m = String(hex || '').match(/^#?([0-9a-f]{2})([0-9a-f]{2})([0-9a-f]{2})$/i);
  if (!m) return '';
  var r = parseInt(m[1], 16), g = parseInt(m[2], 16), b = parseInt(m[3], 16);
  if (g > r + 25 && g > b + 15) return 'vert';
  if (r > g + 60 && r > b + 60) return 'rouge';
  if (r > 200 && g > 180 && b < 160) return 'jaune';
  return '';
}

var COULEUR_PAYE = '#6aa84f', COULEUR_PARTIEL = '#ffe599';

function makeId_(ts, nom, prenom) {
  var t = isDate_(ts) ? String(Math.floor(ts.getTime() / 1000)) : norm_(ts);
  return t + '|' + norm_(nom) + '|' + norm_(prenom);
}

function suiviIndex_() {
  var idx = {};
  readTable_('Suivi').forEach(function (s) { idx[s.ID] = s; });
  return idx;
}

/** Transforme une ligne brute en objet élève sérialisable */
function toEleve_(source, o, suivi) {
  var id = makeId_(o.ts, o.nom, o.prenom);
  var s = suivi[id] || {};
  var e = { id: id, source: source, row: o._row, hasSuivi: !!suivi[id] };
  Object.keys(FIELD_LABELS).forEach(function (k) { e[k] = plain_(o[k]); });
  e.ts = fmtDateTime_(o.ts);
  e.tsRaw = isDate_(o.ts) ? o.ts.getTime() : null;
  e.nom = String(o.nom || '').trim();
  e.prenom = String(o.prenom || '').trim();
  e.age = ageAt_(o.naissance);
  e.emailBest = String(o.emailContact || o.email || '').trim();
  e.alerteSante = !isRas_(o.allergies) || !isRas_(o.sante);
  var manquants = [];
  if (!String(o.urgence || '').trim()) manquants.push('personne à prévenir');
  if (!String(o.soins || '').trim()) manquants.push('autorisation de soins');
  if (!String(o.tel || '').trim()) manquants.push('téléphone');
  e.manquants = manquants;
  // le Sheet fait foi : case Nom verte = payé
  var cNom = couleur_(o._bgNom);
  e.paye = cNom === 'vert' ? 'Oui' : (cNom === 'rouge' ? (s.Paye || 'Non') : (s.Paye === 'Partiel' || cNom === 'jaune' ? 'Partiel' : 'Non'));
  e.payeSource = cNom === 'vert' ? 'Sheet (case verte)' : (s.Paye ? 'ERP' : '');
  e.signale = cNom === 'rouge' || couleur_(o._bgPrenom) === 'rouge';
  e.montantRecu = s.MontantRecu || '';
  e.modeRecu = s.ModeRecu || '';
  e.statut = s.Statut || 'Actif';
  e.notes = s.Notes || '';
  e.derniereRelance = s.DerniereRelance || '';
  e.confirmation = s.Confirmation || '';
  return e;
}

var _elevesMemo = {};
function clearMemo_() { _elevesMemo = {}; invalideEleves_(); }

/**
 * Cache partagé entre les appels (5 min) : lire les Sheets réponses coûte 1 à 3 s par appel Apps Script.
 * Toute écriture passe par clearMemo_ / api() et change la "version", donc jamais de donnée périmée
 * après une action dans l'ERP. Seules les modifs faites à la main dans le Sheet (case verte...)
 * mettent jusqu'à 5 min à apparaître, ou tout de suite avec le bouton Actualiser.
 */
var ELEVES_TTL = 300;
function elevesVersion_() {
  try { return CacheService.getScriptCache().get('eleves_v') || '0'; } catch (e) { return 'x'; }
}
function invalideEleves_() {
  try { _exclusAppel = {}; } catch (e) { }
  try { CacheService.getScriptCache().put('eleves_v', String(Date.now()) + Math.floor(Math.random() * 1e6), 21600); } catch (e) { }
}

function loadEleves_(source) {
  if (_elevesMemo[source]) return _elevesMemo[source].map(function (e) { return JSON.parse(JSON.stringify(e)); });
  var ck = 'eleves_' + source + '_' + elevesVersion_();
  var hit = null;
  try { hit = cacheGetBig_(ck); } catch (e) { hit = null; }
  if (hit) {
    try { _elevesMemo[source] = JSON.parse(hit); return _elevesMemo[source].map(function (e) { return JSON.parse(JSON.stringify(e)); }); } catch (e) { }
  }
  var data = readResponses_(source);
  var suivi = suiviIndex_();
  migrerOrphelins_(source, data.rows, suivi);
  var list = data.rows.map(function (o) { return toEleve_(source, o, suivi); });
  _elevesMemo[source] = list;
  try { cachePutBig_(ck, JSON.stringify(list), ELEVES_TTL); } catch (e) { }
  return list.map(function (e) { return JSON.parse(JSON.stringify(e)); });
}

/**
 * Quand un parent modifie sa réponse au Form, Google change l'horodatage : l'ID change.
 * On rattache alors la ligne de suivi "orpheline" (même source + même nom/prénom,
 * plus aucune réponse avec son ID) à la nouvelle réponse, pour ne perdre ni paiement ni notes.
 */
function migrerOrphelins_(source, rows, suivi) {
  var ids = {};
  rows.forEach(function (o) { ids[makeId_(o.ts, o.nom, o.prenom)] = o; });
  var orphans = {};
  Object.keys(suivi).forEach(function (id) {
    var s = suivi[id];
    if (s.Source !== source || ids[id]) return;
    var k = id.split('|').slice(1).join('|');
    (orphans[k] = orphans[k] || []).push(s);
  });
  Object.keys(ids).forEach(function (id) {
    if (suivi[id]) return;
    var k = id.split('|').slice(1).join('|');
    var cand = orphans[k];
    var sameName = Object.keys(ids).filter(function (x) { return x.split('|').slice(1).join('|') === k && !suivi[x]; });
    if (cand && cand.length === 1 && sameName.length === 1) {
      var s = cand[0];
      try {
        var cell = erpSheet_('Suivi').getRange(s._row, headersOf_('Suivi').indexOf('ID') + 1);
        if (String(cell.getValue()) !== String(s.ID)) return; // la ligne a bougé entre-temps : on réessaiera au prochain chargement
        cell.setValue(safeCell_(id));
        delete suivi[s.ID]; s.ID = id; suivi[id] = s;
        logAction_('suivi rattaché', k + ' (réponse modifiée)');
      } catch (e) { }
    }
  });
}

// ---------- rattachement aux cours / stages ----------

function coursList_() {
  return readTable_('Cours').map(function (c) {
    return { cle: String(c.Cle), libelle: c.Libelle, jour: c.Jour, horaire: c.Horaire, public: c.Public, tarif: num_(c.Tarif), max: num_(c.Max) || 12,
      code: String(c.Code || '').trim().toUpperCase() || ('C' + c._row), row: c._row };
  });
}

/** Heure de début d'un horaire : « 17h00-18h30 » -> « 17h00 ». Rend '' si illisible. */
function heureDebut_(horaire) {
  var m = /(\d{1,2})\s*[h:]\s*(\d{0,2})/.exec(String(horaire || ''));
  if (!m) return '';
  return (m[1].length === 1 ? '0' + m[1] : m[1]) + 'h' + (m[2] ? (m[2].length === 1 ? m[2] + '0' : m[2]) : '00');
}

/** Toutes les heures citées dans un texte, au format 17h00 */
function heuresTexte_(t) {
  var out = [], re = /(\d{1,2})\s*[h:]\s*(\d{0,2})/g, m;
  while ((m = re.exec(String(t || '')))) {
    out.push((m[1].length === 1 ? '0' + m[1] : m[1]) + 'h' + (m[2] ? (m[2].length === 1 ? m[2] + '0' : m[2]) : '00'));
  }
  return out;
}

/** Le libellé à coller dans le Form : ce qu'on veut, suivi du code entre crochets */
function libelleCodeForm_(c, restantes) {
  var l = c.libelle + ' - ' + c.jour + ' ' + c.horaire;
  if (restantes !== undefined && restantes !== null) l += ' - ' + (restantes > 0 ? restantes + (restantes > 1 ? ' places restantes' : ' place restante') : 'complet');
  return l + ' [' + c.code + ']';
}

function stagesList_() {
  return readTable_('Stages').map(function (s) {
    return { cle: String(s.Cle), groupe: norm_(s.Groupe), nom: s.Nom, dates: s.Dates, public: s.Public, tarif: num_(s.Tarif), max: num_(s.Max) || 12, debut: s.Debut,
      id: norm_(s.Cle) + '#' + norm_(s.Groupe) };
  });
}

/**
 * Rattache un créneau coché dans le Form à un ou plusieurs cours. Trois niveaux, dans l'ordre :
 * 1. le code court entre crochets, ex. [C3] : insensible à tout ce qu'on écrit autour
 * 2. la colonne Cle cherchée telle quelle dans le texte (comportement historique)
 * 3. repli : le jour et l'heure de début retrouvés dans le texte
 * Le niveau 3 évite qu'un libellé retouché dans le Form (« 1 PLACE RESTANTE ... ») fasse
 * disparaître l'élève des listes d'appel.
 */
function matchCours_(creneauxText, cours) {
  var brut = String(creneauxText || '');
  if (!brut.trim()) return [];
  var t = ' ' + norm_(brut) + ' ';

  var codes = (brut.toUpperCase().match(/\[[A-Z0-9_-]{1,10}\]/g) || []).map(function (c) { return c.slice(1, -1); });
  if (codes.length) {
    var parCode = cours.filter(function (c) { return codes.indexOf(String(c.code).toUpperCase()) >= 0; });
    if (parCode.length) return parCode;
  }

  var parCle = cours.filter(function (c) { return norm_(c.cle) && t.indexOf(norm_(c.cle)) >= 0; });
  if (parCle.length) return parCle;

  var heures = heuresTexte_(brut);
  return cours.filter(function (c) {
    var j = norm_(c.jour), h = heureDebut_(c.horaire);
    return j && h && t.indexOf(j) >= 0 && heures.indexOf(h) >= 0;
  });
}

function matchStages_(stageText, stages) {
  // un parent peut cocher plusieurs stages : on découpe sur les virgules qui suivent un "€"
  var t = norm_(stageText);
  var parts = t.split(/€\s*,|\n/);
  var out = [];
  stages.forEach(function (s) {
    var key = ' ' + norm_(s.cle);
    for (var i = 0; i < parts.length; i++) {
      var p = ' ' + parts[i] + ' ';
      if (p.indexOf(key) >= 0 && p.indexOf(s.groupe) >= 0) { out.push(s); break; }
    }
  });
  return out;
}

// ---------- API ----------

/** Onglet(s) "desistement" du Sheet réponses : liste en lecture seule */
function desistements_(source) {
  try {
    var c = getConfig_();
    var ss = SpreadsheetApp.openById(source === 'cours' ? c.ID_COURS_REPONSES : c.ID_STAGES_REPONSES);
    var out = [];
    ss.getSheets().forEach(function (sh) {
      if (norm_(sh.getName()).indexOf('desist') < 0) return;
      var v = sh.getDataRange().getValues();
      var hr = -1, map = null;
      for (var i = 0; i < Math.min(5, v.length); i++) { var m = mapHeaders_(v[i].map(String)); if (m.nom !== undefined) { hr = i; map = m; break; } }
      if (!map) { map = mapHeaders_(v[0] ? v[0].map(String) : []); }
      // onglet sans en-tête : on suppose l'ordre du Form (Nom en C, Prénom en D, créneau en F)
      if (map.nom === undefined) map = { nom: 2, prenom: 3, creneaux: 5 };
      for (var r = hr + 1; r < v.length; r++) {
        var nom = String(v[r][map.nom] || '').trim(), prenom = String(v[r][map.prenom] || '').trim();
        if (!nom && !prenom) continue;
        out.push({ nom: nom, prenom: prenom, creneaux: plain_(v[r][map.creneaux !== undefined ? map.creneaux : map.stage]) , onglet: sh.getName() });
      }
    });
    return out;
  } catch (e) { return []; }
}

/**
 * Rattache une inscription non classée au cours choisi. On réécrit le créneau dans le Sheet
 * réponses avec le libellé conseillé (code entre crochets), donc le rattachement tient même
 * si quelqu'un retouche ensuite les libellés du formulaire.
 */
function apiRattacherCours_(p) {
  var k = coursList_().filter(function (c) { return String(c.cle) === String(p.cle); })[0];
  if (!k) throw new Error('Créneau inconnu : ' + p.cle);
  var r = apiModifierInscription_({ source: 'cours', id: p.id, fields: { creneaux: libelleCodeForm_(k) } });
  logAction_('rattachement', p.id + ' -> ' + k.cle);
  return { id: r.id, creneau: libelleCodeForm_(k), cle: k.cle };
}

/** Vue d'ensemble des cours avec élèves rattachés */
function apiCours_() {
  var _att = attenteParGroupe_('cours');
  var cours = coursList_();
  var eleves = loadEleves_('cours');
  var c = getConfig_();
  var min = num_(c.SEUIL_MIN_COURS) || 4;
  var byKey = {};
  cours.forEach(function (k) { byKey[k.cle] = []; });
  var nonClasses = [];
  eleves.forEach(function (e) {
    if (e.statut === 'Désinscrit') return;
    var m = matchCours_(e.creneaux, cours);
    e.coursCles = m.map(function (x) { return x.cle; });
    if (!m.length) nonClasses.push(e);
    m.forEach(function (k) { byKey[k.cle].push(e); });
  });
  var list = cours.map(function (k) {
    var el = byKey[k.cle].sort(function (a, b) { return a.nom.localeCompare(b.nom); });
    var n = el.length;
    return {
      cle: k.cle, libelle: k.libelle, jour: k.jour, horaire: k.horaire, public: k.public, tarif: k.tarif, max: k.max,
      code: k.code, libelleForm: libelleCodeForm_(k, Math.max(k.max - n, 0)),
      inscrits: n, libres: Math.max(k.max - n, 0),
      statut: n >= k.max ? 'Complet' : (n < min ? 'Sous le minimum' : 'Ouvert'),
      alertes: el.filter(function (e) { return e.alerteSante; }).length,
      impayes: el.filter(function (e) { return e.paye !== 'Oui'; }).length,
      attente: _att[k.cle] || 0,
      eleves: el
    };
  });
  return {
    cours: list, nonClasses: nonClasses,
    desinscrits: eleves.filter(function (e) { return e.statut === 'Désinscrit'; }),
    desistements: desistements_('cours'),
    payes: eleves.filter(function (e) { return e.statut !== 'Désinscrit' && e.paye === 'Oui'; }).length,
    total: eleves.filter(function (e) { return e.statut !== 'Désinscrit'; }).length,
    fiches: 'https://docs.google.com/spreadsheets/d/' + c.ID_FICHES_COURS,
    sheet: 'https://docs.google.com/spreadsheets/d/' + c.ID_COURS_REPONSES
  };
}

function apiStages_() {
  var _att = attenteParGroupe_('stages');
  var stages = stagesList_();
  var eleves = loadEleves_('stages');
  var c = getConfig_();
  var min = num_(c.SEUIL_MIN_STAGE) || 6;
  var byId = {};
  stages.forEach(function (s) { byId[s.id] = []; });
  var nonClasses = [];
  eleves.forEach(function (e) {
    if (e.statut === 'Désinscrit') return;
    var m = matchStages_(e.stage, stages);
    if (!m.length) nonClasses.push(e);
    m.forEach(function (s) { byId[s.id].push(e); });
  });
  var list = stages.map(function (s) {
    var el = byId[s.id].sort(function (a, b) { return a.nom.localeCompare(b.nom); });
    var n = el.length;
    return {
      id: s.id, nom: s.nom, dates: s.dates, groupe: s.groupe, public: s.public, tarif: s.tarif, max: s.max, debut: s.debut,
      inscrits: n, libres: Math.max(s.max - n, 0),
      statut: n === 0 ? 'Aucun inscrit' : (n >= s.max ? 'Complet' : (n < min ? 'Sous le seuil' : 'Confirmé')),
      alertes: el.filter(function (e) { return e.alerteSante; }).length,
      impayes: el.filter(function (e) { return e.paye !== 'Oui'; }).length,
      attente: _att[s.id] || 0,
      eleves: el
    };
  });
  return {
    stages: list, nonClasses: nonClasses,
    desinscrits: eleves.filter(function (e) { return e.statut === 'Désinscrit'; }),
    desistements: desistements_('stages'),
    payes: eleves.filter(function (e) { return e.statut !== 'Désinscrit' && e.paye === 'Oui'; }).length,
    total: eleves.filter(function (e) { return e.statut !== 'Désinscrit'; }).length,
    fiches: 'https://docs.google.com/spreadsheets/d/' + c.ID_FICHES_STAGES,
    sheet: 'https://docs.google.com/spreadsheets/d/' + c.ID_STAGES_REPONSES
  };
}

/** Recherche tous élèves (cours + stages) : nom, prénom, responsable, tel, email */
function apiRecherche_(p) {
  var q = norm_(p && p.q);
  var all = loadEleves_('cours').concat(loadEleves_('stages'));
  if (!q) return all.filter(function (e) { return e.alerteSante; }).slice(0, 200);
  return all.filter(function (e) {
    var hay = norm_([e.nom, e.prenom, e.responsable, e.tel, e.emailBest, e.creneaux, e.stage].join(' '));
    return q.split(' ').every(function (w) { return hay.indexOf(w) >= 0; });
  }).slice(0, 100);
}

/** Infos urgentes de TOUS les enfants (filtrées dans le navigateur : recherche instantanée) */
function apiUrgencesTous_() {
  var champs = ['id', 'source', 'nom', 'prenom', 'age', 'creneaux', 'stage', 'alerteSante', 'allergies', 'sante', 'urgence',
    'responsable', 'second', 'repartSeul', 'autorises', 'soins', 'tel', 'emailBest', 'statut'];
  return loadEleves_('cours').concat(loadEleves_('stages')).map(function (e) {
    var o = {}; champs.forEach(function (k) { if (e[k] !== undefined && e[k] !== null && e[k] !== '') o[k] = e[k]; });
    if (o.age === undefined) o.age = '';
    return o;
  }).sort(function (a, b) {
    return (b.alerteSante ? 1 : 0) - (a.alerteSante ? 1 : 0) || String(a.prenom).localeCompare(String(b.prenom));
  });
}

function apiEleve_(p) {
  var list = loadEleves_(p.source);
  for (var i = 0; i < list.length; i++) if (list[i].id === p.id) {
    var e = list[i];
    e.groupes = groupesDe_(p.source, e);
    return e;
  }
  throw new Error('Inscription introuvable (elle a peut-être été modifiée dans le Sheet).');
}

// ---------- liste d'attente (cours ou stage complet) ----------
// Les personnes en attente NE SONT PAS des élèves : elles vivent dans l'onglet Attente de l'ERP,
// jamais dans les Sheets réponses, et ne sont comptées nulle part dans les effectifs.

function attenteSheet_() {
  var ss = erpSS_();
  var sh = ss.getSheetByName('Attente');
  if (!sh) { sh = ss.insertSheet('Attente'); sh.getRange(1, 1, 1, TABLES.Attente.length).setValues([TABLES.Attente]); }
  return sh;
}

function attenteListe_() {
  attenteSheet_();
  return readTable_('Attente').map(function (r) {
    r.actif = norm_(r.Statut) !== 'inscrit' && norm_(r.Statut) !== 'abandon';
    return r;
  });
}

/** p = { source?, cle? } : tout, ou seulement un créneau */
function apiAttente_(p) {
  p = p || {};
  var l = attenteListe_();
  if (p.source) l = l.filter(function (r) { return String(r.Source) === String(p.source); });
  if (p.cle) l = l.filter(function (r) { return String(r.CleGroupe) === String(p.cle); });
  return l.sort(function (a, b) { return String(a.Date).localeCompare(String(b.Date)); });
}

/** Compte par créneau, pour les pastilles "X en attente" */
function attenteParGroupe_(source) {
  var out = {};
  attenteListe_().forEach(function (r) {
    if (!r.actif || String(r.Source) !== source) return;
    out[String(r.CleGroupe)] = (out[String(r.CleGroupe)] || 0) + 1;
  });
  return out;
}

function apiSaveAttente_(p) {
  attenteSheet_();
  if (!String(p.Nom || '').trim() && !String(p.Prenom || '').trim()) throw new Error('Nom ou prénom obligatoire.');
  if (!p.ID && !p.CleGroupe) throw new Error('Choisis le cours ou le stage.');
  var rec = {
    ID: p.ID || uid_(), Date: p.Date ? parseDate_(p.Date) || new Date() : new Date(),
    Source: p.Source === 'stages' ? 'stages' : 'cours', CleGroupe: p.CleGroupe,
    Nom: p.Nom, Prenom: p.Prenom, Naissance: p.Naissance || '', Responsable: p.Responsable || '',
    Tel: p.Tel || '', Email: p.Email || '', Notes: p.Notes || '', Statut: p.Statut || 'En attente',
    MajPar: currentUserLabel_(), MajLe: new Date()
  };
  upsertRow_('Attente', rec, 'ID');
  logAction_('liste attente', rec.Prenom + ' ' + rec.Nom + ' -> ' + rec.CleGroupe + ' (' + rec.Statut + ')');
  return rec;
}

function apiStatutAttente_(p) {
  var l = attenteListe_().filter(function (r) { return String(r.ID) === String(p.id); })[0];
  if (!l) throw new Error('Personne introuvable dans la liste d\'attente.');
  upsertRow_('Attente', { ID: p.id, Statut: p.statut, MajPar: currentUserLabel_(), MajLe: new Date() }, 'ID');
  logAction_('liste attente statut', l.Prenom + ' ' + l.Nom + ' -> ' + p.statut);
  return true;
}

function apiSupprimerAttente_(p) {
  attenteSheet_();
  if (!deleteRowById_('Attente', p.id)) throw new Error('Personne introuvable.');
  logAction_('liste attente suppression', String(p.id));
  return true;
}

/** Réglages d'un créneau (places max, tarif, horaire, libellé) : écrit dans l'onglet Cours / Stages de l'ERP */
function apiReglagesGroupe_(p) {
  var table = p.source === 'cours' ? 'Cours' : 'Stages';
  var rows = readTable_(table);
  var ligne = null;
  rows.forEach(function (r) {
    if (ligne) return;
    if (p.source === 'cours') { if (String(r.Cle) === String(p.cle)) ligne = r; }
    else if (norm_(r.Cle) + '#' + norm_(r.Groupe) === String(p.cle)) ligne = r;
  });
  if (!ligne) throw new Error('Créneau introuvable dans l\'onglet ' + table + '.');
  if (p.lire) return { Libelle: ligne.Libelle || ligne.Nom, Horaire: ligne.Horaire || ligne.Dates, Public: ligne.Public, Tarif: num_(ligne.Tarif), Max: num_(ligne.Max) };
  var sh = erpSheet_(table);
  var h = headersOf_(table);
  var champs = { Max: num_(p.Max), Tarif: num_(p.Tarif) };
  if (p.Public !== undefined) champs.Public = String(p.Public);
  if (p.source === 'cours' && p.Horaire !== undefined) champs.Horaire = String(p.Horaire);
  if (!champs.Max || champs.Max < 1 || champs.Max > 99) throw new Error('Places max : un nombre entre 1 et 99.');
  Object.keys(champs).forEach(function (k) {
    var i = h.indexOf(k);
    if (i >= 0 && champs[k] !== undefined && champs[k] !== '') sh.getRange(ligne._row, i + 1).setValue(safeCell_(champs[k]));
  });
  clearMemo_();
  logAction_('réglages ' + p.source, p.cle + ' max=' + champs.Max + ' tarif=' + champs.Tarif);
  return true;
}

// ---------- déplacer / retirer d'un cours ou d'un stage ----------

/** Libellé au format du Google Form, pour que les fiches de présence et le Modele_economique le reconnaissent */
/**
 * Intitulé à écrire dans la colonne Créneau(x) / Stage(s) : on reprend le texte EXACT du Form
 * (sinon les fiches de présence, qui cherchent l'intitulé exact, ne trouvent pas l'élève).
 * 1) choix du Google Form, 2) texte déjà présent dans les réponses, 3) intitulé reconstruit.
 */
function choixForm_(source) {
  var cache = CacheService.getScriptCache();
  var ck = 'choixForm_' + source;
  var hit = cache.get(ck);
  if (hit) { try { return JSON.parse(hit); } catch (e) { } }
  var out = [];
  var c = getConfig_();
  var idForm = source === 'cours' ? c.ID_FORM_COURS : c.ID_FORM_STAGES;
  if (idForm && typeof FormApp !== 'undefined') {
    try {
      FormApp.openById(idForm).getItems().forEach(function (it) {
        var t = norm_(it.getTitle());
        if (source === 'cours' ? t.indexOf('creneau') < 0 : t.indexOf('stage') < 0) return;
        var typ = String(it.getType());
        var x = typ === 'CHECKBOX' ? it.asCheckboxItem() : typ === 'MULTIPLE_CHOICE' ? it.asMultipleChoiceItem() : typ === 'LIST' ? it.asListItem() : null;
        if (x) x.getChoices().forEach(function (ch) { out.push(String(ch.getValue()).trim()); });
      });
    } catch (e) { /* Form inaccessible : on se rabat sur les réponses */ }
  }
  // textes déjà écrits par le Form dans le Sheet réponses (le plus fréquent d'abord)
  try {
    var vus = {};
    readResponses_(source).rows.forEach(function (o) {
      morceaux_(source, source === 'cours' ? o.creneaux : o.stage).forEach(function (m) { vus[m] = (vus[m] || 0) + 1; });
    });
    Object.keys(vus).sort(function (a, b) { return vus[b] - vus[a]; }).forEach(function (m) { if (out.indexOf(m) < 0) out.push(m); });
  } catch (e) { }
  try { cache.put(ck, JSON.stringify(out), 21600); } catch (e) { }
  return out;
}

function libelleForm_(source, g, choix) {
  choix = choix || choixForm_(source);
  for (var i = 0; i < choix.length; i++) if (morceauEst_(source, choix[i], g)) return choix[i];
  if (source === 'cours') return g.libelle + ' \u2014 ' + g.jour + ' ' + g.horaire + ' (' + g.tarif + ' \u20ac/an )';
  return g.dates + ' - ' + (g.groupe === 'matin' ? 'Matin' : 'Après-midi') + ' ' + String(g.public || '').replace(/^(matin|apr[eè]s-midi)\s*/i, '') + ' \u2014 ' + g.tarif + ' \u20ac';
}

function groupesListe_(source) {
  return source === 'cours'
    ? coursList_().map(function (k) { return { cle: k.cle, label: k.jour + ' ' + k.horaire + ' · ' + k.libelle, g: k }; })
    : stagesList_().map(function (s) { return { cle: s.id, label: s.nom + ' · ' + (s.groupe === 'matin' ? 'Matin' : 'Après-midi') + ' · ' + s.dates, g: s }; });
}

function groupesDe_(source, e) {
  var m = source === 'cours' ? matchCours_(e.creneaux, coursList_()) : matchStages_(e.stage, stagesList_());
  return m.map(function (g) { return source === 'cours' ? { cle: g.cle, label: g.jour + ' ' + g.horaire + ' · ' + g.libelle } : { cle: g.id, label: g.nom + ' · ' + (g.groupe === 'matin' ? 'Matin' : 'Après-midi') }; });
}

/** Liste des cours / stages possibles (pour le menu Déplacer) */
function apiGroupes_(p) {
  return groupesListe_(p.source).map(function (x) { return { cle: x.cle, label: x.label }; });
}

/** Découpe le texte "créneau(x)" du Form en morceaux (un par case cochée) */
function morceaux_(source, texte) {
  var t = String(texte || '');
  if (!t.trim()) return [];
  if (source === 'stages') return t.split(/(?<=\u20ac)\s*,\s*/).map(function (x) { return x.trim(); }).filter(String);
  // une case cochée du Form finit par "(xxx €/an )" : on coupe après
  return t.split(/(?<=\/an\s?\))\s*,\s*/).map(function (x) { return x.trim(); }).filter(String);
}

function morceauEst_(source, morceau, g) {
  var n = ' ' + norm_(morceau) + ' ';
  if (source === 'cours') return n.indexOf(norm_(g.cle)) >= 0;
  return n.indexOf(' ' + norm_(g.cle)) >= 0 && n.indexOf(g.groupe) >= 0;
}

/**
 * p = { source, id, de: cle (ou '' pour un simple ajout), vers: cle (ou '' pour retirer) }
 * Réécrit la colonne Créneau(x) / Stage(s) du Sheet réponses.
 */
function apiDeplacer_(p) {
  if (!p.de && !p.vers) throw new Error('Rien à faire.');
  var lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    var data = readResponses_(p.source);
    var col = p.source === 'cours' ? data.map.creneaux : data.map.stage;
    if (col === undefined) throw new Error('Colonne créneau / stage introuvable dans le Sheet réponses.');
    var row = findRowById_(data, p.id);
    if (row < 0) throw new Error('Inscription introuvable.');
    var liste = groupesListe_(p.source);
    var gDe = p.de ? liste.filter(function (x) { return x.cle === p.de; })[0] : null;
    var gVers = p.vers ? liste.filter(function (x) { return x.cle === p.vers; })[0] : null;
    if (p.de && !gDe) throw new Error('Cours / stage de départ inconnu.');
    if (p.vers && !gVers) throw new Error('Cours / stage d\'arrivée inconnu.');
    var cell = data.sheet.getRange(row, col + 1);
    var parts = morceaux_(p.source, cell.getValue());
    if (gDe) {
      var avant = parts.length;
      parts = parts.filter(function (m) { return !morceauEst_(p.source, m, gDe.g); });
      if (parts.length === avant) throw new Error('Cet élève n\'est pas inscrit dans « ' + gDe.label + ' ».');
    }
    if (gVers && !gDe && parts.some(function (m) { return morceauEst_(p.source, m, gVers.g); })) throw new Error('Déjà inscrit dans « ' + gVers.label + ' ».');
    if (gVers && !parts.some(function (m) { return morceauEst_(p.source, m, gVers.g); })) parts.push(libelleForm_(p.source, gVers.g));
    if (!parts.length) throw new Error('Ce serait son seul cours : utilise plutôt Désinscrire ou Supprimer.');
    cell.setValue(safeCell_(parts.join(', ')));
    clearMemo_();
    logAction_('déplacement ' + p.source, p.id + ' : ' + (gDe ? gDe.label : '-') + ' -> ' + (gVers ? gVers.label : 'retiré'));
    return { creneaux: parts.join(', ') };
  } finally { lock.releaseLock(); }
}

/** Liste des champs modifiables pour le formulaire d'ajout (selon les colonnes réelles) */
function apiChamps_(p) {
  var data = readResponses_(p.source);
  var keys = Object.keys(data.map).filter(function (k) { return k !== 'ts' && FIELD_LABELS[k]; });
  return {
    champs: keys.map(function (k) {
      var c = { key: k, label: FIELD_LABELS[k], header: data.headers[data.map[k]] };
      // réponses types (Oui / Non, J'autorise...) reprises du Sheet pour les proposer en liste
      if (['essai', 'niveau', 'repartSeul', 'image', 'soins', 'vaccins', 'assurance', 'adhesion', 'modePaiement', 'lien', 'reductions', 'certification'].indexOf(k) >= 0) {
        var vus = {};
        data.rows.forEach(function (o) { var v = String(o[k] || '').trim(); if (v && v.length < 160) vus[v] = (vus[v] || 0) + 1; });
        c.choix = Object.keys(vus).sort(function (a, b) { return vus[b] - vus[a]; }).slice(0, 8);
      }
      return c;
    }),
    // menu déroulant : libellé court + texte exact du Form écrit dans le Sheet
    groupes: (function () { var ch = choixForm_(p.source); return groupesListe_(p.source).map(function (x) { return { cle: x.cle, label: x.label, form: libelleForm_(p.source, x.g, ch) }; }); })()
  };
}

function findRowById_(data, id) {
  for (var i = 0; i < data.rows.length; i++) {
    var o = data.rows[i];
    if (makeId_(o.ts, o.nom, o.prenom) === id) return o._row;
  }
  return -1;
}

/** Ajout d'une inscription directement dans le Sheet réponses */
function apiAjouterInscription_(p) {
  var lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    var data = readResponses_(p.source);
    if (!p.fields || !String(p.fields.nom || '').trim() || !String(p.fields.prenom || '').trim()) throw new Error('Nom et prénom obligatoires.');
    var line = data.headers.map(function () { return ''; });
    var now = new Date(Math.floor(Date.now() / 1000) * 1000);
    line[data.map.ts] = now;
    Object.keys(p.fields).forEach(function (k) {
      if (data.map[k] === undefined || k === 'ts') return;
      var v = p.fields[k];
      if (k === 'naissance' && v) { var d = parseDate_(v); if (d) v = d; }
      line[data.map[k]] = safeCell_(v);
    });
    data.sheet.appendRow(line);
    clearMemo_();
    var id = makeId_(now, p.fields.nom, p.fields.prenom);
    saveSuivi_({ ID: id, Source: p.source, Nom: p.fields.nom, Prenom: p.fields.prenom, Statut: 'Actif', Paye: 'Non', Notes: 'Ajouté depuis l\'ERP' });
    logAction_('ajout ' + p.source, p.fields.prenom + ' ' + p.fields.nom);
    return { id: id };
  } finally { lock.releaseLock(); }
}

/** Modification des champs d'une inscription (dans le Sheet réponses) */
function apiModifierInscription_(p) {
  var lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    var data = readResponses_(p.source);
    var row = findRowById_(data, p.id);
    if (row < 0) throw new Error('Inscription introuvable.');
    var changes = [];
    Object.keys(p.fields || {}).forEach(function (k) {
      if (data.map[k] === undefined || k === 'ts') return;
      var v = p.fields[k];
      if (k === 'naissance' && v) { var d = parseDate_(v); if (d) v = d; }
      data.sheet.getRange(row, data.map[k] + 1).setValue(safeCell_(v));
      changes.push(k);
    });
    clearMemo_();
    var newId = p.id;
    if (p.fields && (p.fields.nom !== undefined || p.fields.prenom !== undefined)) {
      var o = data.rows.filter(function (r) { return r._row === row; })[0];
      newId = makeId_(o.ts, p.fields.nom !== undefined ? p.fields.nom : o.nom, p.fields.prenom !== undefined ? p.fields.prenom : o.prenom);
      if (newId !== p.id) {
        var s = suiviIndex_()[p.id];
        if (s) { deleteRowById_('Suivi', p.id); s.ID = newId; s.Nom = p.fields.nom || s.Nom; s.Prenom = p.fields.prenom || s.Prenom; delete s._row; saveSuivi_(s); }
      }
    }
    logAction_('modif ' + p.source, p.id + ' : ' + changes.join(', '));
    return { id: newId };
  } finally { lock.releaseLock(); }
}

/** Paiement / statut / notes (onglet Suivi de l'ERP, le Sheet réponses n'est pas touché) */
function apiSuivi_(p) {
  var s = suiviIndex_()[p.id] || { ID: p.id, Source: p.source, Nom: p.nom || '', Prenom: p.prenom || '' };
  delete s._row;
  ['Paye', 'MontantRecu', 'ModeRecu', 'Statut', 'Notes'].forEach(function (k) { if (p[k] !== undefined) s[k] = p[k]; });
  if (p.Paye !== undefined) colorierPaiement_(p.source, p.id, p.Paye);
  saveSuivi_(s);
  logAction_('suivi', p.id + ' ' + JSON.stringify(p).slice(0, 200));
  return true;
}

/** Répercute le paiement dans le Sheet réponses : case Nom verte (payé), jaune (partiel), sinon on enlève la couleur */
function colorierPaiement_(source, id, paye) {
  var data = readResponses_(source);
  var row = findRowById_(data, id);
  if (row < 0 || data.map.nom === undefined) return;
  var cell = data.sheet.getRange(row, data.map.nom + 1);
  var actuel = couleur_(cell.getBackground());
  if (actuel === 'rouge') return; // case signalée en rouge par l'équipe : on n'y touche pas (le suivi ERP garde le statut)
  if (paye === 'Oui') { if (actuel !== 'vert') cell.setBackground(COULEUR_PAYE); }
  else if (paye === 'Partiel') cell.setBackground(COULEUR_PARTIEL);
  else if (actuel === 'vert' || actuel === 'jaune') cell.setBackground(null);
}

function saveSuivi_(s) {
  clearMemo_();
  s.MajPar = currentUserLabel_();
  s.MajLe = new Date();
  upsertRow_('Suivi', s, 'ID');
}

/**
 * Suppression définitive : efface la ligne du Sheet réponses.
 * (La réponse reste dans l'onglet Réponses du Google Form lui-même.)
 */
function apiSupprimerInscription_(p) {
  if (p.confirm !== 'SUPPRIMER') throw new Error('Confirmation manquante.');
  var lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    var data = readResponses_(p.source);
    var row = findRowById_(data, p.id);
    if (row < 0) throw new Error('Inscription introuvable.');
    data.sheet.deleteRow(row);
    clearMemo_();
    deleteRowById_('Suivi', p.id);
    logAction_('suppression ' + p.source, p.id);
    return true;
  } finally { lock.releaseLock(); }
}

/**
 * Crée la ligne manquante dans l'onglet Suivi pour chaque élève, avec le paiement
 * repris de la couleur de sa case Nom. L'onglet devient lisible tel quel dans le classeur.
 * Ne touche jamais une ligne qui existe déjà.
 */
function completerSuivi_() {
  var idx = suiviIndex_();
  var ajoutes = 0;
  ['cours', 'stages'].forEach(function (source) {
    loadEleves_(source).forEach(function (e) {
      if (idx[e.id]) return;
      upsertRow_('Suivi', {
        ID: e.id, Source: source, Nom: e.nom, Prenom: e.prenom,
        Paye: e.paye, MontantRecu: '', ModeRecu: '', Statut: e.statut || 'Actif',
        Notes: '', DerniereRelance: '', Confirmation: '',
        MajPar: currentUserLabel_(), MajLe: new Date()
      }, 'ID');
      ajoutes++;
    });
  });
  if (ajoutes) clearMemo_();
  logAction_('completer suivi', ajoutes + ' ligne(s) ajoutee(s)');
  return ajoutes;
}

function completerSuivi() {
  proprietaireSeulement_();
  var n = completerSuivi_();
  try { SpreadsheetApp.getUi().alert(n ? n + ' ligne(s) ajoutee(s) dans l onglet Suivi.' : 'L onglet Suivi etait deja complet.'); } catch (e) { }
  return n;
}

/** Liste "urgences" d'un groupe : tous les élèves avec contacts d'urgence, alertes en premier */
function apiUrgencesGroupe_(p) {
  var groupe;
  if (p.source === 'cours') groupe = apiCours_().cours.filter(function (c) { return c.cle === p.cle; })[0];
  else groupe = apiStages_().stages.filter(function (s) { return s.id === p.cle; })[0];
  if (!groupe) throw new Error('Groupe introuvable');
  var el = groupe.eleves.slice().sort(function (a, b) { return (b.alerteSante ? 1 : 0) - (a.alerteSante ? 1 : 0) || a.nom.localeCompare(b.nom); });
  return { titre: (groupe.libelle || groupe.nom) + ' · ' + (groupe.jour ? groupe.jour + ' ' + groupe.horaire : groupe.dates + ' ' + (groupe.groupe === 'matin' ? 'matin' : 'après-midi')), eleves: el };
}

/** Pousse les nombres d'inscrits dans Modele_economique (colonne Inscrits), sur demande */
function apiSyncModeleEco_() {
  var c = getConfig_();
  var ss = SpreadsheetApp.openById(c.ID_MODELE_ECO);
  var cours = apiCours_().cours;
  var stages = apiStages_().stages;
  var done = [];
  ss.getSheets().forEach(function (sh) {
    var v = sh.getDataRange().getValues();
    for (var r = 0; r < v.length; r++) {
      var h = v[r].map(norm_);
      var iCours = h.indexOf('cours'), iInscrits = h.indexOf('inscrits');
      if (iCours >= 0 && iInscrits >= 0) {
        for (var k = r + 1; k < v.length && norm_(v[k][iCours]) !== 'total' && String(v[k][iCours]).trim(); k++) {
          var label = norm_(v[k][iCours]);
          var match = cours.filter(function (x) { return label.indexOf(norm_(x.cle)) === 0; })[0];
          if (match) { if (setIfNotFormula_(sh.getRange(k + 1, iInscrits + 1), match.inscrits)) done.push(match.cle + ' = ' + match.inscrits); }
        }
      }
      var iStage = h.indexOf('stage'), iMat = h.indexOf('inscrits matin'), iAp = h.indexOf('inscrits aprem');
      if (iStage >= 0 && iMat >= 0 && iAp >= 0) {
        for (var k2 = r + 1; k2 < v.length && norm_(v[k2][iStage]) !== 'total' && String(v[k2][iStage]).trim(); k2++) {
          var nomStage = norm_(v[k2][iStage]);
          var mat = stages.filter(function (s) { return norm_(s.nom) === nomStage && s.groupe === 'matin'; })[0];
          var ap = stages.filter(function (s) { return norm_(s.nom) === nomStage && s.groupe === 'apres-midi'; })[0];
          if (mat) setIfNotFormula_(sh.getRange(k2 + 1, iMat + 1), mat.inscrits);
          if (ap) setIfNotFormula_(sh.getRange(k2 + 1, iAp + 1), ap.inscrits);
          if (mat || ap) done.push(v[k2][iStage] + ' = ' + (mat ? mat.inscrits : '-') + '/' + (ap ? ap.inscrits : '-'));
        }
      }
    }
  });
  logAction_('sync modele eco', done.join(' ; '));
  return done;
}

function setIfNotFormula_(range, v) {
  if (range.getFormula()) return false;
  range.setValue(v);
  return true;
}
