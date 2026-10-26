/**
 * Divers.gs : accueil, rappels, notes site, agenda, médiathèque comm, docs, rapports Claude
 */

var JOURS = ['Dimanche', 'Lundi', 'Mardi', 'Mercredi', 'Jeudi', 'Vendredi', 'Samedi'];

function apiAccueil_() {
  var now = new Date();
  var jour = JOURS[now.getDay()];
  var cours = apiCours_();
  var stages = apiStages_();
  var coursDuJour = cours.cours.filter(function (k) { return norm_(k.jour) === norm_(jour); }).map(function (k) {
    return { cle: k.cle, libelle: k.libelle, horaire: k.horaire, inscrits: k.inscrits, alertes: k.alertes };
  });
  var today = Utilities.formatDate(now, TZ, 'yyyy-MM-dd');
  var rappels = readTable_('Rappels').filter(function (r) { return norm_(r.Fait) !== 'oui'; }).map(function (r) {
    var d = parseDate_(r.Date);
    var iso = d ? Utilities.formatDate(d, TZ, 'yyyy-MM-dd') : '';
    r.iso = iso; r.enRetard = iso && iso < today; r.aujourdhui = iso === today;
    return r;
  }).filter(function (r) {
    return r.enRetard || r.aujourdhui || (r.iso && r.iso <= Utilities.formatDate(new Date(now.getTime() + 3 * 86400000), TZ, 'yyyy-MM-dd'));
  }).sort(function (a, b) { return a.iso.localeCompare(b.iso); });

  // agenda et compteurs Gmail : lents (1-2 s chacun), gardés 5 min
  var cache = CacheService.getScriptCache();
  var agenda = null;
  try { agenda = JSON.parse(cache.get('accueil_agenda') || 'null'); } catch (e) { agenda = null; }
  if (!agenda) {
    try { agenda = agenda_(3); cache.put('accueil_agenda', JSON.stringify(agenda), 300); } catch (e) { agenda = []; }
  }

  var all = [];
  cours.cours.forEach(function (k) { all = all.concat(k.eleves); });
  all = all.concat(cours.nonClasses);
  var uniqCours = {}; all.forEach(function (e) { uniqCours[e.id] = e; });
  var elevesCours = Object.keys(uniqCours).map(function (k) { return uniqCours[k]; });
  var st = []; stages.stages.forEach(function (s) { st = st.concat(s.eleves); }); st = st.concat(stages.nonClasses);
  var uniqSt = {}; st.forEach(function (e) { uniqSt[e.id] = e; });
  var elevesStages = Object.keys(uniqSt).map(function (k) { return uniqSt[k]; });
  var tous = elevesCours.concat(elevesStages);

  var nonLus = 0, brouillons = 0;
  var gm = null;
  try { gm = JSON.parse(cache.get('accueil_gmail') || 'null'); } catch (e) { gm = null; }
  if (!gm) {
    try {
      gm = { nonLus: GmailApp.getInboxUnreadCount(), brouillons: GmailApp.getDrafts().length };
      cache.put('accueil_gmail', JSON.stringify(gm), 300);
    } catch (e) { gm = { nonLus: 0, brouillons: 0 }; }
  }
  nonLus = gm.nonLus; brouillons = gm.brouillons;
  var aRel = 0;
  try { aRel = aRelancer_().length; } catch (e) { }

  var prochainsStages = stages.stages.filter(function (s) {
    var d = parseDate_(s.debut); return d && d.getTime() >= now.getTime() - 7 * 86400000;
  }).slice(0, 4).map(function (s) { return { id: s.id, nom: s.nom, dates: s.dates, groupe: s.groupe, inscrits: s.inscrits, max: s.max, statut: s.statut }; });

  return {
    date: Utilities.formatDate(now, TZ, 'dd/MM/yyyy'), jour: jour,
    coursDuJour: coursDuJour, rappels: rappels, agenda: agenda, prochainsStages: prochainsStages,
    nouveaux: tous.filter(function (e) { return e.tsRaw; }).sort(function (a, b) { return b.tsRaw - a.tsRaw; }).slice(0, 6)
      .map(function (e) { return { id: e.id, source: e.source, nom: e.nom, prenom: e.prenom, ts: e.ts }; }),
    coursFragiles: cours.cours.filter(function (k) { return k.statut === 'Sous le minimum'; }).map(function (k) { return k.jour + ' ' + k.horaire + ' (' + k.inscrits + ')'; }),
    stats: {
      eleves: elevesCours.length, stagiaires: elevesStages.length,
      impayes: tous.filter(function (e) { return e.paye !== 'Oui'; }).length,
      alertes: tous.filter(function (e) { return e.alerteSante; }).length,
      nonClasses: cours.nonClasses.length + stages.nonClasses.length,
      nonLus: nonLus, brouillons: brouillons, aRelancer: aRel
    }
  };
}

// ---------- Rappels ----------
function apiRappels_() {
  return readTable_('Rappels').sort(function (a, b) {
    var da = parseDate_(a.Date), db = parseDate_(b.Date);
    return (da ? da.getTime() : 0) - (db ? db.getTime() : 0);
  });
}
function apiSaveRappel_(p) {
  if (!p.Titre) throw new Error('Titre obligatoire.');
  var r = { ID: p.ID || uid_(), Date: parseDate_(p.Date) || '', Heure: p.Heure || '', Type: p.Type || 'Autre', Titre: p.Titre, Detail: p.Detail || '', Fait: p.Fait || 'non' };
  upsertRow_('Rappels', r, 'ID');
  return r.ID;
}
function apiDeleteRappel_(p) { return deleteRowById_('Rappels', p.ID); }

// ---------- Notes site ----------
function apiNotesSite_() { return readTable_('NotesSite').reverse(); }
function apiSaveNoteSite_(p) {
  if (!p.Description) throw new Error('Description obligatoire.');
  var n = { ID: p.ID || uid_(), Date: p.ID ? (parseDate_(p.Date) || new Date()) : new Date(), Page: p.Page || '', Description: p.Description,
    Priorite: p.Priorite || 'Normale', Statut: p.Statut || 'À faire', Auteur: p.Auteur || '' };
  upsertRow_('NotesSite', n, 'ID');
  return n.ID;
}
function apiStatutNoteSite_(p) {
  var n = readTable_('NotesSite').filter(function (r) { return String(r.ID) === String(p.ID); })[0];
  if (!n) throw new Error('Note introuvable.');
  var statut = p.Statut || (String(n.Statut) === 'Fait' ? 'À faire' : 'Fait');
  upsertRow_('NotesSite', { ID: n.ID, Statut: statut }, 'ID');
  logAction_('note site', String(n.Description).slice(0, 60) + ' -> ' + statut);
  return statut;
}
function apiDeleteNoteSite_(p) { return deleteRowById_('NotesSite', p.ID); }

// ---------- Agenda (agenda Google du compte) ----------
/** Couleurs de l'agenda Google : numéro de couleur -> { categorie, hex } (réglable dans Config) */
function couleursAgenda_() {
  var out = {};
  String(getConfig_().COULEURS_AGENDA || '').split(',').forEach(function (x) {
    var m = x.trim().match(/^(\d+)\s*=\s*([^|]+)\|\s*(#[0-9a-f]{3,8})$/i);
    if (m) out[m[1]] = { categorie: m[2].trim(), hex: m[3] };
  });
  return out;
}

function agenda_(jours) {
  var cols = couleursAgenda_();
  var now = new Date();
  var start = new Date(now.getFullYear(), now.getMonth(), now.getDate());
  var fin = new Date(now.getFullYear(), now.getMonth(), now.getDate() + jours); // jours entiers : 3 = aujourd'hui, demain, après-demain
  return CalendarApp.getDefaultCalendar().getEvents(start, fin).slice(0, 60).map(function (ev) {
    var allDay = ev.isAllDayEvent();
    var d = ev.getStartTime();
    var cid = '';
    try { cid = String(ev.getColor() || ''); } catch (e) { cid = ''; }
    var c = cols[cid] || null;
    return {
      id: ev.getId(), titre: ev.getTitle(), lieu: ev.getLocation(), allDay: allDay,
      couleur: c ? c.hex : '', categorie: c ? c.categorie : '',
      iso: Utilities.formatDate(d, TZ, 'yyyy-MM-dd'),
      quand: JOURS[d.getDay()].slice(0, 3) + ' ' + Utilities.formatDate(d, TZ, allDay ? 'dd/MM' : 'dd/MM HH:mm'),
      fin: allDay ? '' : Utilities.formatDate(ev.getEndTime(), TZ, 'HH:mm')
    };
  });
}
function apiAgenda_(p) { return agenda_((p && p.jours) || 21); }
function apiCreerEvenement_(p) {
  if (!p.titre || !p.date) throw new Error('Titre et date obligatoires.');
  var d = parseDate_(p.date);
  var cal = CalendarApp.getDefaultCalendar();
  var ev;
  if (p.heure) {
    var hm = String(p.heure).split(':');
    var s = new Date(d.getFullYear(), d.getMonth(), d.getDate(), +hm[0], +(hm[1] || 0));
    var e = new Date(s.getTime() + (num_(p.duree) || 60) * 60000);
    ev = cal.createEvent(p.titre, s, e, { location: p.lieu || '', description: p.detail || '' });
  } else {
    ev = cal.createAllDayEvent(p.titre, d, { location: p.lieu || '', description: p.detail || '' });
  }
  logAction_('agenda', p.titre + ' ' + p.date);
  return ev.getId();
}

// ---------- Médiathèque (images + textes des posts) ----------
function apiMedias_() {
  var hit = cacheGetBig_('medias_v3');
  if (hit) return JSON.parse(hit);
  var cfg = getConfig_();
  // Comm = uniquement le dossier Photo & Affiche (et ses sous-dossiers)
  var ids = String(cfg.ID_DOSSIER_PHOTOS || cfg.IDS_DOSSIERS_MEDIAS || '').split(',').map(function (s) { return s.trim(); }).filter(String);
  var out = [];
  var seen = {};
  function walk(folder, path, depth) {
    if (depth > 4 || out.length > 600) return;
    var files = folder.getFiles();
    while (files.hasNext() && out.length <= 600) {
      var f = files.next();
      if (seen[f.getId()]) continue;
      seen[f.getId()] = true;
      var mime = f.getMimeType();
      var type = mime.indexOf('image/') === 0 ? 'image' : (mime === MimeType.GOOGLE_DOCS ? 'texte' : (mime.indexOf('video/') === 0 ? 'video' : (mime === MimeType.PDF ? 'pdf' : 'autre')));
      out.push({ id: f.getId(), nom: f.getName(), type: type, dossier: path, url: f.getUrl(), modif: f.getLastUpdated().getTime(), date: fmtDate_(f.getLastUpdated()) });
    }
    var subs = folder.getFolders();
    while (subs.hasNext()) { var s = subs.next(); walk(s, path + ' / ' + s.getName(), depth + 1); }
  }
  ids.forEach(function (id) {
    try { var f = DriveApp.getFolderById(id); walk(f, f.getName(), 0); } catch (e) { logAction_('médias', 'dossier inaccessible ' + id); }
  });
  out.sort(function (a, b) { return b.modif - a.modif; });
  cachePutBig_('medias_v3', JSON.stringify(out), 900);
  return out;
}
function apiRafraichirMedias_() { cacheRemoveBig_('medias_v3'); return apiMedias_(); }

/** Sous-dossiers de Photo & Affiche (pour choisir où ranger une image envoyée depuis l'ERP) */
function apiMediasDossiers_() {
  var c = getConfig_();
  var root = DriveApp.getFolderById(c.ID_DOSSIER_PHOTOS || String(c.IDS_DOSSIERS_MEDIAS || '').split(',')[0]);
  var out = [{ id: root.getId(), nom: root.getName() }];
  var it = root.getFolders();
  while (it.hasNext()) { var f = it.next(); out.push({ id: f.getId(), nom: root.getName() + ' / ' + f.getName() }); }
  return out;
}

/** Envoi d'une image (ou PDF / vidéo courte) depuis l'ERP : p = { name, mime, data (base64), dossier } */
function apiUploadMedia_(p) {
  var mime = String(p.mime || '');
  if (!/^image\//.test(mime) && mime !== 'application/pdf' && !/^video\//.test(mime)) throw new Error('Seulement des images, des PDF ou des vidéos.');
  var dossiers = apiMediasDossiers_();
  var cible = dossiers.filter(function (d) { return d.id === p.dossier; })[0] || dossiers[0];
  var bytes = Utilities.base64Decode(String(p.data || ''));
  if (!bytes.length) throw new Error('Fichier vide.');
  if (bytes.length > 30 * 1024 * 1024) throw new Error('Fichier trop lourd (30 Mo max).');
  var nom = String(p.name || 'image').replace(/[\\/:*?"<>|#%]/g, '_').slice(0, 120);
  var file = DriveApp.getFolderById(cible.id).createFile(Utilities.newBlob(bytes, mime, nom));
  cacheRemoveBig_('medias_v3');
  logAction_('image ajoutée', cible.nom + ' / ' + nom);
  return { id: file.getId(), nom: nom, url: file.getUrl(), dossier: cible.nom };
}

/** CacheService limite chaque valeur à 100 Ko : on découpe en morceaux */
function cachePutBig_(key, str, ttl) {
  try {
    var cache = CacheService.getScriptCache();
    var size = 45000, parts = {}, n = Math.ceil(str.length / size);
    if (n > 60) return;
    for (var i = 0; i < n; i++) parts[key + '_' + i] = str.substr(i * size, size);
    parts[key + '_n'] = String(n);
    cache.putAll(parts, ttl);
  } catch (e) { }
}
function cacheGetBig_(key) {
  try {
    var cache = CacheService.getScriptCache();
    var n = +(cache.get(key + '_n') || 0);
    if (!n) return null;
    var keys = []; for (var i = 0; i < n; i++) keys.push(key + '_' + i);
    var got = cache.getAll(keys), out = '';
    for (var j = 0; j < n; j++) { if (got[keys[j]] === undefined || got[keys[j]] === null) return null; out += got[keys[j]]; }
    return out;
  } catch (e) { return null; }
}
function cacheRemoveBig_(key) {
  try { CacheService.getScriptCache().remove(key + '_n'); } catch (e) { }
}

/** Miniature en base64 (marche même si le navigateur n'est pas connecté au compte Google de la structure) */
function apiMiniature_(p) {
  var key = 'th_' + p.id;
  var cache = CacheService.getScriptCache();
  var hit = cache.get(key);
  if (hit) return hit;   // n'est en cache que si le contrôle d'accès est déjà passé
  assertAutorise_(p.id);
  try {
    var blob = DriveApp.getFileById(p.id).getThumbnail();
    if (!blob) return '';
    var data = 'data:' + (blob.getContentType() || 'image/png') + ';base64,' + Utilities.base64Encode(blob.getBytes());
    if (data.length < 95000) cache.put(key, data, 21600);
    return data;
  } catch (e) { return ''; }
}

/** Texte d'un Doc (post, plan de publication...) */
function apiTexteDoc_(p) {
  assertAutorise_(p.id);
  var doc = DocumentApp.openById(p.id);
  return { nom: doc.getName(), texte: doc.getBody().getText().slice(0, 8000), url: doc.getUrl() };
}

// ---------- Docs importants + rapports Claude ----------
function listFolder_(id, max) {
  if (!id) return [];
  var folder = DriveApp.getFolderById(id);
  var out = [];
  var it = folder.getFiles();
  while (it.hasNext() && out.length < (max || 100)) {
    var f = it.next();
    out.push({ id: f.getId(), nom: f.getName(), mime: f.getMimeType(), url: f.getUrl(), modif: f.getLastUpdated().getTime(), date: fmtDateTime_(f.getLastUpdated()) });
  }
  return out.sort(function (a, b) { return b.modif - a.modif; });
}

function apiDocs_() {
  var c = getConfig_();
  var fixed = [
    ['Modèle économique (prévisionnel + caisse)', c.ID_MODELE_ECO, 'sheet'],
    ['Réponses inscriptions cours', c.ID_COURS_REPONSES, 'sheet'],
    ['Réponses inscriptions stages', c.ID_STAGES_REPONSES, 'sheet'],
    ['Fiches de présence cours', c.ID_FICHES_COURS, 'sheet'],
    ['Fiches de présence stages', c.ID_FICHES_STAGES, 'sheet'],
    ['Modèle de facture', c.ID_MODELE_FACTURE, 'doc']
  ].map(function (x) {
    return { nom: x[0], url: x[2] === 'sheet' ? 'https://docs.google.com/spreadsheets/d/' + x[1] : 'https://docs.google.com/document/d/' + x[1] };
  });
  var admin = [];
  try { admin = listFolder_(c.ID_DOSSIER_DOCS, 60); } catch (e) { }
  return { fixes: fixed, admin: admin };
}

function apiRapports_() {
  var c = getConfig_();
  var files = [];
  try { files = listFolder_(c.ID_DOSSIER_RAPPORTS, 60); } catch (e) { }
  return files.map(function (f) {
    var n = norm_(f.nom);
    f.type = n.indexOf('veille') >= 0 ? 'Veille matos' : (n.indexOf('mail') >= 0 ? 'Tri mails' : 'Autre');
    return f;
  });
}

/**
 * Médias / docs : on ne lit que des fichiers rangés dans les dossiers configurés
 * (médias, rapports, admin), pas n'importe quel fichier du Drive du compte.
 */
function assertAutorise_(fileId) {
  var c = getConfig_();
  var allowed = String(c.IDS_DOSSIERS_MEDIAS || '').split(',').concat([c.ID_DOSSIER_PHOTOS, c.ID_DOSSIER_RAPPORTS, c.ID_DOSSIER_DOCS])
    .map(function (s) { return String(s || '').trim(); }).filter(String);
  var seen = {};
  function up(file, depth) {
    if (depth > 6) return false;
    var it = file.getParents();
    while (it.hasNext()) {
      var f = it.next();
      var id = f.getId();
      if (allowed.indexOf(id) >= 0) return true;
      if (!seen[id]) { seen[id] = true; if (up(f, depth + 1)) return true; }
    }
    return false;
  }
  if (!up(DriveApp.getFileById(fileId), 0)) throw new Error('Fichier hors des dossiers autorisés.');
}
