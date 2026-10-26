/** ERP Association - 2.0.0 - genere par build.js - ne pas editer a la main */

/* ===== Util.gs ===== */
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

/* ===== Config.gs ===== */
/**
 * Config.gs : parametres + installation (setup)
 *
 * TOUTES les valeurs d'identite sont vides par defaut. Rien n'est code en dur :
 * apres `setup()`, elles se remplissent dans l'onglet "Config" du classeur ERP,
 * sans toucher au code. C'est le seul fichier a renseigner pour adapter l'ERP
 * a une autre structure.
 *
 * Trois familles de cles :
 *   1. identite      : nom, contacts, coordonnees de paiement, mentions legales
 *   2. raccordement   : identifiants des Google Sheets / Forms / Drive utilises
 *   3. vocabulaire    : les mots affiches aux utilisateurs et dans les mails
 *                       (adherent / cours / stage / lieu), pour qu'une MJC,
 *                       une ecole de musique ou un club sportif n'aient pas a
 *                       modifier une ligne de code.
 */

var DEFAULT_CONFIG = [
  // [cle, valeur, description]

  // --- 1. identite de la structure ---------------------------------------
  ['NOM_STRUCTURE', '', 'Nom affiche partout (ERP, mails, documents)'],
  ['EMAIL_STRUCTURE', '', 'Adresse qui signe les mails sortants'],
  ['TEL_STRUCTURE', '', 'Telephone affiche dans la signature des mails'],
  ['ADRESSE_STRUCTURE', '', 'Adresse postale, utilisee sur les devis et factures'],
  ['SITE_WEB', '', 'Site public, affiche en pied de mail (sans https://)'],
  ['NOM_ASSO', '', 'Nom juridique, utilise sur les factures que les prestataires adressent a la structure'],
  ['SIRET_ASSO', '', 'SIRET (ou numero d\'immatriculation) de la structure'],
  ['PIN', '', "Code d'acces a l'ERP (vide = pas de code). Chiffres ou texte"],
  ['COULEUR_ACCENT', '#9b1017', 'Couleur du titre et des intertitres dans les mails'],
  ['COULEUR_FILET', '#ffd400', 'Couleur des filets du bandeau de mail'],
  ['BASELINE', '', 'Sous-titre du bandeau de mail (ex. "Ecole de musique associative")'],
  ['MOT_ACCENTUE', '', 'Mot du nom a teinter dans la couleur d\'accent sur le bandeau de mail (vide = nom entier en noir)'],

  // --- 2. raccordement aux fichiers Google -------------------------------
  // A remplir avec les identifiants de VOS fichiers. L'identifiant est la
  // partie centrale de l'URL : docs.google.com/spreadsheets/d/<IDENTIFIANT>/edit
  ['ID_COURS_REPONSES', '', 'Sheet des reponses du Google Form d\'inscription aux cours'],
  ['ONGLET_COURS_REPONSES', '', 'Nom de l\'onglet des reponses cours (vide = 1er onglet)'],
  ['ID_STAGES_REPONSES', '', 'Sheet des reponses du Google Form d\'inscription aux stages'],
  ['ONGLET_STAGES_REPONSES', '', 'Nom de l\'onglet des reponses stages (vide = 1er onglet)'],
  ['ID_FORM_COURS', '', 'Google Form cours (lu pour recopier les intitules exacts des creneaux)'],
  ['ID_FORM_STAGES', '', 'Google Form stages (intitules exacts des stages)'],
  ['ID_FICHES_COURS', '', 'Classeur des feuilles de presence des cours (appel depuis l\'ERP)'],
  ['ID_FICHES_STAGES', '', 'Classeur des feuilles de presence des stages'],
  ['ID_MODELE_ECO', '', 'Classeur du modele economique (previsionnel et compta reelle)'],
  ['ONGLET_COMPTE', 'Compte', 'Onglet du modele economique ou sont les depenses / recettes reelles'],
  ['ONGLET_SYNTHESE', 'Synthese', 'Onglet du modele economique avec le previsionnel affiche dans la Compta ERP'],
  ['ID_LOGO', '', 'Image du logo affichee dans l\'ERP (fichier Drive)'],
  ['ID_MODELE_FACTURE', '', 'Google Doc servant de modele de devis / facture'],
  ['ID_DOSSIER_DEVIS', '', 'Dossier Drive des devis (les factures vont dans un sous-dossier Factures)'],
  ['ID_DOSSIER_JUSTIFICATIFS', '', 'Dossier des justificatifs compta (cree au setup si vide)'],
  ['ID_DOSSIER_RAPPORTS', '', 'Dossier des rapports generes (cree au setup si vide)'],
  ['ID_DOSSIER_PHOTOS', '', 'Dossier des visuels affiche dans l\'ecran Comm'],
  ['IDS_DOSSIERS_MEDIAS', '', 'Dossiers medias supplementaires, separes par des virgules'],
  ['ID_DOSSIER_DOCS', '', 'Dossier des documents administratifs importants'],

  // --- 3. vocabulaire affiche --------------------------------------------
  // Change ces mots et toute l'interface, les mails et les documents suivent.
  ['VOC_MEMBRE', 'adherent', 'Comment s\'appelle une personne inscrite (adherent, eleve, licencie, participant)'],
  ['VOC_MEMBRES', 'adherents', 'Le meme mot au pluriel'],
  ['VOC_COURS', 'cours', 'Activite reguliere hebdomadaire (cours, atelier, entrainement, seance)'],
  ['VOC_COURS_PL', 'cours', 'Le meme mot au pluriel'],
  ['VOC_STAGE', 'stage', 'Activite intensive sur quelques jours (stage, camp, session)'],
  ['VOC_STAGE_PL', 'stages', 'Le meme mot au pluriel'],
  ['VOC_LIEU', 'sur place', 'Ou l\'on remet un cheque ou se deroule l\'activite ("sur place", "au local", "au gymnase")'],
  ['VOC_ACTIVITE', 'activite', 'Nature de la remuneration de terrain, par opposition a l\'administratif et a la technique'],

  // --- 4. paiements et facturation ---------------------------------------
  ['IBAN_TITULAIRE', '', 'Titulaire du compte, tel qu\'il doit apparaitre sur un virement'],
  ['IBAN', '', 'IBAN communique aux adherents et aux clients'],
  ['BIC', '', 'BIC du compte'],
  ['ADHESION', '0', 'Adhesion annuelle par personne inscrite, en euros (0 = pas d\'adhesion)'],
  ['PROCHAIN_NUM_FACTURE', '2026-001', 'Numero de la prochaine facture emise'],
  ['ECHEANCE_FACTURE_JOURS', '30', 'Delai de paiement affiche sur les factures emises'],
  ['PROCHAIN_NUM_DEVIS', 'DEV-2026-001', 'Numero du prochain devis'],
  ['VALIDITE_DEVIS_JOURS', '30', 'Duree de validite d\'un devis'],
  ['SIGNATAIRE', '', 'Nom et fonction de la personne qui signe les devis'],
  ['MENTION_TVA', 'TVA non applicable, article 293 B du CGI', 'Mention a reprendre sur les factures. Vider si assujetti'],
  ['DELAI_PAIEMENT_INTERV', '60', 'Delai affiche sur les factures des prestataires (jours). En France le maximum legal entre professionnels est 60 jours date de facture, ou 45 jours fin de mois'],
  ['PENALITES_INTERV', 'En cas de retard de paiement, penalites au taux de 3 fois le taux d\'interet legal et indemnite forfaitaire de recouvrement de 40 EUR (art. L441-10 et D441-5 du code de commerce).', 'Mention legale entre professionnels. Vider cette ligne pour ne pas l\'afficher'],
  ['DATE_CONVENTION', '', 'Date de signature de la convention annuelle de prestation (rappelee sur les factures lissees)'],

  // --- 5. fonctionnement -------------------------------------------------
  ['COULEURS_AGENDA', '', 'Google Agenda : "numero de couleur = categorie|code couleur", separes par des virgules. Ex. 7=Cours|#2b5f9e, 5=Stage|#e0a800'],
  ['DEBUT_SAISON_COURS', '', 'Date (yyyy-mm-dd) du lundi de la premiere seance. L\'appel s\'ouvre sur la seance de la semaine en cours'],
  ['VACANCES_COURS', '', 'Jours sans cours (vacances et feries), format debut>fin separes par des point-virgules'],
  ['RELANCE_APRES_JOURS', '10', 'Brouillon de relance si pas paye X jours apres inscription'],
  ['RELANCE_INTERVALLE_JOURS', '7', 'Pas plus d\'une relance tous les X jours'],
  ['RELANCES_AUTO', 'non', 'Brouillons de relance crees automatiquement chaque matin (oui/non). Sinon, bouton dans l\'ERP'],
  ['CONFIRMATION_AUTO', 'oui', 'Brouillon de confirmation a chaque nouvelle inscription (oui/non)'],
  ['DIGEST_EMAILS', '', 'Recap quotidien 7h (adresses separees par des virgules, vide = desactive)'],
  ['INTERVENANTS', '', 'Personnes remunerees, separees par des point-virgules'],
  ['CONTACT_PAIEMENTS', '', 'Qui suit les arrangements de paiement. Cite dans les relances pour que les familles ayant deja convenu de modalites n\'en tiennent pas compte. Vider pour ne pas le mentionner'],
  ['REDUC_MOTIFS', '', 'Motifs de reduction propres a la structure (nom d\'une commune conventionnee, dispositif local...), separes par des point-virgules. Une inscription portant un de ces motifs ne recoit jamais de montant chiffre par mail : le calcul reste humain'],
  ['SEUIL_MIN_COURS', '4', 'Inscrits minimum pour maintenir un cours'],
  ['SEUIL_MIN_STAGE', '6', 'Inscrits minimum par groupe de stage']
];

/**
 * Exemples de creneaux hebdomadaires, ecrits dans l'onglet "Cours" au premier
 * setup. Remplace-les par les tiens, ou vide ce tableau pour partir d'une
 * feuille blanche. La colonne Cle est le texte cherche dans la reponse du
 * Google Form pour rattacher une inscription a un creneau.
 */
var DEFAULT_COURS = [
  // Cle | Libelle | Jour | Horaire | Public | Tarif annuel | Places max
  ['Lundi 17h00', 'Groupe 6-9 ans', 'Lundi', '17h00-18h30', 'Enfants 6-9 ans', 300, 12],
  ['Mardi 17h00', 'Groupe 9-13 ans', 'Mardi', '17h00-19h00', 'Enfants 9-13 ans', 350, 12],
  ['Mercredi 11h00', 'Groupe 4-6 ans', 'Mercredi', '11h00-12h00', 'Enfants 4-6 ans', 250, 12],
  ['Jeudi 19h15', 'Groupe adultes', 'Jeudi', '19h15-21h15', 'Adultes 18+', 350, 8]
];

/**
 * Exemples de sessions courtes (vacances scolaires), ecrits dans l'onglet
 * "Stages" au premier setup. Meme logique : la Cle est cherchee dans la
 * reponse du Form.
 */
var DEFAULT_STAGES = [
  // Cle dates | Groupe | Nom | Dates affichees | Public | Tarif | Max | Debut (yyyy-mm-dd)
  ['semaine 1', 'matin', 'Session 1', 'Du 26 au 30 octobre 2026', 'Matin 10h-12h, 4-6 ans', 200, 12, '2026-10-26'],
  ['semaine 1', 'apres-midi', 'Session 1', 'Du 26 au 30 octobre 2026', 'Apres-midi 13h30-16h30, 7 ans et +', 250, 12, '2026-10-26']
];

var TABLES = {
  Config: ['Cle', 'Valeur', 'Description'],
  Cours: ['Cle', 'Libelle', 'Jour', 'Horaire', 'Public', 'Tarif', 'Max', 'Code'],
  Stages: ['Cle', 'Groupe', 'Nom', 'Dates', 'Public', 'Tarif', 'Max', 'Debut'],
  Suivi: ['ID', 'Source', 'Nom', 'Prenom', 'Paye', 'MontantRecu', 'ModeRecu', 'Statut', 'Notes', 'DerniereRelance', 'Confirmation', 'MajPar', 'MajLe'],
  Factures: ['ID', 'Numero', 'Date', 'Client', 'Objet', 'Montant', 'Echeance', 'Statut', 'PayeeLe', 'LienDoc', 'LienPdf', 'Type', 'Details'],
  NotesSite: ['ID', 'Date', 'Page', 'Description', 'Priorite', 'Statut', 'Auteur'],
  Rappels: ['ID', 'Date', 'Heure', 'Type', 'Titre', 'Detail', 'Fait'],
  Attente: ['ID', 'Date', 'Source', 'CleGroupe', 'Nom', 'Prenom', 'Naissance', 'Responsable', 'Tel', 'Email', 'Notes', 'Statut', 'MajPar', 'MajLe'],
  Achats: ['ID', 'Date', 'Beneficiaire', 'Type', 'Objet', 'Montant', 'Echeance', 'Statut', 'PayeLe', 'Mode', 'Lien', 'Pieces', 'Notes', 'MajPar', 'MajLe'],
  Intervenants: ['ID', 'Nom', 'Statut', 'Adresse', 'CP', 'Ville', 'SIRET', 'Email', 'Tel', 'IBAN', 'BIC', 'MentionTVA', 'Assurance', 'ProchainNum', 'Actif', 'MajPar', 'MajLe'],
  FacturesInterv: ['ID', 'Numero', 'Date', 'Intervenant', 'PeriodeDebut', 'PeriodeFin', 'Periode', 'Lignes', 'Total', 'Fichier', 'LienPdf', 'LienDoc', 'Statut', 'AchatID', 'Notes', 'MajPar', 'MajLe'],
  Presences: ['Cle', 'Source', 'CleGroupe', 'Onglet', 'Seance', 'Nom', 'Statut', 'Date', 'Par', 'Prenom'],
  Log: ['Date', 'Qui', 'Action', 'Detail']
};

/**
 * A lancer UNE fois depuis l'editeur Apps Script (bouton Executer).
 * Cree les onglets, pre-remplit la config, cree les dossiers et installe les
 * declencheurs. Relancable sans risque : les onglets et les valeurs existants
 * ne sont pas ecrases, seules les colonnes et les cles manquantes sont ajoutees.
 */
function setup() {
  proprietaireSeulement_();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error('Le script doit etre lie au classeur ERP (Extensions > Apps Script depuis le Sheet).');
  PropertiesService.getScriptProperties().setProperty('ERP_SS_ID', ss.getId());

  Object.keys(TABLES).forEach(function (name) {
    var sh = ss.getSheetByName(name);
    if (!sh) sh = ss.insertSheet(name);
    if (sh.getLastRow() === 0) {
      sh.getRange(1, 1, 1, TABLES[name].length).setValues([TABLES[name]])
        .setFontWeight('bold').setBackground('#151313').setFontColor('#F4C20D');
      sh.setFrozenRows(1);
    } else {
      // ajoute les colonnes manquantes a droite (mise a jour de version)
      var h = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(String);
      TABLES[name].forEach(function (col) {
        if (h.indexOf(col) < 0) {
          sh.getRange(1, sh.getLastColumn() + 1).setValue(col).setFontWeight('bold');
        }
      });
    }
  });

  var cfg = ss.getSheetByName('Config');
  cfg.getRange('B:B').setNumberFormat('@'); // texte brut : "2026-001" ne doit pas devenir une date
  if (cfg.getLastRow() < 2) cfg.getRange(2, 1, DEFAULT_CONFIG.length, 3).setValues(DEFAULT_CONFIG);
  else {
    var existing = cfg.getRange(2, 1, cfg.getLastRow() - 1, 1).getValues().map(function (r) { return String(r[0]); });
    DEFAULT_CONFIG.forEach(function (row) { if (existing.indexOf(row[0]) < 0) cfg.appendRow(row); });
  }
  var cours = ss.getSheetByName('Cours');
  if (cours.getLastRow() < 2 && DEFAULT_COURS.length) cours.getRange(2, 1, DEFAULT_COURS.length, 7).setValues(DEFAULT_COURS);
  var stages = ss.getSheetByName('Stages');
  if (stages.getLastRow() < 2 && DEFAULT_STAGES.length) stages.getRange(2, 1, DEFAULT_STAGES.length, 8).setValues(DEFAULT_STAGES);

  // Dossiers crees a cote du classeur ERP
  var parent = DriveApp.getFileById(ss.getId()).getParents();
  var root = parent.hasNext() ? parent.next() : DriveApp.getRootFolder();
  var c = getConfig_(true);
  if (!c.ID_DOSSIER_JUSTIFICATIFS) setConfig_('ID_DOSSIER_JUSTIFICATIFS', getOrCreateFolder_(root, 'ERP - Justificatifs compta').getId());
  if (!c.ID_DOSSIER_RAPPORTS) setConfig_('ID_DOSSIER_RAPPORTS', getOrCreateFolder_(root, 'ERP - Rapports').getId());

  installTriggers();
  logAction_('setup', 'installation OK');
  return 'Setup termine. Remplis maintenant l\'onglet Config, puis Deployer > Nouveau deploiement > Application Web.';
}

/** Ce qui manque encore pour que l'ERP soit utilisable. Affiche par le menu. */
function verifierConfig() {
  var c = getConfig_(true);
  var obligatoires = {
    NOM_STRUCTURE: 'nom de la structure',
    EMAIL_STRUCTURE: 'adresse qui signe les mails',
    ID_COURS_REPONSES: 'Sheet des reponses du Form cours'
  };
  var conseillees = {
    ID_STAGES_REPONSES: 'Sheet des reponses du Form stages',
    ID_MODELE_ECO: 'classeur du modele economique',
    IBAN: 'IBAN pour les relances',
    DEBUT_SAISON_COURS: 'date de la premiere seance (appel)',
    PIN: 'code d\'acces'
  };
  var manque = [], tiede = [];
  Object.keys(obligatoires).forEach(function (k) { if (!String(c[k] || '').trim()) manque.push(k + ' (' + obligatoires[k] + ')'); });
  Object.keys(conseillees).forEach(function (k) { if (!String(c[k] || '').trim()) tiede.push(k + ' (' + conseillees[k] + ')'); });
  var msg = manque.length ? 'A remplir avant utilisation :\n- ' + manque.join('\n- ') : 'Configuration minimale complete.';
  if (tiede.length) msg += '\n\nOptionnel, desactive tant que c\'est vide :\n- ' + tiede.join('\n- ');
  try { SpreadsheetApp.getUi().alert('Configuration ERP', msg, SpreadsheetApp.getUi().ButtonSet.OK); } catch (e) { }
  return msg;
}

function getOrCreateFolder_(parent, name) {
  var it = parent.getFoldersByName(name);
  return it.hasNext() ? it.next() : parent.createFolder(name);
}

var _cfgCache = null;
function getConfig_(fresh) {
  if (_cfgCache && !fresh) return _cfgCache;
  var out = {};
  DEFAULT_CONFIG.forEach(function (r) { out[r[0]] = r[1]; });
  try {
    var sh = erpSS_().getSheetByName('Config');
    if (sh && sh.getLastRow() > 1) {
      sh.getRange(2, 1, sh.getLastRow() - 1, 2).getValues().forEach(function (r) {
        if (r[0]) out[String(r[0]).trim()] = String(r[1]).trim();
      });
    }
  } catch (e) { /* avant setup */ }
  _cfgCache = out;
  return out;
}

function setConfig_(key, value) {
  var sh = erpSheet_('Config');
  var vals = sh.getRange(2, 1, Math.max(sh.getLastRow() - 1, 1), 1).getValues();
  for (var i = 0; i < vals.length; i++) {
    if (String(vals[i][0]).trim() === key) {
      sh.getRange(i + 2, 2).setValue(value);
      _cfgCache = null;
      return;
    }
  }
  sh.appendRow([key, value, '']);
  _cfgCache = null;
}

/** (Re)installe les declencheurs : inscriptions (form submit) + taches quotidiennes 7h */
function installTriggers() {
  proprietaireSeulement_();
  ScriptApp.getProjectTriggers().forEach(function (t) {
    var f = t.getHandlerFunction();
    if (['onInscriptionCours', 'onInscriptionStage', 'dailyJob', 'rechaufferCaches'].indexOf(f) >= 0) ScriptApp.deleteTrigger(t);
  });
  var c = getConfig_(true);
  if (c.ID_COURS_REPONSES) ScriptApp.newTrigger('onInscriptionCours').forSpreadsheet(c.ID_COURS_REPONSES).onFormSubmit().create();
  if (c.ID_STAGES_REPONSES) ScriptApp.newTrigger('onInscriptionStage').forSpreadsheet(c.ID_STAGES_REPONSES).onFormSubmit().create();
  ScriptApp.newTrigger('dailyJob').timeBased().everyDays(1).atHour(7).inTimezone(TZ).create();
  // rechauffage du cache toutes les 10 min : les ecrans lisent une reponse deja prete
  poserRechauffage_();
}

/** Menu dans le Sheet ERP */
function onOpen() {
  SpreadsheetApp.getUi().createMenu('ERP')
    .addItem('Installer / mettre a jour (setup)', 'setup')
    .addItem('Verifier la configuration', 'verifierConfig')
    .addItem('Reinstaller les declencheurs', 'installTriggers')
    .addItem('Lancer la tache quotidienne maintenant', 'dailyJob')
    .addItem('Preparer les ecrans maintenant (cache)', 'rechaufferCaches')
    .addItem("Completer l'onglet Suivi (une ligne par inscrit)", 'completerSuivi')
    .addItem('Mettre a jour le modele economique (remuneration par intervenant)', 'majModeleEco')
    .addToUi();
}

/** setup / declencheurs : uniquement depuis l'editeur ou le menu du Sheet, par le proprietaire */
function proprietaireSeulement_() {
  var actif = '', effectif = '';
  try { actif = Session.getActiveUser().getEmail(); effectif = Session.getEffectiveUser().getEmail(); } catch (e) { }
  if (!actif || actif !== effectif) throw new Error('Reserve au proprietaire du script (editeur Apps Script ou menu ERP).');
  try {
    var owner = DriveApp.getFileById(erpSS_().getId()).getOwner();
    if (owner && owner.getEmail() && owner.getEmail() !== actif) throw new Error('Reserve au proprietaire du classeur ERP (' + owner.getEmail() + ').');
  } catch (e) { if (/Reserve/.test(e.message)) throw e; }
}

/** Appel legitime = vrai declencheur installe (triggerUid connu) ou lancement manuel par le proprietaire */
function viaDeclencheurOuProprio_(ev) {
  if (ev && ev.triggerUid) {
    var ok = ScriptApp.getProjectTriggers().some(function (t) { return t.getUniqueId() === String(ev.triggerUid); });
    if (ok) return;
    throw new Error('Declencheur inconnu.');
  }
  proprietaireSeulement_();
}

/* ===== Libelles.gs ===== */
/**
 * Libelles.gs : couche vocabulaire.
 *
 * Le modele de donnees garde des noms stables en interne (eleve, cours, stage).
 * Ce qui s'affiche a l'ecran, part dans un mail ou atterrit dans un document
 * passe par ici, et vient de l'onglet Config. Une ecole de musique met
 * "eleve / cours / stage", un club sportif "licencie / entrainement / camp",
 * une MJC "adherent / atelier / session", sans toucher une ligne de code.
 *
 *   L_('membre')        -> 'adherent'
 *   L_('membres')       -> 'adherents'
 *   Lmaj_('membres')    -> 'Adherents'
 *   Lle_('membre')      -> "l'adherent"  / 'le licencie'
 *   Ldu_('membre')      -> "de l'adherent" / 'du licencie'
 *   Lun_('membre')      -> 'un adherent'
 *
 * Les cles inconnues renvoient la cle elle-meme : un oubli se voit a l'ecran
 * mais ne casse rien.
 */

var VOC_DEFAUT = {
  membre: 'adherent',
  membres: 'adherents',
  cours: 'cours',
  coursPl: 'cours',
  stage: 'stage',
  stagePl: 'stages',
  lieu: 'sur place',
  activite: 'activite'
};

var VOC_CLES = {
  membre: 'VOC_MEMBRE',
  membres: 'VOC_MEMBRES',
  cours: 'VOC_COURS',
  coursPl: 'VOC_COURS_PL',
  stage: 'VOC_STAGE',
  stagePl: 'VOC_STAGE_PL',
  lieu: 'VOC_LIEU',
  activite: 'VOC_ACTIVITE'
};

var _vocCache = null;

/** Le vocabulaire complet, lu une fois par execution. */
function vocabulaire_(fresh) {
  if (_vocCache && !fresh) return _vocCache;
  var c = getConfig_(fresh);
  var out = {};
  Object.keys(VOC_CLES).forEach(function (k) {
    var v = String(c[VOC_CLES[k]] || '').trim();
    out[k] = v || VOC_DEFAUT[k];
  });
  _vocCache = out;
  return out;
}

/** Libelle en minuscules. */
function L_(cle) {
  var v = vocabulaire_();
  return v[cle] !== undefined ? v[cle] : String(cle);
}

/** Libelle avec une majuscule initiale, pour un debut de phrase ou un titre. */
function Lmaj_(cle) {
  var s = L_(cle);
  return s ? s.charAt(0).toUpperCase() + s.slice(1) : s;
}

/** "le cours" / "l'atelier" : elide l'article devant une voyelle ou un h muet. */
function Lle_(cle) {
  var s = L_(cle);
  return (voyelle_(s) ? "l'" : 'le ') + s;
}

/** "un cours" / "un atelier" */
function Lun_(cle) {
  return 'un ' + L_(cle);
}

/** "du licencie" / "de l'adherent" : complement du nom, elide si besoin. */
function Ldu_(cle) {
  var s = L_(cle);
  return (voyelle_(s) ? "de l'" : 'du ') + s;
}

function voyelle_(s) {
  return /^[aeiouyhâàéèêîïôöùûü]/i.test(s || '');
}

/**
 * Remplace les marqueurs {{membre}}, {{cours}}, {{lieu}}... dans un texte.
 * Sert aux gabarits de mails et de documents, ou l'on ecrit la phrase une fois
 * et ou le vocabulaire se substitue a la lecture.
 */
function Ltxt_(txt) {
  if (!txt) return '';
  var v = vocabulaire_();
  return String(txt).replace(/\{\{(\w+)\}\}/g, function (tout, cle) {
    if (v[cle] !== undefined) return v[cle];
    var maj = cle.charAt(0).toLowerCase() + cle.slice(1);
    if (v[maj] !== undefined) return v[maj].charAt(0).toUpperCase() + v[maj].slice(1);
    return tout;
  });
}

/* ===== Inscriptions.gs ===== */
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

/* ===== Compta.gs ===== */
/**
 * Compta.gs : travaille sur le petit tableau de caisse existant de Modele_economique
 * (Affichage, Matos : boule, Portique, ... puis Cash / Chèque / Virement / Total).
 *
 * Règles :
 *  - le tableau est retrouvé par ses libellés (cellules "Cash" et "Total"), pas par adresse fixe
 *  - une DÉPENSE = une nouvelle ligne insérée juste au-dessus de "Cash" (montant négatif)
 *  - une RECETTE = ajoutée au total du mode (Cash / Chèque / Virement), la cellule garde
 *    l'historique sous forme de formule (=1000+300+250) + une note datée
 *  - 2 colonnes ajoutées à droite du montant : Date et Justificatif (lien Drive)
 *  - on insère des CELLULES (pas des lignes entières) pour ne pas décaler les autres tableaux
 */

var MODES = ['cash', 'cheque', 'virement'];

function comptaLocate_() {
  var c = getConfig_();
  var ss = SpreadsheetApp.openById(c.ID_MODELE_ECO);
  var sheets = ss.getSheets();
  // l'onglet "Compte aussi" (Config ONGLET_COMPTE) est cherché en premier
  var pref = norm_(c.ONGLET_COMPTE || '');
  if (pref) sheets = sheets.slice().sort(function (a, b) { return (norm_(b.getName()) === pref) - (norm_(a.getName()) === pref); });
  for (var s = 0; s < sheets.length; s++) {
    var sh = sheets[s];
    var v = sh.getDataRange().getValues();
    for (var r = 0; r < v.length; r++) {
      for (var col = 0; col < v[r].length; col++) {
        if (norm_(v[r][col]) !== 'cash') continue;
        // on a "Cash" : cherche Chèque, Virement, Total dans la même colonne en dessous
        var rows = { cash: r };
        var totalRow = -1;
        for (var k = r + 1; k < Math.min(r + 12, v.length); k++) {
          var lab = norm_(v[k][col]);
          if (lab === 'cheque' || lab === 'cheques') rows.cheque = k;
          else if (lab === 'virement' || lab === 'virements') rows.virement = k;
          else if (lab === 'total') { totalRow = k; break; }
        }
        if (totalRow < 0) continue;
        // début du bloc : on remonte jusqu'à 3 lignes vides d'affilée
        var scanStart = r, empties = 0;
        for (var u = r - 1; u >= 0 && empties < 3; u--) {
          var filled = String(v[u][col]).trim() !== '' || String(v[u][col + 1]).trim() !== '';
          if (filled) { scanStart = u; empties = 0; } else empties++;
        }
        var L = { ss: ss, sheet: sh, col: col + 1, cashRow: r + 1, totalRow: totalRow + 1, bloque: '',
          modeRows: { cash: rows.cash + 1, cheque: rows.cheque !== undefined ? rows.cheque + 1 : null, virement: rows.virement !== undefined ? rows.virement + 1 : null } };
        var tf = sh.getRange(totalRow + 1, col + 2).getFormula();
        var fm = tf && tf.match(/^=\s*SUM\(\s*\$?([A-Z]+)\$?(\d+)\s*:\s*\$?([A-Z]+)\$?(\d+)\s*\)\s*$/i);
        var amountCol = columnLetter_(col + 2);
        if (fm && fm[1].toUpperCase() === amountCol && fm[3].toUpperCase() === amountCol && Math.min(+fm[2], +fm[4]) < r + 1) {
          L.totalMode = 'sum'; L.start = Math.min(+fm[2], +fm[4]);
        } else if (!tf) {
          L.totalMode = 'value'; L.start = scanStart + 1;
          var somme = 0;
          for (var q = scanStart; q < totalRow; q++) somme += num_(v[q][col + 1]);
          if (Math.abs(somme - num_(v[totalRow][col + 1])) > 0.005) L.bloque = 'Le Total de la caisse est une valeur tapée à la main qui ne correspond pas à la somme des lignes. Mets une formule =SUM(...) sur la cellule Total, puis réessaie.';
        } else {
          L.totalMode = 'autre'; L.start = scanStart + 1;
          L.bloque = 'La cellule Total de la caisse contient une formule que l\'ERP ne sait pas étendre sans risque (' + tf + '). Remplace-la par =SUM(' + amountCol + (scanStart + 1) + ':' + amountCol + totalRow + ') puis réessaie.';
        }
        // colonnes Date / Justificatif (à droite du montant) : vides, dates ou liens uniquement
        if (!L.bloque) {
          for (var q2 = L.start - 1; q2 <= totalRow; q2++) {
            for (var dc = 2; dc <= 3; dc++) {
              var x = v[q2] ? v[q2][col + dc] : '';
              if (x === '' || x === null || x === undefined || isDate_(x)) continue;
              var xs = norm_(x);
              if (/^https?:\/\//.test(xs) || xs === 'date' || xs === 'justificatif') continue;
              L.bloque = 'La caisse a quelque chose à droite du montant (' + columnLetter_(col + 1 + dc) + (q2 + 1) + ' : "' + String(x).slice(0, 30) + '"). L\'ERP a besoin des 2 colonnes à droite du montant libres pour la date et le justificatif.';
              break;
            }
            if (L.bloque) break;
          }
        }
        // en dessous : un tableau à cheval sur ces colonnes serait décalé par l'insertion
        if (!L.bloque) {
          for (var q3 = totalRow + 1; q3 < v.length; q3++) {
            var inBlock = false;
            for (var cc = col; cc <= col + 3; cc++) if (String(v[q3][cc] === undefined ? '' : v[q3][cc]).trim() !== '') inBlock = true;
            var left = col > 0 && String(v[q3][col - 1]).trim() !== '';
            var right = col + 4 < v[q3].length && String(v[q3][col + 4]).trim() !== '';
            if (inBlock && (left || right)) {
              L.bloque = 'Sous la caisse (ligne ' + (q3 + 1) + '), un tableau déborde sur les colonnes de la caisse : insérer une ligne le décalerait. Déplace la caisse ou ce tableau.';
              break;
            }
          }
        }
        return L;
      }
    }
  }
  throw new Error('Tableau de caisse introuvable dans Modele_economique (il faut une cellule "Cash" puis "Total" en dessous).');
}

/** Lecture du tableau + synthèse */
function apiCompta_() {
  var L = comptaLocate_();
  var sh = L.sheet;
  var n = L.totalRow - L.start + 1;
  var rng = sh.getRange(L.start, L.col, n, 4);
  var vals = rng.getValues();
  var forms = rng.getFormulas();
  var notes = rng.getNotes();
  var lignes = [];
  var modes = [];
  for (var i = 0; i < n; i++) {
    var row = L.start + i;
    var label = String(vals[i][0]).trim();
    var montant = vals[i][1];
    if (!label && (montant === '' || montant === null)) continue;
    var key = norm_(label).replace(/s$/, '');
    if (row === L.totalRow) continue;
    var item = { row: row, label: label, montant: num_(montant), formule: forms[i][1] || '', date: plain_(vals[i][2]), justificatif: String(vals[i][3] || ''), note: notes[i][1] || '' };
    if (MODES.indexOf(key) >= 0) modes.push(item); else lignes.push(item);
  }
  var total = num_(sh.getRange(L.totalRow, L.col + 1).getValue());
  var depenses = lignes.filter(function (l) { return l.montant < 0; }).reduce(function (a, l) { return a + l.montant; }, 0);
  var recettesAutres = lignes.filter(function (l) { return l.montant > 0; }).reduce(function (a, l) { return a + l.montant; }, 0);
  var recettesModes = modes.reduce(function (a, m) { return a + m.montant; }, 0);
  return {
    lignes: lignes, modes: modes, total: total,
    depenses: depenses, recettes: recettesModes + recettesAutres,
    synthese: syntheseModeleEco_(L.ss),
    synthesePoles: synthesePoles_(L.ss),
    onglet: sh.getName(),
    factures: apiFactures_().filter(function (f) { return f.Type !== 'Devis' && f.Statut !== 'Payée' && f.Statut !== 'Annulée'; }),
    bloque: L.bloque,
    lien: 'https://docs.google.com/spreadsheets/d/' + L.ss.getId() + '/edit#gid=' + sh.getSheetId(),
    emplacement: sh.getName() + '!' + columnLetter_(L.col) + L.start + ':' + columnLetter_(L.col + 3) + L.totalRow
  };
}

function columnLetter_(c) {
  var s = '';
  while (c > 0) { var m = (c - 1) % 26; s = String.fromCharCode(65 + m) + s; c = Math.floor((c - 1) / 26); }
  return s;
}

/**
 * Le tableau des poles de la Synthese : recettes, cout intervenants, marge, et la colonne
 * de scenario optimal quand elle existe. On ne recopie plus l onglet entier : seul ce
 * tableau est utile dans l ERP, le reste se lit dans le Sheet.
 */
function synthesePoles_(ss) {
  var nom = getConfig_().ONGLET_SYNTHESE || 'Synthese';
  var sh = ss.getSheetByName(nom);
  if (!sh) sh = ss.getSheets().filter(function (x) { return norm_(x.getName()).indexOf('synth') === 0; })[0];
  if (!sh) return null;
  var disp = sh.getDataRange().getDisplayValues();
  var raw = sh.getDataRange().getValues();
  for (var r = 0; r < disp.length; r++) {
    if (norm_(disp[r][0]) !== 'pole') continue;
    var larg = disp[r].length;
    while (larg > 0 && String(disp[r][larg - 1]).trim() === '') larg--;
    if (larg < 2) continue;
    var nett = function (x) { return String(x).replace(/\s+/g, ' ').trim(); };
    var entetes = disp[r].slice(0, larg).map(nett);
    var lignes = [], total = null;
    for (var i = r + 1; i < disp.length; i++) {
      var lab = String(disp[i][0]).trim();
      if (!lab) break;
      var o = {
        cells: disp[i].slice(0, larg).map(nett),
        nums: raw[i].slice(0, larg).map(function (x) { return typeof x === 'number' ? x : null; })
      };
      if (norm_(lab).indexOf('total') === 0) { total = o; break; }
      lignes.push(o);
    }
    if (lignes.length) return { entetes: entetes, lignes: lignes, total: total };
  }
  return null;
}

/** Quelques chiffres clés de la synthèse, retrouvés par libellé */
function syntheseModeleEco_(ss) {
  // clé cherchée -> libellé affiché. La comparaison se fait sur le DEBUT du libellé,
  // pour survivre aux precisions ajoutees a la main (« TOTAL ANNEE (PAYE) »...).
  var wanted = [
    ['total annee reel', 'Total encaissé à ce jour'],
    ['total annee', 'Total année (recettes)'],
    ['caisse', 'Caisse'],
    ['total a payer aux intervenants', 'À payer aux intervenants'],
    ['total equipe', 'À payer aux intervenants'],
    ['total potentiel', 'Rémunération potentielle'],
    ['resultat optimal', 'Résultat si 7 élèves partout'],
    ['resultat previ net', 'Résultat net'],
    ['resultat net reel', 'Résultat net réel'],
    ['resultat net', 'Résultat net'],
    ['cours annuels', 'Recettes cours annuels'],
    ['prestations collectivites et ecoles', 'Recettes prestations'],
    ['stages vacances', 'Recettes stages'],
    ['charges de structure - previ', 'Charges de structure'],
    ['charges de structure - reel', 'Charges réelles'],
    ['total charges', 'Charges de structure']
  ];
  var out = [], seen = {};
  // la clé d'unicité est le libellé affiché : deux intitulés différents du Sheet qui
  // désignent la même chose (TOTAL EQUIPE / TOTAL A PAYER) ne font qu'une ligne
  var pose = function (cle, libelle, valeur) {
    if (seen[libelle] || typeof valeur !== 'number') return;
    seen[libelle] = true; out.push({ label: libelle, valeur: valeur });
  };
  ss.getSheets().forEach(function (sh) {
    var v = sh.getDataRange().getValues();
    if (norm_(v.map(function (r) { return r.join(' '); }).join(' ')).indexOf('synthese annuelle') < 0) return;
    for (var r = 0; r < v.length; r++) {
      for (var c = 0; c < v[r].length; c++) {
        var k = norm_(v[r][c]).replace(/'/g, ' ');
        if (!k) continue;
        for (var w = 0; w < wanted.length; w++) {
          if (k.indexOf(wanted[w][0]) !== 0) continue;
          // la valeur est juste a droite, sinon (cas « TOTAL CHARGES » en legende) juste a gauche
          var droite = c + 1 < v[r].length ? v[r][c + 1] : null;
          var gauche = c > 0 ? v[r][c - 1] : null;
          if (typeof droite === 'number') pose(wanted[w][0], wanted[w][1], droite);
          else if (typeof gauche === 'number') pose(wanted[w][0], wanted[w][1], gauche);
          break;
        }
      }
    }
  });
  return out;
}

function assertWritable_(L) {
  if (L.bloque) throw new Error(L.bloque);
}

/**
 * Après insertion / suppression : le Total doit couvrir tout le bloc (du début jusqu'à la ligne au-dessus de Total).
 * Seulement si le Total est un SUM simple (ou une valeur qui collait à la somme).
 */
function fixTotal_(L, insertedRow) {
  var sh = L.sheet;
  var cell = sh.getRange(L.totalRow, L.col + 1);
  var f = cell.getFormula();
  var colL = columnLetter_(L.col + 1);
  var fm = f && f.match(/^=\s*SUM\(\s*\$?([A-Z]+)\$?(\d+)\s*:\s*\$?([A-Z]+)\$?(\d+)\s*\)\s*$/i);
  var a = fm ? Math.min(+fm[2], +fm[4]) : L.start;
  if (insertedRow && insertedRow < a) a = insertedRow;
  var nf = '=SUM(' + colL + a + ':' + colL + (L.totalRow - 1) + ')';
  if (f.replace(/\$/g, '').replace(/\s/g, '').toUpperCase() !== nf.toUpperCase()) {
    cell.setFormula(nf);
    logAction_('compta', 'Formule Total : ' + (f || '(valeur)') + ' -> ' + nf);
  }
}

/** Ajoute une dépense (ligne au-dessus de Cash) */
function apiAjouterDepense_(p) {
  var lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    var montant = Math.abs(num_(p.montant));
    if (!p.libelle || !montant) throw new Error('Libellé et montant obligatoires.');
    var L = comptaLocate_();
    assertWritable_(L);
    var sh = L.sheet;
    var ins = L.cashRow;
    sh.getRange(ins, L.col, 1, 4).insertCells(SpreadsheetApp.Dimension.ROWS);
    var d = parseDate_(p.date) || new Date();
    var just = /^https:\/\//.test(String(p.justificatif || '')) ? p.justificatif : '';
    sh.getRange(ins, L.col, 1, 4).setValues([[safeCell_(String(p.libelle)), -montant, d, just]]);
    sh.getRange(L.cashRow, L.col + 2).setNumberFormat('dd/mm/yyyy');
    sh.getRange(L.cashRow, L.col + 1).setNumberFormat(sh.getRange(L.start, L.col + 1).getNumberFormat() || '#,##0 €');
    SpreadsheetApp.flush();
    fixTotal_(comptaLocate_(), ins);
    logAction_('dépense', p.libelle + ' -' + montant);
    return true;
  } finally { lock.releaseLock(); }
}

/**
 * Modifie une ligne de la caisse : p = { row, label (contrôle), libelle, sens: 'depense'|'entree', montant (positif), date }
 * Lignes Cash / Chèque / Virement : seul le montant change (libellé fixe). L'ancienne valeur est gardée en note.
 */
function apiModifierLigneCompta_(p) {
  var lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    var L = comptaLocate_();
    assertWritable_(L);
    var sh = L.sheet;
    var row = Number(p.row);
    if (!(row >= L.start && row < L.totalRow)) throw new Error('Cette ligne ne peut pas être modifiée ici.');
    var ancienLabel = String(sh.getRange(row, L.col).getValue());
    if (norm_(ancienLabel) !== norm_(p.label)) throw new Error('Le tableau a bougé, recharge la page.');
    var brut = String(p.montant === undefined ? '' : p.montant).replace(/\s/g, '').replace(',', '.');
    if (brut === '' || isNaN(Number(brut))) throw new Error('Montant invalide.');
    var estMode = MODES.indexOf(norm_(ancienLabel).replace(/s$/, '')) >= 0;
    var montant = Math.abs(Number(brut));
    if (!estMode && p.sens === 'depense') montant = -montant;
    var cell = sh.getRange(row, L.col + 1);
    var avant = num_(cell.getValue());
    var libelle = estMode ? ancienLabel : String(p.libelle || '').trim();
    if (!libelle) throw new Error('Libellé obligatoire.');
    if (!estMode && libelle !== ancienLabel) sh.getRange(row, L.col).setValue(safeCell_(libelle));
    if (Math.abs(avant - montant) > 0.001) {
      cell.setValue(montant);
      var note = cell.getNote();
      cell.setNote((note ? note + '\n' : '') + fmtDate_(new Date()) + ' : ' + avant + ' € -> ' + montant + ' € (modifié par ' + currentUserLabel_() + ')');
    }
    var d = parseDate_(p.date);
    if (d) sh.getRange(row, L.col + 2).setValue(d).setNumberFormat('dd/mm/yyyy');
    SpreadsheetApp.flush();
    fixTotal_(comptaLocate_());
    logAction_('ligne compta modifiée', ancienLabel + ' ' + avant + ' -> ' + libelle + ' ' + montant);
    return true;
  } finally { lock.releaseLock(); }
}

/** Supprime une ligne de dépense (jamais les lignes Cash / Chèque / Virement / Total) */
function apiSupprimerLigneCompta_(p) {
  var lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    var L = comptaLocate_();
    assertWritable_(L);
    var row = Number(p.row);
    var label = String(L.sheet.getRange(row, L.col).getValue());
    if (norm_(label) !== norm_(p.label)) throw new Error('Le tableau a bougé, recharge la page.');
    if (row < L.start || row >= L.cashRow) throw new Error('Cette ligne ne peut pas être supprimée ici.');
    L.sheet.getRange(row, L.col, 1, 4).deleteCells(SpreadsheetApp.Dimension.ROWS);
    SpreadsheetApp.flush();
    fixTotal_(comptaLocate_());
    logAction_('suppression ligne compta', label);
    return true;
  } finally { lock.releaseLock(); }
}

/** Ajoute (ou retire si négatif) une recette sur un mode de paiement */
function apiAjouterRecette_(p) {
  var lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    var mode = norm_(p.mode).replace(/s$/, '');
    if (mode === 'especes' || mode === 'espece' || mode === 'liquide') mode = 'cash';
    if (MODES.indexOf(mode) < 0) throw new Error('Mode inconnu : Cash, Chèque ou Virement.');
    var montant = num_(p.montant);
    if (!montant) throw new Error('Montant obligatoire.');
    var L = comptaLocate_();
    assertWritable_(L);
    var row = L.modeRows[mode];
    if (!row) throw new Error('Ligne "' + p.mode + '" introuvable dans le tableau.');
    var cell = L.sheet.getRange(row, L.col + 1);
    var f = cell.getFormula();
    var base = f ? f.replace(/^=/, '') : String(num_(cell.getValue()));
    var add = montant >= 0 ? '+' + montant : String(montant);
    cell.setFormula('=' + base + add);
    var d = parseDate_(p.date) || new Date();
    var note = cell.getNote();
    var line = fmtDate_(d) + ' ' + (montant >= 0 ? '+' : '') + montant + ' € ' + (p.libelle || '');
    cell.setNote((note ? note + '\n' : '') + line);
    L.sheet.getRange(row, L.col + 2).setValue(d).setNumberFormat('dd/mm/yyyy');
    SpreadsheetApp.flush();
    fixTotal_(comptaLocate_());
    logAction_('recette', line + ' (' + mode + ')');
    return true;
  } finally { lock.releaseLock(); }
}

/**
 * Total de la caisse corrigé à la main (l'info vient souvent de l'oral : "il y a 2 350 € en caisse").
 * On ne casse pas le =SUM : on met l'écart sur une ligne unique « Ajustement caisse » (au-dessus de Cash),
 * dont la cellule garde l'historique (=-120+45) et une note datée à chaque correction.
 */
function apiModifierCaisse_(p) {
  var lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    var voulu = num_(p.total);
    if (p.total === '' || p.total === null || p.total === undefined || isNaN(voulu)) throw new Error('Indique le nouveau total de la caisse.');
    var L = comptaLocate_();
    assertWritable_(L);
    var sh = L.sheet;
    var actuel = num_(sh.getRange(L.totalRow, L.col + 1).getValue());
    var ecart = Math.round((voulu - actuel) * 100) / 100;
    if (Math.abs(ecart) < 0.005) return { total: actuel, ecart: 0 };
    var d = new Date();
    var ligne = fmtDate_(d) + ' : total ' + actuel + ' € -> ' + voulu + ' € (' + (ecart > 0 ? '+' : '') + ecart + ' €)' + (p.raison ? ' ' + String(p.raison).slice(0, 120) : '') + ' par ' + currentUserLabel_();
    // ligne d'ajustement existante ?
    var n = L.cashRow - L.start;
    var labels = n > 0 ? sh.getRange(L.start, L.col, n, 1).getValues() : [];
    var row = -1;
    for (var i = 0; i < labels.length; i++) if (norm_(labels[i][0]) === 'ajustement caisse') { row = L.start + i; break; }
    if (row < 0) {
      row = L.cashRow;
      sh.getRange(row, L.col, 1, 4).insertCells(SpreadsheetApp.Dimension.ROWS);
      sh.getRange(row, L.col, 1, 3).setValues([['Ajustement caisse', ecart, d]]);
      sh.getRange(row, L.col + 1).setNumberFormat(sh.getRange(L.start, L.col + 1).getNumberFormat() || '#,##0 €');
    } else {
      var cell = sh.getRange(row, L.col + 1);
      var f = cell.getFormula();
      var base = f ? f.replace(/^=/, '') : String(num_(cell.getValue()));
      cell.setFormula('=' + base + (ecart >= 0 ? '+' + ecart : String(ecart)));
      sh.getRange(row, L.col + 2).setValue(d);
    }
    sh.getRange(row, L.col + 2).setNumberFormat('dd/mm/yyyy');
    var c2 = sh.getRange(row, L.col + 1);
    var note = c2.getNote();
    c2.setNote((note ? note + '\n' : '') + ligne);
    SpreadsheetApp.flush();
    fixTotal_(comptaLocate_(), row);
    logAction_('caisse', ligne);
    return { total: voulu, ecart: ecart };
  } finally { lock.releaseLock(); }
}

/** Upload d'un justificatif (photo / PDF) : renvoie le lien Drive */
/** Sous-dossier des justificatifs, créé au besoin. '' = racine des justificatifs. */
function dossierJustif_(nom) {
  var c = getConfig_();
  if (!c.ID_DOSSIER_JUSTIFICATIFS) throw new Error('Dossier justificatifs non configuré (relance setup).');
  var racine = DriveApp.getFolderById(c.ID_DOSSIER_JUSTIFICATIFS);
  if (!nom) return racine;
  var it = racine.getFoldersByName(nom);
  return it.hasNext() ? it.next() : racine.createFolder(nom);
}

var SOUS_DOSSIERS = { fournisseurs: 'Factures fournisseurs', intervenants: 'Factures intervenants', caisse: 'Justificatifs caisse' };

/**
 * Dépôt d'un ou plusieurs justificatifs.
 * p = { name, mime, data } pour un fichier, ou p = { fichiers: [{name, mime, data}], dossier: 'fournisseurs' }.
 * Renvoie { url, id, fichiers: [{url, id, nom}] } : url reste le premier fichier, pour l'existant.
 */
function apiUploadJustificatif_(p) {
  var liste = p.fichiers && p.fichiers.length ? p.fichiers : [{ name: p.name, mime: p.mime, data: p.data }];
  if (liste.length > 10) throw new Error('10 fichiers maximum à la fois.');
  var folder = dossierJustif_(SOUS_DOSSIERS[p.dossier] || '');
  var jour = Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd');
  var out = liste.map(function (f) {
    var bytes = Utilities.base64Decode(f.data);
    if (bytes.length > 15 * 1024 * 1024) throw new Error('Fichier trop lourd (15 Mo max) : ' + (f.name || ''));
    var name = jour + ' - ' + String(f.name || 'justificatif').replace(/[\\/:*?"<>|]/g, '_');
    var file = folder.createFile(Utilities.newBlob(bytes, f.mime || 'application/octet-stream', name));
    return { url: file.getUrl(), id: file.getId(), nom: name };
  });
  logAction_('justificatif', out.length + ' fichier(s) dans ' + folder.getName());
  return { url: out[0].url, id: out[0].id, fichiers: out };
}

// ---------- Factures à payer (intervenants, auto-entreprise, fournisseurs) ----------

/** Onglet Achats du classeur ERP (créé au besoin) */
function achatsSheet_() {
  var ss = erpSS_();
  var sh = ss.getSheetByName('Achats');
  if (!sh) { sh = ss.insertSheet('Achats'); sh.getRange(1, 1, 1, TABLES.Achats.length).setValues([TABLES.Achats]); }
  return sh;
}

/** Liste des personnes / sociétés qu'on paie (Config INTERVENANTS) */
function beneficiaires_() {
  var c = getConfig_();
  return String(c.INTERVENANTS || '').split(/[;\n]/).map(function (s) { return s.trim(); }).filter(String);
}

function apiAchats_() {
  achatsSheet_();
  var auj = new Date(); auj.setHours(0, 0, 0, 0);
  var liste = readTable_('Achats').map(function (r) {
    r.montant = num_(r.Montant);
    r.paye = /^pay/i.test(String(r.Statut || ''));
    var ech = parseDate_(r.Echeance);
    r.retard = !r.paye && !!ech && ech < auj;
    r.jours = ech ? Math.round((ech - auj) / 86400000) : null;
    r.pieces = [r.Lien].concat(String(r.Pieces || '').split('|')).map(function (x) { return String(x || '').trim(); })
      .filter(function (x) { return /^https:\/\//.test(x); });
    return r;
  }).sort(function (a, b) {
    if (a.paye !== b.paye) return a.paye ? 1 : -1;
    var da = parseDate_(a.Echeance) || parseDate_(a.Date), db = parseDate_(b.Echeance) || parseDate_(b.Date);
    return (da ? da.getTime() : 0) - (db ? db.getTime() : 0);
  });
  var aPayer = liste.filter(function (r) { return !r.paye; });
  return {
    liste: liste,
    beneficiaires: beneficiaires_(),
    aPayer: aPayer.reduce(function (a, r) { return a + Math.abs(r.montant); }, 0),
    nbAPayer: aPayer.length,
    retards: aPayer.filter(function (r) { return r.retard; }).length,
    payeTotal: liste.filter(function (r) { return r.paye; }).reduce(function (a, r) { return a + Math.abs(r.montant); }, 0)
  };
}

/** Enregistre (ou modifie) une facture à payer */
function apiSaveAchat_(p) {
  achatsSheet_();
  var montant = Math.abs(num_(p.Montant));
  if (!String(p.Beneficiaire || '').trim()) throw new Error('Choisis qui doit être payé.');
  if (!montant) throw new Error('Montant obligatoire.');
  var rec = {
    ID: p.ID || uid_(),
    Date: parseDate_(p.Date) || new Date(),
    Beneficiaire: String(p.Beneficiaire).trim(),
    Type: p.Type || 'Intervenant',
    Objet: p.Objet || '',
    Montant: montant,
    Echeance: parseDate_(p.Echeance) || '',
    Statut: p.Statut || 'À payer',
    PayeLe: parseDate_(p.PayeLe) || '',
    Mode: p.Mode || '',
    Lien: /^https:\/\//.test(String(p.Lien || '')) ? p.Lien : '',
    // pièces jointes supplémentaires, séparées par « | » (une facture fournisseur peut tenir en plusieurs pages)
    Pieces: String(p.Pieces || '').split('|').map(function (x) { return x.trim(); })
      .filter(function (x) { return /^https:\/\//.test(x); }).join(' | '),
    Notes: p.Notes || '',
    MajPar: currentUserLabel_(), MajLe: new Date()
  };
  upsertRow_('Achats', rec, 'ID');
  logAction_('facture à payer', rec.Beneficiaire + ' ' + montant + ' € (' + rec.Statut + ')');
  return rec.ID;
}

/**
 * Marque une facture payée. p = { id, mode, date, caisse: true/false }
 * caisse = true : la dépense est aussi écrite dans le tableau de caisse (Compte aussi).
 */
function apiPayerAchat_(p) {
  achatsSheet_();
  var a = readTable_('Achats').filter(function (r) { return String(r.ID) === String(p.id); })[0];
  if (!a) throw new Error('Facture introuvable.');
  if (/^pay/i.test(String(a.Statut || ''))) throw new Error('Cette facture est déjà notée payée.');
  var d = parseDate_(p.date) || new Date();
  var montant = Math.abs(num_(a.Montant));
  var libelle = String(a.Beneficiaire) + (a.Objet ? ' · ' + a.Objet : '');
  if (p.caisse) apiAjouterDepense_({ libelle: libelle, montant: montant, date: d, justificatif: a.Lien || '' });
  upsertRow_('Achats', { ID: a.ID, Statut: 'Payée', PayeLe: d, Mode: p.mode || a.Mode || '', MajPar: currentUserLabel_(), MajLe: new Date() }, 'ID');
  logAction_('paiement', libelle + ' -' + montant + (p.caisse ? ' (caisse)' : ' (hors caisse)'));
  return { montant: montant, caisse: !!p.caisse };
}

function apiSupprimerAchat_(p) {
  achatsSheet_();
  if (!deleteRowById_('Achats', p.id)) throw new Error('Facture introuvable.');
  logAction_('facture à payer supprimée', String(p.id));
  return true;
}

/* ===== Factures.gs ===== */
/**
 * Factures.gs : génère un DEVIS ou une FACTURE à partir de Modele_facture
 * (copie du Doc + remplissage + PDF), suivis dans l'onglet Factures.
 * Nomenclature (comme DEV-2026-001_Villeneuve.pdf) :
 *   DEV-2026-002_<Client>.pdf  /  FAC-2026-001_<Client>.pdf  à la racine du dossier Devis / Facture,
 *   le Google Doc modifiable dans le sous-dossier « Modifiables (Docs) ».
 */

function eur_(n) {
  var v = Math.round(num_(n) * 100) / 100;
  var s = v.toFixed(2).split('.');
  var ent = s[0].replace(/\B(?=(\d{3})+(?!\d))/g, '\u202f');
  return ent + ',' + s[1] + ' €';
}

function apiFactures_() {
  return readTable_('Factures').map(function (f) {
    f.Type = f.Type || (/^DEV/i.test(String(f.Numero)) ? 'Devis' : 'Facture');
    return f;
  }).sort(function (a, b) {
    var da = parseDate_(a.Date), db = parseDate_(b.Date);
    return ((db ? db.getTime() : 0) - (da ? da.getTime() : 0)) || String(b.Numero).localeCompare(String(a.Numero));
  });
}

function nextNumero_(n) {
  var m = String(n).match(/^(.*?)(\d+)$/);
  if (!m) return n + '-1';
  var next = String(+m[2] + 1);
  while (next.length < m[2].length) next = '0' + next;
  return m[1] + next;
}

function dossiersFactures_() {
  var c = getConfig_();
  var root = c.ID_DOSSIER_DEVIS ? DriveApp.getFolderById(c.ID_DOSSIER_DEVIS) : DriveApp.getRootFolder();
  return { root: root, docs: getOrCreateFolder_(root, 'Modifiables (Docs)') };
}

/** Numéro du prochain devis / de la prochaine facture (DEV-2026-002, FAC-2026-001) */
function numeroDoc_(type, c) {
  var an = Utilities.formatDate(new Date(), TZ, 'yyyy');
  if (type === 'devis') return String(c.PROCHAIN_NUM_DEVIS || ('DEV-' + an + '-001')).trim();
  var n = String(c.PROCHAIN_NUM_FACTURE || ('FAC-' + an + '-001')).trim();
  if (/^\d{4}-\d+$/.test(n)) n = 'FAC-' + n; // ancien format 2026-001
  return n;
}

/** "Mairie de Villeneuve" -> "Villeneuve" (pour le nom de fichier) */
function nomCourt_(client) {
  var t = String(client || '').trim()
    .replace(/^(mairie|commune|ville|ecole|école|groupe scolaire|college|collège|centre de loisirs|accueil de loisirs|association|communauté de communes)\s+(de\s+la\s+|de\s+l'|du\s+|des\s+|de\s+|d')?/i, '');
  t = t.replace(/[\\/:*?"<>|#%]/g, '').replace(/['’]/g, '-').replace(/\s+/g, '-').replace(/-+/g, '-').replace(/^-|-$/g, '');
  return (t || 'Client').slice(0, 40);
}

function ensureColumns_(name) {
  var sh = erpSheet_(name);
  var h = headersOf_(name);
  (TABLES[name] || []).forEach(function (col) {
    if (h.indexOf(col) < 0) { sh.getRange(1, h.length + 1).setValue(col); h.push(col); }
  });
}

/**
 * p = { type: 'devis' | 'facture', client, service, adresse, cpVille, siret, engagement, periode,
 *       date (yyyy-mm-dd), conditions?, depuisDevis? (ID), lignes: [{designation, detail, qte, pu}] }
 */
function apiCreerFacture_(p) {
  var lock = LockService.getScriptLock(); lock.waitLock(30000);
  try {
    var c = getConfig_(true);
    var type = p.type === 'devis' ? 'devis' : 'facture';
    if (!p.client) throw new Error('Client obligatoire.');
    var lignes = (p.lignes || []).filter(function (l) { return l.designation && num_(l.qte) && num_(l.pu); });
    if (!lignes.length) throw new Error('Au moins une ligne avec désignation, quantité et prix.');
    var numero = numeroDoc_(type, c);
    var dateF = parseDate_(p.date) || new Date();
    var jours = type === 'devis' ? (num_(c.VALIDITE_DEVIS_JOURS) || 30) : (num_(c.ECHEANCE_FACTURE_JOURS) || 30);
    var ech = new Date(dateF.getTime() + jours * 86400000);
    var total = lignes.reduce(function (a, l) { return a + num_(l.qte) * num_(l.pu); }, 0);
    var court = p.nomCourt ? nomCourt_(p.nomCourt) : nomCourt_(p.client);
    var nom = numero + '_' + court;

    var f = dossiersFactures_();
    var copy = DriveApp.getFileById(c.ID_MODELE_FACTURE).makeCopy(nom, f.docs);
    var doc = DocumentApp.openById(copy.getId());
    var body = doc.getBody();

    function rep(pattern, value) { body.replaceText(pattern, String(value || '').replace(/\$/g, '\\$')); }
    rep('N° ?\\d{4}-\\d{3}', 'N° ' + numero);
    rep('\\[Nom de la collectivité\\]', p.client);
    rep('\\[Service\\]', p.service || '');
    rep('\\[Adresse\\]', p.adresse || '');
    rep('\\[Code postal et ville\\]', p.cpVille || '');
    rep('\\[SIRET du client\\]', p.siret || '-');
    rep('IBAN \\[à compléter\\]', 'IBAN ' + (c.IBAN || ''));
    rep('BIC \\[à compléter\\]', 'BIC ' + (c.BIC || ''));
    fillTable_(body, lignes, total);

    if (type === 'facture') {
      rep('Émise le \\[JJ/MM/AAAA\\]', 'Émise le ' + fmtDate_(dateF));
      rep('Échéance le \\[JJ/MM/AAAA\\]', 'Échéance le ' + fmtDate_(ech));
      rep('\\[bon de commande\\]', p.engagement || '-');
      rep('\\[période\\]', p.periode || '');
    } else {
      rep('^\\s*FACTURE\\s*$', 'DEVIS');
      rep('Émise le \\[JJ/MM/AAAA\\]', 'Établi le ' + fmtDate_(dateF));
      rep('Échéance le \\[JJ/MM/AAAA\\]', 'Valable ' + jours + ' jours');
      rep('FACTURÉ À', 'DEVIS ÉTABLI POUR');
      rep("N° d'engagement \\[bon de commande\\]", p.engagement ? "N° d'engagement " + p.engagement : '');
      rep('Prestation réalisée \\[période\\]', p.periode ? 'Prestation : ' + p.periode : '');
      rep('\\[période\\]', p.periode || '');
      rep('TOTAL À PAYER', 'TOTAL DU DEVIS');
      blocsDevis_(body, c, p.conditions);
    }
    doc.saveAndClose();

    var pdf = f.root.createFile(copy.getAs('application/pdf').setName(nom + '.pdf'));
    ensureColumns_('Factures');
    var details = { service: p.service || '', adresse: p.adresse || '', cpVille: p.cpVille || '', siret: p.siret || '', engagement: p.engagement || '',
      periode: p.periode || '', nomCourt: court, lignes: lignes, depuisDevis: p.depuisDevis || '' };
    var rec = {
      ID: uid_(), Numero: numero, Date: dateF, Client: p.client, Objet: lignes.map(function (l) { return l.designation; }).join(' / ').slice(0, 200),
      Montant: total, Echeance: ech, Statut: type === 'devis' ? 'Envoyé' : 'Émise', PayeeLe: '', LienDoc: copy.getUrl(), LienPdf: pdf.getUrl(),
      Type: type === 'devis' ? 'Devis' : 'Facture', Details: JSON.stringify(details)
    };
    upsertRow_('Factures', rec, 'ID');
    setConfig_(type === 'devis' ? 'PROCHAIN_NUM_DEVIS' : 'PROCHAIN_NUM_FACTURE', nextNumero_(numero));
    if (p.depuisDevis) upsertRow_('Factures', { ID: p.depuisDevis, Statut: 'Facturé' }, 'ID');
    logAction_(type, numero + ' ' + p.client + ' ' + total);
    return { numero: numero, type: type, doc: copy.getUrl(), pdf: pdf.getUrl(), total: total, fichier: nom + '.pdf' };
  } finally { lock.releaseLock(); }
}

/** Devis : le cadre RÈGLEMENT devient CONDITIONS, MENTIONS LÉGALES devient BON POUR ACCORD */
function blocsDevis_(body, c, conditions) {
  var tables = body.getTables();
  for (var i = 0; i < tables.length; i++) {
    var t = tables[i];
    if (!t.getNumRows() || t.getRow(0).getNumCells() < 2) continue;
    if (norm_(t.getRow(0).getCell(0).getText()).indexOf('reglement') !== 0) continue;
    var cond = conditions || ('Devis valable ' + (num_(c.VALIDITE_DEVIS_JOURS) || 30) + ' jours. Lieu : sur place, ' + c.ADRESSE_STRUCTURE +
      '. Facturation à l\'issue de la prestation, règlement par virement (Chorus Pro pour les clients publics). Toute séance annulée moins de 48 heures à l\'avance est due.');
    setCellParas_(t.getRow(0).getCell(0), [{ t: 'CONDITIONS', b: true }, { t: cond }, { t: 'TVA non applicable, article 261-7-1° du Code général des impôts.', i: true }]);
    setCellParas_(t.getRow(0).getCell(1), [{ t: 'BON POUR ACCORD', b: true }, { t: 'Date, signature et cachet, précédés de la mention « Bon pour accord »' }, { t: ' ' }, { t: ' ' },
      { t: 'Pour L\'Ain\'stant Activite : ' + (c.SIGNATAIRE || '') }]);
    return;
  }
}

function setCellParas_(cell, paras) {
  for (var i = cell.getNumChildren() - 1; i >= 1; i--) cell.getChild(i).removeFromParent();
  var first = cell.getChild(0).asParagraph();
  paras.forEach(function (x, k) {
    var par = k === 0 ? first : cell.appendParagraph('');
    par.setText(x.t || ' ');
    var tx = par.editAsText();
    tx.setBold(!!x.b); tx.setItalic(!!x.i);
  });
}

function fillTable_(body, lignes, total) {
  var tables = body.getTables();
  var t = null;
  for (var i = 0; i < tables.length; i++) {
    if (norm_(tables[i].getRow(0).getText()).indexOf('designation') >= 0) { t = tables[i]; break; }
  }
  if (!t) throw new Error('Tableau DÉSIGNATION introuvable dans le modèle de facture.');
  // lignes "article" = entre l'en-tête et la ligne Sous-total
  var subRow = -1;
  for (var r = 1; r < t.getNumRows(); r++) if (norm_(t.getRow(r).getText()).indexOf('sous-total') >= 0) { subRow = r; break; }
  if (subRow < 0) throw new Error('Ligne Sous-total introuvable dans le modèle.');
  var slots = subRow - 1;
  // ajoute des lignes si besoin (copie de la dernière ligne vide)
  while (slots < lignes.length) { t.insertTableRow(subRow, t.getRow(subRow - 1).copy()); subRow++; slots++; }
  for (var k = 0; k < slots; k++) {
    var row = t.getRow(1 + k);
    var l = lignes[k];
    setCell_(row.getCell(0), l ? l.designation : '', l ? l.detail : '');
    setCell_(row.getCell(1), l ? String(num_(l.qte)) : '');
    setCell_(row.getCell(2), l ? eur_(l.pu) : '');
    setCell_(row.getCell(3), l ? eur_(num_(l.qte) * num_(l.pu)) : '');
  }
  // Sous-total et TOTAL : le montant est dans la cellule qui suit le libellé
  for (var r2 = subRow; r2 < t.getNumRows(); r2++) {
    var rw = t.getRow(r2);
    for (var cI = 0; cI < rw.getNumCells() - 1; cI++) {
      var tx = norm_(rw.getCell(cI).getText());
      if (tx.indexOf('sous-total') >= 0 || tx.indexOf('total a payer') >= 0 || tx.indexOf('total du devis') >= 0) {
        var cell = rw.getCell(cI + 1);
        var bold = tx.indexOf('total a payer') >= 0;
        setCell_(cell, eur_(total));
        if (bold) cell.editAsText().setBold(true);
        break;
      }
    }
  }
}

function setCell_(cell, main, detail) {
  var n = cell.getNumChildren();
  var p0 = cell.getChild(0).asParagraph();
  var wasBold = p0.getText().length ? p0.editAsText().isBold(0) : false;
  for (var i = n - 1; i >= 1; i--) cell.getChild(i).removeFromParent();
  p0.setText(main || ' ');
  if (wasBold && main) p0.editAsText().setBold(0, main.length - 1, true);
  if (detail) {
    var p1 = cell.appendParagraph(detail);
    p1.editAsText().setBold(false).setFontSize(9);
  }
}

/** PDF présents dans le dossier Devis / Facture du Drive (y compris ceux faits à la main) */
function apiDossierDevis_() {
  var c = getConfig_();
  if (!c.ID_DOSSIER_DEVIS) return [];
  var suivis = {};
  readTable_('Factures').forEach(function (f) { suivis[norm_(f.Numero)] = f.Statut || ''; });
  var out = [];
  try {
    var folder = DriveApp.getFolderById(c.ID_DOSSIER_DEVIS);
    var it = folder.getFiles();
    while (it.hasNext() && out.length < 60) {
      var f = it.next();
      var nom = f.getName();
      if (!/\.pdf$/i.test(nom)) continue;
      var num = (nom.match(/^(DEV|FAC)-\d{4}-\d+/i) || [''])[0];
      var type = /^DEV/i.test(num) || /devis/i.test(nom) ? 'Devis' : 'Facture';
      out.push({ nom: nom, numero: num, type: type, url: f.getUrl(), date: fmtDate_(f.getLastUpdated()),
        modif: f.getLastUpdated().getTime(), statut: num ? (suivis[norm_(num)] || '') : '' });
    }
  } catch (e) { return []; }
  out.sort(function (a, b) { return b.modif - a.modif; });
  return out;
}

/** Changement de statut ; si payée et option, ajoute la recette en Virement dans la caisse */
function apiStatutFacture_(p) {
  var f = readTable_('Factures').filter(function (x) { return String(x.ID) === String(p.id); })[0];
  if (!f) throw new Error('Facture introuvable.');
  var upd = { ID: p.id, Statut: p.statut };
  if (p.statut === 'Payée') upd.PayeeLe = new Date();
  upsertRow_('Factures', upd, 'ID');
  if (p.statut === 'Payée' && p.ajouterCaisse) {
    apiAjouterRecette_({ mode: p.mode || 'Virement', montant: num_(f.Montant), libelle: 'Facture ' + f.Numero + ' ' + f.Client });
  }
  logAction_('facture statut', f.Numero + ' -> ' + p.statut);
  return true;
}

/* ===== Mails.gs ===== */
/**
 * Mails.gs : brouillons Gmail (rien ne part sans validation humaine),
 * déclencheurs d'inscription, tâche quotidienne (relances optionnelles + récap équipe).
 */

/**
 * Signature commune à tous les mails sortants : même formule, même adresse,
 * téléphone + mail + site. Tout vient de l'onglet Config.
 */
function signature_() {
  var c = getConfig_();
  var adresse = String(c.ADRESSE_STRUCTURE || '').trim();
  var contact = [c.TEL_STRUCTURE, c.EMAIL_STRUCTURE, c.SITE_WEB].filter(String).join(' · ');
  return '\n\nMerci à vous, et à très vite ' + L_('lieu') + '.\n\nL\'équipe de ' + c.NOM_STRUCTURE + '\n' + adresse + '\n' + contact;
}

function infosPaiement_(mode) {
  var c = getConfig_();
  var m = norm_(mode);
  var lines = [];
  if (!m || m.indexOf('virement') >= 0) {
    lines.push('Par virement : titulaire ' + c.IBAN_TITULAIRE + ', IBAN ' + c.IBAN + ', BIC ' + c.BIC + ' (merci d\'indiquer le nom ' + Ldu_('membre') + ' en référence).');
  }
  if (!m || m.indexOf('cheque') >= 0) lines.push('Par chèque : à l\'ordre de ' + (c.IBAN_TITULAIRE || c.NOM_ASSO || c.NOM_STRUCTURE) + ', à remettre ' + L_('lieu') + '.');
  if (!m || m.indexOf('espece') >= 0 || m.indexOf('cash') >= 0) lines.push('En espèces : directement ' + L_('lieu') + ', un reçu vous sera remis.');
  return lines.join('\n');
}

function tarifsEleve_(e) {
  if (e.source === 'cours') {
    var m = matchCours_(e.creneaux, coursList_());
    return m.map(function (k) { return { label: k.libelle + ', ' + k.jour + ' ' + k.horaire, prix: k.tarif }; });
  }
  var s = matchStages_(e.stage, stagesList_());
  return s.map(function (k) { return { label: 'Stage ' + k.nom + ' (' + k.dates + '), ' + k.public, prix: k.tarif }; });
}

function corpsConfirmation_(e) {
  var c = getConfig_();
  var t = tarifsEleve_(e);
  var total = t.reduce(function (a, x) { return a + x.prix; }, 0);
  var bonjour = 'Bonjour' + (e.responsable ? ' ' + String(e.responsable).trim() : '') + ',';
  var qui = e.prenom + ' ' + e.nom;
  var lignes = t.length ? t.map(function (x) { return '  • ' + x.label + ' : ' + x.prix + ' €'; }).join('\n') : '  • ' + (e.creneaux || e.stage);
  var txt = bonjour + '\n\nMerci, nous avons bien reçu l\'inscription de ' + qui + ' à ' + c.NOM_STRUCTURE + '.\n\n' +
    'Récapitulatif :\n' + lignes + '\n';
  if (t.length > 1) txt += '  Total : ' + total + ' € (réduction fratrie éventuelle appliquée à la validation)\n';
  if (e.source === 'cours') {
    txt += '\nLe tarif annuel comprend l\'adhésion à l\'association.';
    if (norm_(e.essai).indexOf('oui') === 0) txt += ' Le premier cours sert de cours d\'essai gratuit : l\'inscription est confirmée après ce cours.';
    txt += '\n';
  } else {
    txt += '\nL\'adhésion à l\'association (' + c.ADHESION + ' € par an et par enfant) s\'ajoute au prix du stage pour les nouveaux inscrits.\n';
  }
  txt += '\nRèglement :\n' + infosPaiement_(e.modePaiement) + '\n';
  if (e.manquants && e.manquants.length) {
    txt += '\nIl nous manque encore : ' + e.manquants.join(', ') + '. Vous pouvez simplement répondre à ce mail avec ces informations.\n';
  }
  txt += '\nRendez-vous ' + c.ADRESSE_STRUCTURE + '. Prévoir une tenue souple et une gourde.';
  txt += signature_();
  var sujet = (e.source === 'cours' ? 'Inscription confirmée' : 'Inscription au stage confirmée') + ' : ' + e.prenom + ' · ' + c.NOM_STRUCTURE;
  return { sujet: sujet, corps: txt };
}

/**
 * Une inscription offerte, un tarif « à la carte », une réduction fratrie ou commune :
 * dans ces cas le tarif du créneau n'est PAS le montant dû. On ne met alors aucun chiffre
 * dans le mail, plutôt que d'en annoncer un faux. C'est la règle centrale de tout ce
 * fichier : un montant incertain ne sort jamais vers une famille.
 *
 * REDUC_A_VERIFIER se complète depuis l'onglet Config (clé REDUC_MOTIFS) pour ajouter
 * les motifs propres à la structure, par exemple le nom d'une commune conventionnée.
 */
var REDUC_A_VERIFIER = ['fratrie', 'demi ?annee', 'demi-annee', 'a la carte', 'preferentiel', 'convention', 'tarif reduit', 'quotient'];

function motifsReduc_() {
  var sup = String(getConfig_().REDUC_MOTIFS || '').split(/[;,]/)
    .map(function (s) { return norm_(s).trim(); }).filter(Boolean);
  return new RegExp(REDUC_A_VERIFIER.concat(sup.map(escapeRegex_)).join('|'));
}

function reducEleve_(e) {
  var r = norm_(e.reductions) + ' ' + norm_(e.notes) + ' ' + norm_(e.commentaire);
  if (/offert|gratuit|invit/.test(r)) return { rien: true, libelle: 'offert' };
  var m = r.match(motifsReduc_());
  return m ? { aVerifier: true, libelle: String(e.reductions || e.notes || '').trim() } : {};
}

/** Montant à réclamer, ou null quand il demande une vérification humaine */
function montantDu_(e) {
  var red = reducEleve_(e);
  if (red.rien || red.aVerifier) return null;
  var total = tarifsEleve_(e).reduce(function (a, x) { return a + x.prix; }, 0);
  if (!total) return null;
  var recu = num_(e.montantRecu);
  return { total: total, recu: recu, reste: Math.max(total - recu, 0) };
}

/**
 * Relance de PAIEMENT. Elle ne parle que du règlement : pas de demande de complément
 * d'informations, qui est un autre sujet et un autre message.
 * eleves = un ou plusieurs enfants de la même famille, pour n'envoyer qu'un mail par foyer.
 */
function corpsRelance_(eleves) {
  var c = getConfig_();
  var liste = [].concat(eleves);
  var e = liste[0];
  var bonjour = 'Bonjour' + (e.responsable ? ' ' + String(e.responsable).trim() : '') + ',';
  var lignes = [], total = 0, sur = true, modes = {};

  liste.forEach(function (x) {
    var m = montantDu_(x);
    var t = tarifsEleve_(x);
    var quoi = t.length ? t.map(function (y) { return y.label; }).join(' et ') : (x.creneaux || x.stage || '');
    if (m) { total += m.reste; lignes.push('  • ' + x.prenom + ' ' + x.nom + (quoi ? ' — ' + quoi : '') + ' : ' + m.reste + ' €' + (m.recu ? ' (reste à régler, ' + m.recu + ' € déjà reçus)' : '')); }
    else { sur = false; lignes.push('  • ' + x.prenom + ' ' + x.nom + (quoi ? ' — ' + quoi : '')); }
    if (x.modePaiement) modes[norm_(x.modePaiement)] = x.modePaiement;
  });

  var txt = bonjour + '\n\n';
  txt += liste.length > 1
    ? 'Nous revenons vers vous au sujet des inscriptions de vos enfants à ' + c.NOM_STRUCTURE + '.'
    : 'Nous revenons vers vous au sujet de l\'inscription de ' + e.prenom + ' ' + e.nom + ' à ' + c.NOM_STRUCTURE + '.';
  txt += '\n\nSauf erreur de notre part, le règlement ne nous est pas encore parvenu :\n' + lignes.join('\n');
  if (sur && liste.length > 1) txt += '\n  Total : ' + total + ' €';
  if (!sur) txt += '\n\nUn tarif particulier s\'applique à votre situation : nous vous confirmons le montant exact par retour de mail.';
  txt += '\n\n' + infosPaiement_(Object.keys(modes).length === 1 ? modes[Object.keys(modes)[0]] : '');
  // beaucoup de familles ont déjà convenu d'un échelonnement de vive voix : on le dit,
  // sinon elles reçoivent une relance qui a l'air de ne pas tenir compte de leur accord
  var suivi = String(c.CONTACT_PAIEMENTS || '').trim();
  txt += '\n\nSi le règlement a déjà été effectué' +
    (suivi ? ', ou si vous avez déjà convenu de modalités de paiement avec ' + suivi : '') +
    ', merci de ne pas tenir compte de ce message. Pour toute difficulté ou pour étaler le règlement en plusieurs fois, répondez simplement à ce mail, nous trouverons une solution.';
  txt += signature_();

  var sujet = liste.length > 1
    ? 'Règlement des inscriptions · ' + c.NOM_STRUCTURE
    : 'Règlement de l\'inscription de ' + e.prenom + ' · ' + c.NOM_STRUCTURE;
  return { sujet: sujet, corps: txt, montant: sur ? total : null };
}

// ---------- mise en page officielle des mails (bandeau + couleurs de la structure) ----------

function logoBlob_() {
  var d = logo_(getConfig_().ID_LOGO);
  var m = d && d.match(/^data:([^;]+);base64,(.*)$/);
  if (!m) return null;
  return Utilities.newBlob(Utilities.base64Decode(m[2]), m[1], 'logo.png');
}

/** Texte brut du mail -> HTML aux couleurs de la structure (le texte brut reste la version de secours) */
function mailHtml_(corps, logoSrc) {
  var c = getConfig_();
  var acc = c.COULEUR_ACCENT || '#9b1017';
  var blocs = String(corps || '').replace(/\r/g, '').split(/\n{2,}/);
  var puce = /^\s*[•\-]\s+/;
  var html = blocs.map(function (b) {
    // un bloc = paragraphe ; les lignes à puces consécutives deviennent une vraie liste
    var out = '', para = [], liste = [];
    function flushPara() { if (para.length) { out += '<p style="margin:0 0 14px">' + para.join('<br>') + '</p>'; para = []; } }
    function flushListe() { if (liste.length) { out += '<ul style="margin:0 0 14px;padding-left:20px">' + liste.join('') + '</ul>'; liste = []; } }
    b.split('\n').forEach(function (l, i) {
      if (puce.test(l)) { flushPara(); liste.push('<li style="margin:3px 0">' + esc_(l.replace(puce, '')) + '</li>'); return; }
      flushListe();
      var t = esc_(l);
      // titre de section : une ligne courte tout en majuscules, comme dans les envois groupés
      if (i === 0 && l.length < 70 && l.trim() && l === l.toUpperCase() && /[A-ZÀ-Ý]/.test(l)) {
        flushPara();
        out += '<div style="margin:22px 0 10px;color:' + acc + ';font-size:13px;font-weight:bold;letter-spacing:.6px">' + t + '</div>';
        return;
      }
      if (i === 0 && /:\s*$/.test(l) && l.length < 60) t = '<b style="color:' + acc + '">' + t + '</b>';
      para.push(t);
    });
    flushPara(); flushListe();
    return out;
  }).join('');
  // Bandeau : titre en serif, baseline et intertitres dans la couleur d'accent,
  // filets dans la couleur secondaire, corps en sans-serif. AUCUNE image : pas de
  // pièce jointe, pas de « télécharger les images » chez le destinataire, et le
  // même rendu dans Gmail, Outlook et sur mobile.
  //
  // MOT_ACCENTUE permet de teinter un mot du nom (par exemple le mot central d'un
  // nom composé) sans toucher au code : il est cherché tel quel dans NOM_STRUCTURE.
  var nom = String(c.NOM_STRUCTURE || '');
  var accent = acc;
  var mot = String(c.MOT_ACCENTUE || '').trim();
  var titre;
  if (mot && nom.indexOf(mot) >= 0) {
    var i = nom.indexOf(mot);
    titre = span_(nom.slice(0, i), '#141414') + span_(mot, accent) + span_(nom.slice(i + mot.length), '#141414');
  } else {
    titre = span_(nom, '#141414');
  }
  var site = String(c.SITE_WEB || '').replace(/^https?:\/\//, '').replace(/\/$/, '');
  var pied = [c.NOM_STRUCTURE, c.NOM_ASSO, String(c.ADRESSE_STRUCTURE || '').trim()].filter(String).map(esc_).join(' · ');
  return '<div data-erp="1" style="font-family:Calibri,Arial,sans-serif;font-size:15px;color:#2b2b2b;line-height:1.55;max-width:640px">' +
    '<div style="text-align:center;padding:6px 0 4px 0">' + titre +
    (c.BASELINE ? '<div style="font-size:13px;font-weight:bold;color:' + accent + ';padding-top:4px">' + esc_(c.BASELINE) + '</div>' : '') +
    '</div>' +
    filet_(c, '8px 0 18px 0') +
    html +
    filet_(c, '18px 0 0 0') +
    '<p style="text-align:center;color:#6b6b6b;font-size:12px;margin:8px 0 0 0">' + pied +
    (site ? ' · <a href="https://' + esc_(site) + '" style="color:' + accent + '">' + esc_(site) + '</a>' : '') +
    '</p></div>';
}

var MAIL_TITRE = 'font-family:Georgia,serif;font-size:30px;font-weight:bold';

function span_(txt, couleur) {
  return txt ? '<span style="' + MAIL_TITRE + ';color:' + couleur + '">' + esc_(txt) + '</span>' : '';
}

function filet_(c, marge) {
  return '<table cellpadding="0" cellspacing="0" border="0" width="100%" style="border-collapse:collapse;margin:' + marge + '">' +
    '<tr><td style="background-color:' + (c.COULEUR_FILET || '#ffd400') + ';height:4px;line-height:4px;font-size:0">&nbsp;</td></tr></table>';
}

/**
 * Le gabarit n'a plus d'image : plus de pièce jointe logo, plus de « télécharger les images »
 * chez le destinataire, et le même rendu partout (Gmail, Outlook, mobile).
 */
function optionsMail_(corps) {
  return { name: getConfig_().NOM_STRUCTURE, htmlBody: mailHtml_(corps) };
}

/** Aperçu dans l'ERP : exactement le HTML qui part */
function htmlApercu_(corps) {
  return mailHtml_(corps);
}

function verifDest_(to) {
  var t = String(to || '').trim();
  if (!t) throw new Error('Pas d\'adresse email.');
  t.split(/\s*,\s*/).forEach(function (a) {
    if (!/^[^\s@<>"]+@[^\s@<>"]+\.[^\s@<>"]+$/.test(a)) throw new Error('Adresse email invalide : ' + a);
  });
  return t;
}

function creerBrouillon_(to, sujet, corps) {
  if (!to) throw new Error('Pas d\'adresse email pour cette inscription.');
  var d = GmailApp.createDraft(verifDest_(to), sujet, corps, optionsMail_(corps));
  return d.getId();
}

function noterEnvoi_(e, type) {
  var s = suiviIndex_()[e.id] || { ID: e.id, Source: e.source, Nom: e.nom, Prenom: e.prenom, Statut: 'Actif', Paye: 'Non' };
  delete s._row;
  if (type === 'relance') s.DerniereRelance = new Date(); else s.Confirmation = new Date();
  saveSuivi_(s);
}

/**
 * Un mail de relance peut couvrir plusieurs enfants d'un même foyer : la date de relance
 * doit alors être notée sur chacun, sinon les frères et sœurs ressortent dans la liste.
 */
function noterEnvoiTous_(p, e, type) {
  var l = (p && p.items || []).filter(function (x) { return x && x.id; });
  if (!l.length) return noterEnvoi_(e, type);
  l.forEach(function (x) {
    try {
      var y = (x.source === e.source && String(x.id) === String(e.id)) ? e : apiEleve_({ source: x.source, id: x.id });
      noterEnvoi_(y, type);
    } catch (err) { }
  });
}

/** Bouton "Créer le brouillon" (texte éventuellement retouché dans l'ERP) */
function apiBrouillon_(p) {
  var e = apiEleve_({ source: p.source, id: p.id });
  var m = p.type === 'relance' ? corpsRelance_(e) : corpsConfirmation_(e);
  var to = p.to || e.emailBest, sujet = p.sujet || m.sujet, corps = p.corps || m.corps;
  creerBrouillon_(to, sujet, corps);
  noterEnvoiTous_(p, e, p.type);
  logAction_('brouillon ' + p.type, e.prenom + ' ' + e.nom);
  return { to: to, sujet: sujet };
}

/** Bouton "Envoyer maintenant" (après aperçu dans l'ERP) */
function apiEnvoyerMail_(p) {
  var e = apiEleve_({ source: p.source, id: p.id });
  var m = p.type === 'relance' ? corpsRelance_(e) : corpsConfirmation_(e);
  var to = verifDest_(p.to || e.emailBest), sujet = p.sujet || m.sujet, corps = p.corps || m.corps;
  GmailApp.sendEmail(to, sujet, corps, optionsMail_(corps));
  noterEnvoiTous_(p, e, p.type);
  logAction_('mail envoyé ' + p.type, e.prenom + ' ' + e.nom + ' -> ' + to);
  return { to: to, sujet: sujet };
}

/** Aperçu du mail avant création / envoi */
function apiApercuMail_(p) {
  var e = apiEleve_({ source: p.source, id: p.id });
  var m = p.type === 'relance' ? corpsRelance_(e) : corpsConfirmation_(e);
  return { to: e.emailBest, sujet: m.sujet, corps: m.corps, html: htmlApercu_(m.corps) };
}

// ---------- brouillons Gmail gérés depuis l'ERP ----------

/**
 * Les brouillons préparés par Claude ont une étiquette dans l'objet ([À COMPLÉTER], [À RELIRE], [PRÊT])
 * pour dire s'il faut encore les retoucher. Elle est retirée avant l'envoi : le destinataire ne la voit jamais.
 */
var RE_ETIQUETTE = /^\s*\[\s*(a\s*completer|à\s*compléter|a\s*relire|à\s*relire|pret|prêt)[^\]]*\]\s*/i;
function etiquette_(sujet) {
  var m = String(sujet || '').match(RE_ETIQUETTE);
  return m ? m[0].trim() : '';
}
function sansEtiquette_(sujet) {
  return String(sujet || '').replace(RE_ETIQUETTE, '').trim();
}

function brouillon_(id) {
  var d = GmailApp.getDraft(String(id || ''));
  if (!d) throw new Error('Brouillon introuvable (déjà envoyé ou supprimé ?).');
  return d;
}

function apiBrouillonVoir_(p) {
  var d = brouillon_(p.id), m = d.getMessage();
  var sujet = m.getSubject() || '';
  var html = m.getBody() || '';
  var logo = '';
  try { logo = logo_(getConfig_().ID_LOGO); } catch (e) { logo = ''; }
  if (logo) html = html.replace(/(src=["'])cid:[^"']+(["'])/gi, '$1' + logo + '$2');
  var pj = 0;
  try { pj = m.getAttachments({ includeInlineImages: false }).length; } catch (e) { pj = 0; }
  return { id: p.id, a: m.getTo(), cc: m.getCc(), sujet: sansEtiquette_(sujet), etiquette: etiquette_(sujet), corps: m.getPlainBody(), html: html, piecesJointes: pj, date: fmtDateTime_(m.getDate()) };
}

function apiBrouillonEnvoyer_(p) {
  var d = brouillon_(p.id), m = d.getMessage();
  if (!String(m.getTo() || '').trim()) throw new Error('Ce brouillon n\'a pas de destinataire.');
  // l'étiquette de relecture ne part jamais chez le destinataire
  if (etiquette_(m.getSubject())) {
    var pj = [];
    try { pj = m.getAttachments({ includeInlineImages: false }); } catch (e) { pj = []; }
    var opt = { htmlBody: m.getBody(), name: getConfig_().NOM_STRUCTURE };
    if (pj.length) opt.attachments = pj;
    d.update(m.getTo(), sansEtiquette_(m.getSubject()), m.getPlainBody(), opt);
    d = brouillon_(p.id); m = d.getMessage();
  }
  d.send();
  logAction_('brouillon envoyé', m.getSubject() + ' -> ' + m.getTo());
  return true;
}

function apiBrouillonModifier_(p) {
  var d = brouillon_(p.id);
  var to = verifDest_(p.a);
  var sujet = sansEtiquette_(p.sujet);
  if (!sujet) throw new Error('Objet vide.');
  d.update(to, sujet, p.corps || '', optionsMail_(p.corps || ''));
  logAction_('brouillon modifié', sujet);
  return true;
}

function apiBrouillonSupprimer_(p) {
  var d = brouillon_(p.id), sujet = d.getMessage().getSubject();
  d.deleteDraft();
  logAction_('brouillon supprimé', sujet);
  return true;
}

/** Élèves à relancer (pas payés depuis X jours ou infos manquantes) */
/**
 * Qui est à relancer, et surtout POURQUOI les autres ne le sont pas.
 * Les compteurs se répartissent exactement : total = relancés + chaque motif d'exclusion.
 */
function aRelancerDetail_() {
  var c = getConfig_();
  var jours = num_(c.RELANCE_APRES_JOURS) || 10;
  var joursInter = num_(c.RELANCE_INTERVALLE_JOURS) || 7;
  var apres = jours * 86400000, inter = joursInter * 86400000;
  var now = Date.now();
  var st = { total: 0, desinscrits: 0, sansEmail: 0, payes: 0, offerts: 0, recents: 0, relanceRecente: 0, aRelancer: 0, foyers: 0, aVerifier: 0, jours: jours, joursInter: joursInter };
  var out = [];
  loadEleves_('cours').concat(loadEleves_('stages')).forEach(function (e) {
    st.total++;
    if (e.statut === 'Désinscrit') { st.desinscrits++; return; }
    if (!e.emailBest) { st.sansEmail++; return; }
    // relance de PAIEMENT : qui a payé n'y figure pas, même s'il manque des informations
    // (le complément d'infos est un autre message, à envoyer séparément)
    if (e.paye === 'Oui') { st.payes++; return; }
    if (reducEleve_(e).rien) { st.offerts++; return; }
    if (e.tsRaw && now - e.tsRaw < apres) { st.recents++; return; }
    var last = parseDate_(e.derniereRelance);
    if (last && now - last.getTime() < inter) { st.relanceRecente++; return; }
    out.push(e);
  });
  st.aRelancer = out.length;
  // regroupement par foyer : une famille avec trois enfants reçoit UN mail, pas trois
  var par = {}, ordre = [];
  out.forEach(function (e) {
    var k = norm_(e.emailBest);
    if (!par[k]) { par[k] = []; ordre.push(k); }
    par[k].push(e);
  });
  var foyers = ordre.map(function (k) {
    var l = par[k];
    var m = corpsRelance_(l);
    return { email: l[0].emailBest, eleves: l, sujet: m.sujet, corps: m.corps, montant: m.montant };
  });
  st.foyers = foyers.length;
  st.aVerifier = foyers.filter(function (f) { return f.montant === null; }).length;
  return { eleves: out, foyers: foyers, stats: st };
}

function aRelancer_() { return aRelancerDetail_().eleves; }

function apiARelancer_() {
  var d = aRelancerDetail_();
  // le logo part une fois pour toutes : recopié dans chaque mail il pèserait cent fois son poids,
  // l'ERP le remet dans l'aperçu à la place de {{LOGO}}
  var logoSrc = '';
  try { logoSrc = logo_(getConfig_().ID_LOGO) || ''; } catch (e) { logoSrc = ''; }
  return {
    liste: d.eleves.map(function (e) {
      return { id: e.id, source: e.source, nom: e.nom, prenom: e.prenom, paye: e.paye, manquants: e.manquants, ts: e.ts, email: e.emailBest, derniereRelance: e.derniereRelance };
    }),
    logo: logoSrc,
    foyers: (d.foyers || []).map(function (f) {
      return {
        email: f.email, sujet: f.sujet, montant: f.montant,
        // texte complet et rendu HTML du mail tel qu'il partira : l'aperçu n'a plus rien à demander au serveur
        corps: f.corps, html: mailHtml_(f.corps, '{{LOGO}}'),
        eleves: f.eleves.map(function (e) {
          return { id: e.id, source: e.source, nom: e.nom, prenom: e.prenom, paye: e.paye, montantRecu: e.montantRecu,
            manquants: e.manquants || [],
            reduction: String(e.reductions || '').trim(), commentaire: String(e.commentaire || '').trim(), ts: e.ts };
        })
      };
    }),
    stats: d.stats
  };
}

/**
 * Note une relance faite en dehors de l'ERP (mail écrit à la main, appel téléphonique).
 * Sans ça, la personne reste éternellement dans « À relancer » puisque la date ne s'écrit
 * que quand c'est l'ERP qui a créé le brouillon.
 */
function apiNoterRelance_(p) {
  var n = 0, err = [];
  (p.items || []).slice(0, 80).forEach(function (it) {
    try { noterEnvoi_(apiEleve_({ source: it.source, id: it.id }), 'relance'); n++; }
    catch (e) { err.push(it.id + ' : ' + e.message); }
  });
  logAction_('relances notées à la main', n + ' inscription(s)');
  return { n: n, erreurs: err };
}

/**
 * Crée les brouillons de relance pour une liste d'ids [{source,id}].
 * Les enfants d'une même famille sont regroupés : un seul brouillon par adresse mail.
 */
function apiBrouillonsRelance_(p) {
  var voulus = {};
  (p.items || []).slice(0, 60).forEach(function (it) { voulus[it.source + '|' + it.id] = 1; });
  var d = aRelancerDetail_();
  var n = 0, err = [];
  (d.foyers || []).forEach(function (f) {
    var pris = f.eleves.filter(function (e) { return voulus[e.source + '|' + e.id]; });
    if (!pris.length) return;
    try {
      var m = corpsRelance_(pris);
      creerBrouillon_(f.email, m.sujet, m.corps);
      pris.forEach(function (e) { noterEnvoi_(e, 'relance'); });
      logAction_('brouillon relance', f.email + ' (' + pris.map(function (e) { return e.prenom; }).join(', ') + ')');
      n++;
    } catch (e) { err.push(f.email + ' : ' + e.message); }
  });
  return { crees: n, erreurs: err };
}

// ---------- déclencheurs ----------

function onInscriptionCours(ev) { viaDeclencheurOuProprio_(ev); handleInscription_('cours', ev); }
function onInscriptionStage(ev) { viaDeclencheurOuProprio_(ev); handleInscription_('stages', ev); }

function handleInscription_(source, ev) {
  try {
    var c = getConfig_(true);
    clearMemo_();
    var list = loadEleves_(source); // rattache aussi le suivi d'une réponse modifiée
    var row = ev && ev.range ? ev.range.getRow() : responseSheet_(source).getLastRow();
    var e = list.filter(function (x) { return x.row === row; })[0];
    if (!e) return;
    // une réponse modifiée par le parent redéclenche le trigger : on ne réécrit pas le suivi
    if (!e.hasSuivi) saveSuivi_({ ID: e.id, Source: source, Nom: e.nom, Prenom: e.prenom, Statut: 'Actif', Paye: 'Non' });
    if (norm_(c.CONFIRMATION_AUTO) === 'oui' && e.emailBest && !e.hasSuivi && !e.confirmation) {
      var m = corpsConfirmation_(e);
      creerBrouillon_(e.emailBest, m.sujet, m.corps);
      saveSuivi_({ ID: e.id, Confirmation: new Date() });
    }
    logAction_('nouvelle inscription ' + source, e.prenom + ' ' + e.nom);
  } catch (err) {
    logAction_('ERREUR inscription ' + source, err.message);
  }
}

/** Tous les jours à 7h */
function dailyJob(ev) {
  viaDeclencheurOuProprio_(ev);
  var c = getConfig_(true);
  if (norm_(c.RELANCES_AUTO) === 'oui') {
    var items = aRelancer_().slice(0, 15);
    apiBrouillonsRelance_({ items: items.map(function (e) { return { source: e.source, id: e.id }; }) });
  }
  // fiches de présence : remet les croix en face des bons noms (ajouts / départs / changements de cours)
  try { syncToutesFiches_(); } catch (e) { logAction_('ERREUR fiches', e.message); }
  if (c.DIGEST_EMAILS) {
    try { envoyerRecap_(c.DIGEST_EMAILS); } catch (e) { logAction_('ERREUR récap', e.message); }
  }
}

function esc_(s) {
  return String(s === undefined || s === null ? '' : s).replace(/[&<>"]/g, function (ch) { return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[ch]; });
}

/** Récap quotidien pour l'équipe (pas de données santé détaillées dans le mail) */
function envoyerRecap_(to) {
  var a = apiAccueil_();
  var h = '<div style="font-family:Arial,sans-serif;max-width:640px;color:#151313">';
  h += '<h2 style="color:#C8262B;margin:0 0 4px">' + esc_(getConfig_().NOM_STRUCTURE) + ' · ' + esc_(a.date) + '</h2>';
  h += '<p style="margin:0 0 16px;color:#555">Récap automatique de l\'ERP</p>';
  if (a.coursDuJour.length) {
    h += '<h3>Cours aujourd\'hui</h3><ul>';
    a.coursDuJour.forEach(function (k) {
      h += '<li><b>' + esc_(k.horaire) + '</b> ' + esc_(k.libelle) + ' : ' + k.inscrits + ' élèves' + (k.alertes ? ' · <span style="color:#C8262B">' + k.alertes + ' alerte(s) santé, voir ERP</span>' : '') + '</li>';
    });
    h += '</ul>';
  }
  if (a.rappels.length) {
    h += '<h3>Rappels</h3><ul>';
    a.rappels.forEach(function (r) { h += '<li>' + (r.enRetard ? '<b style="color:#C8262B">En retard</b> ' : '') + esc_(r.Date) + ' ' + esc_(r.Type) + ' : ' + esc_(r.Titre) + '</li>'; });
    h += '</ul>';
  }
  if (a.agenda.length) {
    h += '<h3>Agenda (7 jours)</h3><ul>';
    a.agenda.slice(0, 12).forEach(function (ev) { h += '<li>' + esc_(ev.quand) + ' : ' + esc_(ev.titre) + '</li>'; });
    h += '</ul>';
  }
  h += '<h3>Chiffres</h3><ul><li>' + a.stats.eleves + ' élèves en cours annuels, ' + a.stats.stagiaires + ' inscrits en stage</li>' +
    '<li>' + a.stats.impayes + ' inscription(s) sans paiement noté</li>' +
    '<li>' + a.stats.aRelancer + ' à relancer · ' + a.stats.brouillons + ' brouillon(s) Gmail en attente · ' + a.stats.nonLus + ' mail(s) non lu(s)</li></ul>';
  if (a.nouveaux.length) {
    h += '<h3>Dernières inscriptions</h3><ul>';
    a.nouveaux.forEach(function (e) { h += '<li>' + esc_(e.ts) + ' : ' + esc_(e.prenom + ' ' + e.nom) + ' (' + esc_(e.source) + ')</li>'; });
    h += '</ul>';
  }
  var url = ScriptApp.getService().getUrl();
  if (url) h += '<p><a href="' + url + '" style="background:#C8262B;color:#fff;padding:10px 16px;border-radius:8px;text-decoration:none">Ouvrir l\'ERP</a></p>';
  h += '</div>';
  MailApp.sendEmail({ to: to, subject: 'Récap du jour · ' + a.date, htmlBody: h, name: 'ERP ' + getConfig_().NOM_STRUCTURE });
}

/** Derniers fils Gmail + brouillons en attente */
function apiMails_() {
  // brouillons (et fil de discussion auquel ils répondent : badge "réponse prête")
  var tous = GmailApp.getDrafts();
  var parFil = {};
  var drafts = tous.slice(0, 40).map(function (d) {
    var m = d.getMessage();
    var fil = '';
    try { fil = m.getThread().getId(); } catch (e) { fil = ''; }
    if (fil && !parFil[fil]) parFil[fil] = d.getId();
    var suj = m.getSubject() || '';
    return { id: d.getId(), sujet: sansEtiquette_(suj) || '(sans objet)', etiquette: etiquette_(suj), a: m.getTo(), date: fmtDateTime_(m.getDate()), fil: fil,
      extrait: String(m.getPlainBody() || '').replace(/\s+/g, ' ').slice(0, 120) };
  });
  // boîte de réception : les 5 derniers mails reçus (lus ou non)
  var list = GmailApp.search('in:inbox', 0, 5).map(function (t) {
    var m = t.getMessages();
    var last = m[m.length - 1];
    return { id: t.getId(), sujet: t.getFirstMessageSubject() || '(sans objet)', de: last.getFrom(), date: fmtDateTime_(t.getLastMessageDate()),
      nonLu: t.isUnread(), extrait: last.getPlainBody().replace(/\s+/g, ' ').slice(0, 160), lien: 'https://mail.google.com/mail/#inbox/' + t.getId(),
      brouillon: parFil[t.getId()] || '' };
  });
  return { threads: list, drafts: drafts, nonLus: GmailApp.getInboxUnreadCount() };
}

/** Aperçu d'un mail reçu, dans l'ERP */
function apiMailVoir_(p) {
  var t = GmailApp.getThreadById(String(p.id || ''));
  if (!t) throw new Error('Mail introuvable (déplacé ou supprimé ?).');
  var ms = t.getMessages();
  var m = ms[ms.length - 1];
  var brouillon = '';
  GmailApp.getDrafts().slice(0, 40).forEach(function (d) {
    if (brouillon) return;
    try { if (d.getMessage().getThread().getId() === t.getId()) brouillon = d.getId(); } catch (e) { }
  });
  return {
    id: t.getId(), sujet: t.getFirstMessageSubject() || '(sans objet)', de: m.getFrom(), a: m.getTo(),
    date: fmtDateTime_(m.getDate()), corps: m.getPlainBody(), html: m.getBody() || '', nonLu: t.isUnread(),
    messages: ms.length, brouillon: brouillon, lien: 'https://mail.google.com/mail/#inbox/' + t.getId()
  };
}

/** Boîte de réception : mails cochés -> corbeille ou marqués comme lus. p = { ids: [threadId], action: 'corbeille' | 'lu' } */
function apiMailsAction_(p) {
  var ids = (p.ids || []).slice(0, 50);
  if (!ids.length) throw new Error('Coche au moins un mail.');
  var n = 0;
  ids.forEach(function (id) {
    var t = null;
    try { t = GmailApp.getThreadById(String(id)); } catch (e) { t = null; }
    if (!t) return;
    if (p.action === 'corbeille') t.moveToTrash(); else t.markRead();
    n++;
  });
  logAction_('mails ' + (p.action === 'corbeille' ? 'supprimés' : 'lus'), String(n));
  return { n: n };
}

/* ===== Divers.gs ===== */
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

/* ===== Presences.gs ===== */
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

/* ===== ModeleEco.gs ===== */
/**
 * ModeleEco.gs : rémunération par intervenant dans Modele_economique.
 *
 * Un seul moteur, deux notions par personne :
 *  - le POTENTIEL : ce que la personne toucherait pour le travail qu'elle assure,
 *    calculé en permanence, qu'elle soit payée ou non ;
 *  - le A PAYER : le potentiel, mais seulement pour les natures activées
 *    (payé activité / payé admin / payé technique). C'est lui qui descend dans la compta.
 *
 * Tout se pilote depuis le tableau INTERVENANTS de l'onglet Parametres, lisible et
 * modifiable depuis l'ERP (écran Rémunération).
 *
 * majModeleEco() se lance depuis l'éditeur Apps Script. Elle fait une sauvegarde,
 * ne supprime aucune ligne de données, conserve les réglages déjà saisis, et se relance sans casse.
 */

var ECO_COLS_INTERV = 3;   // Intervenant 1, 2, 3 sur une ligne d'activité

function ecoSS_() {
  var id = getConfig_().ID_MODELE_ECO;
  if (!id) throw new Error('ID_MODELE_ECO vide dans Config.');
  return SpreadsheetApp.openById(id);
}

function colLettre_(c) {
  var s = '';
  while (c > 0) { var m = (c - 1) % 26; s = String.fromCharCode(65 + m) + s; c = Math.floor((c - 1) / 26); }
  return s;
}

function ecoLigneExacte_(sh, libelle) {
  var v = sh.getRange(1, 1, Math.max(sh.getLastRow(), 1), 1).getValues();
  var cible = norm_(libelle);
  for (var i = 0; i < v.length; i++) if (norm_(v[i][0]) === cible) return i + 1;
  return 0;
}

function ecoLigne_(sh, libelle) {
  var v = sh.getRange(1, 1, Math.max(sh.getLastRow(), 1), 1).getValues();
  var cible = norm_(libelle);
  for (var i = 0; i < v.length; i++) if (norm_(v[i][0]).indexOf(cible) === 0) return i + 1;
  return 0;
}

function ecoExige_(row, libelle, onglet) {
  if (!row) throw new Error('Ligne « ' + libelle + ' » introuvable dans l\'onglet ' + onglet + '. Le classeur a changé, migration annulée.');
  return row;
}

function ecoEntete_(sh, libelle) {
  var v = sh.getDataRange().getValues();
  var cible = norm_(libelle);
  for (var r = 0; r < v.length; r++) for (var c = 0; c < v[r].length; c++) if (norm_(v[r][c]) === cible) return { row: r + 1, col: c + 1 };
  return null;
}

function ecoColonne_(sh, headerRow, libelle) {
  var v = sh.getRange(headerRow, 1, 1, Math.max(sh.getLastColumn(), 1)).getValues()[0];
  var cible = norm_(libelle);
  for (var c = 0; c < v.length; c++) if (norm_(v[c]) === cible) return c + 1;
  return 0;
}

/** Même chose, mais sur une expression : sert aux colonnes dont l'intitulé a bougé à la main. */
function ecoColonneRe_(sh, headerRow, re) {
  var v = sh.getRange(headerRow, 1, 1, Math.max(sh.getLastColumn(), 1)).getValues()[0];
  for (var c = 0; c < v.length; c++) if (re.test(norm_(v[c]))) return c + 1;
  return 0;
}

/**
 * Colonne du cout theorique des intervenants. Son intitule a plusieurs vies selon les versions
 * du classeur et les retouches a la main : « Cout intervenants », « Cout total »,
 * « Cout 2 intervenants (theorique) », « Cout intervenants (theorique) ». On les accepte toutes,
 * sinon la mise a jour s arrete sur une colonne introuvable alors qu elle est juste la.
 * La colonne du cout REEL est exclue, elle a son propre traitement.
 */
var ECO_RE_COUT_THEO = /^cout( \d+)?( par)?( intervenants?| total)( \(theorique\))?$/;

/** Message d'erreur qui dit ce qui manque ET ce que la ligne d'en-tete contient vraiment */
function ecoManque_(sh, hr, onglet, attendu) {
  var v = sh.getRange(hr, 1, 1, Math.max(sh.getLastColumn(), 1)).getValues()[0].filter(String);
  return 'Colonne « ' + attendu + ' » introuvable dans ' + onglet + ' (en-tete ligne ' + hr + '). Trouve : ' + v.join(' | ') + '.';
}

function ecoFinTableau_(sh, headerRow) {
  var v = sh.getRange(1, 1, Math.max(sh.getLastRow(), 1), 1).getValues();
  for (var r = headerRow; r < v.length; r++) if (norm_(v[r][0]).indexOf('total') === 0) return { last: r, totalRow: r + 1 };
  return { last: v.length, totalRow: 0 };
}

function pRef_(row, col) { return 'Parametres!$' + colLettre_(col || 2) + '$' + row; }

// ---------- 1. Parametres : tarifs + tableau INTERVENANTS ----------

// colonnes du tableau INTERVENANTS (A = nom)
var IV_NOM = 1, IV_ACTIVITE = 2, IV_ADMIN = 3, IV_TECH = 4, IV_H_ADMIN = 5, IV_H_TECH = 6, IV_H_ACTIVITE = 7, IV_OBJECTIF = 8, IV_MOIS = 9;
var ECO_INTERV_ENTETE = ['Intervenant', 'Paye activite', 'Paye admin', 'Paye technique',
  'Heures admin / mois', 'Heures technique / mois', 'Heures activite en plus / an', 'Objectif mensuel', 'Mois de lissage'];
// valeurs de depart, reprises de ce qui existe deja dans le classeur quand c'est le cas
var ECO_INTERV_DEFAUT = [
  ['Camille', 'oui', 'non', 'oui', 0, 0, 0, 500, 12],
  ['Alex', 'non', 'non', 'non', 0, 0, 0, 500, 12],
  ['Dominique', 'non', 'oui', 'non', 40, 0, 0, 500, 12]
];
// les 6 natures des fiches intervenant, lignes 5 a 10, dans l ordre
var ECO_NATURE_CLES = ['cours', 'presta', 'stages', 'activitePlus', 'admin', 'technique'];
/**
 * Désignation et descriptif de chaque nature sur les factures des intervenants.
 * Une facture officielle doit dire ce qui est vendu, pas raconter le métier : une phrase
 * courte par nature, assez précise pour tenir devant un contrôle, et la facture reste
 * sur une seule page.
 */
var ECO_NATURE_FACTURE = {
  cours: { designation: 'Encadrement des cours annuels',
    detail: 'Encadrement des cours hebdomadaires de la structure.' },
  presta: { designation: 'Interventions pour collectivites et partenaires',
    detail: 'Ateliers animes pour des collectivites, ecoles et structures partenaires.' },
  stages: { designation: 'Encadrement des stages de vacances',
    detail: 'Encadrement des stages organises pendant les vacances scolaires.' },
  activitePlus: { designation: 'Interventions complementaires',
    detail: 'Ateliers et interventions en dehors du planning annuel.' },
  admin: { designation: 'Prestations administratives, communication et développement',
    detail: 'Gestion administrative de l\'école, communication et outils de gestion.' },
  technique: { designation: 'Prestations techniques',
    detail: 'Entretien, réparation et aménagement du lieu et des installations.' }
};
var ECO_NOTE_INTERV = 'Paye = oui : la personne est facturee, son montant descend dans la compta. Paye = non : elle travaille mais ne facture pas, ' +
  'sa fiche continue d afficher ce qu elle toucherait (potentiel). Les heures sont des heures reelles, 1 seance = 2 h (voir PARAMETRES GENERAUX).';

function ecoParametres_(ss, journal) {
  var sh = ss.getSheetByName('Parametres');
  if (!sh) throw new Error('Onglet Parametres introuvable.');

  var rTitre = ecoLigne_(sh, 'REMUNERATION DES INTERVENANTS');
  if (rTitre) { sh.getRange(rTitre, 1).setValue('REMUNERATION ACTIVITE (prestataires, facturent a l association)'); journal.push('Parametres : bloc activite renomme'); }

  var rNb = ecoExige_(ecoLigneExacte_(sh, 'Nombre d intervenants par seance'), 'Nombre d intervenants par seance', 'Parametres');
  if (num_(sh.getRange(rNb, 2).getValue()) < 2) sh.getRange(rNb, 2).setValue(2);
  sh.getRange(rNb, 3).setValue('Sert au cout theorique quand aucun nom n est saisi sur la ligne');

  var rHeure = ecoLigneExacte_(sh, '1 seance = nbre heure');
  if (!rHeure) {
    var apresGen = ecoExige_(ecoLigne_(sh, 'PARAMETRES GENERAUX'), 'PARAMETRES GENERAUX', 'Parametres');
    sh.insertRowsAfter(apresGen, 1);
    sh.getRange(apresGen + 1, 1, 1, 3).setValues([['1 seance = nbre heure', 2, 'h = 1 seance']]);
    rHeure = apresGen + 1;
    journal.push('Parametres : ligne « 1 seance = nbre heure » ajoutee');
  }

  var rTech = ecoLigneExacte_(sh, 'Remuneration Technique');
  if (!rTech) {
    var base = sh.getLastRow() + 2;
    sh.getRange(base, 1).setValue('REMUNERATION TECHNIQUE (auto-entrepreneurs, facturent a l association)').setFontWeight('bold');
    sh.getRange(base + 1, 1, 1, 3).setValues([['Remuneration Technique', 30, 'Tarif seance (2h) : reparation, amenagement, optimisation du lieu et du site']]);
    rTech = base + 1;
    journal.push('Parametres : bloc REMUNERATION TECHNIQUE ajoute');
  }

  // tableau INTERVENANTS : on récupère ce qui est déjà saisi avant de le réécrire au bon format
  var rTab = ecoLigne_(sh, 'INTERVENANTS (qui');
  var ancien = {};
  if (rTab) {
    var hdr = sh.getRange(rTab + 1, 1, 1, Math.max(sh.getLastColumn(), 2)).getValues()[0].map(function (x) { return norm_(x); });
    for (var r = rTab + 2; r <= sh.getLastRow(); r++) {
      var nom = String(sh.getRange(r, 1).getValue() || '').trim();
      if (!nom || norm_(nom).indexOf('paye = oui') === 0 || norm_(nom).indexOf('facture l activite = non') === 0) break;
      var o = { nom: nom };
      hdr.forEach(function (h, i) { if (h) o[h] = sh.getRange(r, i + 1).getValue(); });
      ancien[norm_(nom)] = o;
    }
  } else {
    rTab = sh.getLastRow() + 2;
  }

  var lignes = ECO_INTERV_DEFAUT.map(function (d) { return d.slice(); });
  Object.keys(ancien).forEach(function (k) {
    if (!lignes.some(function (l) { return norm_(l[0]) === k; })) lignes.push([ancien[k].nom, 'non', 'non', 'non', 0, 0, 0, 500, 12]);
  });
  var ouiNon = function (v, def) { var n = norm_(v); return n === 'oui' || n === 'non' ? n : def; };
  lignes.forEach(function (l) {
    var a = ancien[norm_(l[0])];
    if (!a) return;
    l[IV_ACTIVITE - 1] = ouiNon(a['paye activite'] !== undefined ? a['paye activite'] : a['facture l activite'], l[IV_ACTIVITE - 1]);
    l[IV_ADMIN - 1] = ouiNon(a['paye admin'], l[IV_ADMIN - 1]);
    l[IV_TECH - 1] = ouiNon(a['paye technique'], l[IV_TECH - 1]);
    if (a['heures admin / mois'] !== undefined && a['heures admin / mois'] !== '') l[IV_H_ADMIN - 1] = num_(a['heures admin / mois']);
    if (a['heures technique / mois'] !== undefined && a['heures technique / mois'] !== '') l[IV_H_TECH - 1] = num_(a['heures technique / mois']);
    if (a['heures activite en plus / an'] !== undefined && a['heures activite en plus / an'] !== '') l[IV_H_ACTIVITE - 1] = num_(a['heures activite en plus / an']);
    if (a['objectif mensuel'] !== undefined && a['objectif mensuel'] !== '') l[IV_OBJECTIF - 1] = num_(a['objectif mensuel']);
    if (a['mois de lissage'] !== undefined && a['mois de lissage'] !== '') l[IV_MOIS - 1] = num_(a['mois de lissage']);
  });

  sh.getRange(rTab, 1).setValue('INTERVENANTS (qui est paye, sur quoi, et combien)').setFontWeight('bold');
  sh.getRange(rTab + 1, 1, 1, ECO_INTERV_ENTETE.length).setValues([ECO_INTERV_ENTETE]).setFontWeight('bold');
  sh.getRange(rTab + 2, 1, lignes.length, ECO_INTERV_ENTETE.length).setValues(lignes);
  sh.getRange(rTab + 2 + lignes.length, 1).setValue(ECO_NOTE_INTERV);
  journal.push('Parametres : tableau INTERVENANTS (' + lignes.length + ' personnes, reglages conserves)');

  try {
    var dvON = SpreadsheetApp.newDataValidation().requireValueInList(['oui', 'non'], true).setAllowInvalid(false).build();
    sh.getRange(rTab + 2, IV_ACTIVITE, lignes.length, 3).setDataValidation(dvON);
  } catch (e) { }

  var gens = lignes.map(function (l, i) { return { nom: String(l[0]), row: rTab + 2 + i }; });
  return {
    sheet: sh, gens: gens, entete: rTab + 1, nbInterv: rNb, heureParSeance: rHeure,
    tarifCours: ecoExige_(ecoLigneExacte_(sh, 'Cours annuel - forfait par seance et par intervenant'), 'Cours annuel - forfait par seance', 'Parametres'),
    tarifStage: ecoExige_(ecoLigneExacte_(sh, 'Stage - par jour et par intervenant'), 'Stage - par jour et par intervenant', 'Parametres'),
    tarifAdmin: ecoExige_(ecoLigneExacte_(sh, 'Remuneration Admin / IT'), 'Remuneration Admin / IT', 'Parametres'),
    tarifTech: rTech
  };
}

// ---------- 2. onglets d'activité ----------

function ecoTarifDeFormule_(formule, rowNbInterv, defaut) {
  var refs = String(formule || '').match(/Parametres!\$?B\$?(\d+)/g) || [];
  for (var i = 0; i < refs.length; i++) {
    var n = +refs[i].replace(/\D+/g, '');
    if (n !== rowNbInterv) return n;
  }
  return defaut;
}

function ecoOngletActivite_(ss, P, conf, journal) {
  var sh = ss.getSheetByName(conf.onglet);
  if (!sh) throw new Error('Onglet ' + conf.onglet + ' introuvable.');
  var e = ecoEntete_(sh, conf.entete);
  if (!e) throw new Error('En-tete « ' + conf.entete + ' » introuvable dans ' + conf.onglet + '.');
  var hr = e.row;
  var bornes = ecoFinTableau_(sh, hr);
  var r1 = hr + 1, r2 = bornes.last, tot = bornes.totalRow;
  if (r2 < r1) throw new Error('Aucune ligne de donnees dans ' + conf.onglet + '.');

  var cTheo = ecoColonne_(sh, hr, conf.libelleCout) || ecoColonne_(sh, hr, 'Cout intervenants (theorique)')
    || ecoColonneRe_(sh, hr, ECO_RE_COUT_THEO);
  if (!cTheo) {
    throw new Error('Colonne du cout des intervenants introuvable dans ' + conf.onglet + ' (ligne ' + hr + '). ' +
      'Attendu « ' + conf.libelleCout +' », « Cout intervenants (theorique) » ou « Cout 2 intervenants (theorique) ». ' +
      'Trouve : ' + sh.getRange(hr, 1, 1, Math.max(sh.getLastColumn(), 1)).getValues()[0]
        .filter(String).join(' | ') + '.');
  }

  // tarif par intervenant dans sa propre colonne, repris de l'existant ligne par ligne
  var cTarif = ecoColonne_(sh, hr, 'Cout unitaire par intervenant') || ecoColonne_(sh, hr, 'Tarif par intervenant');
  var ancien = sh.getRange(r1, cTheo, r2 - r1 + 1, 1).getFormulas();
  if (!cTarif && conf.colTarifExistante) {
    cTarif = ecoColonne_(sh, hr, conf.colTarifExistante);
    if (cTarif) {
      var f = sh.getRange(r1, cTarif, r2 - r1 + 1, 1).getFormulas();
      for (var i = 0; i < f.length; i++) {
        var n = ecoTarifDeFormule_(f[i][0], P.nbInterv, 0);
        if (n) sh.getRange(r1 + i, cTarif).setFormula('=' + pRef_(n));
      }
      sh.getRange(hr, cTarif).setValue('Cout unitaire par intervenant');
      journal.push(conf.onglet + ' : colonne tarif nettoyee');
    }
  }
  if (!cTarif) {
    sh.insertColumnsBefore(cTheo, 1);
    cTarif = cTheo; cTheo = cTheo + 1;
    sh.getRange(hr, cTarif).setValue('Tarif par intervenant').setFontWeight('bold');
    for (var j = 0; j < ancien.length; j++) {
      sh.getRange(r1 + j, cTarif).setFormula('=' + pRef_(ecoTarifDeFormule_(ancien[j][0], P.nbInterv, conf.tarifDefaut)));
    }
    journal.push(conf.onglet + ' : colonne Tarif par intervenant');
  }

  // deux colonnes de cout theorique dans le meme tableau : on n en touche qu une et on le dit,
  // supprimer une colonne a la place de Dominique casserait les formules qui pointent dessus
  var doublons = [];
  var ent = sh.getRange(hr, 1, 1, Math.max(sh.getLastColumn(), 1)).getValues()[0];
  for (var dc = 0; dc < ent.length; dc++) {
    if (dc + 1 !== cTheo && ECO_RE_COUT_THEO.test(norm_(ent[dc]))) doublons.push(colLettre_(dc + 1));
  }
  if (doublons.length) {
    journal.push(conf.onglet + ' : ATTENTION, colonne(s) de cout theorique en doublon en ' + doublons.join(', ') +
      '. La migration n utilise que la colonne ' + colLettre_(cTheo) + '. A supprimer a la main une fois verifie.');
  }

  sh.getRange(hr, cTheo).setValue('Cout intervenants (theorique)');
  var cReel = ecoColonne_(sh, hr, 'Cout intervenants (reel)');
  if (!cReel) {
    sh.insertColumnsAfter(cTheo, 1);
    cReel = cTheo + 1;
    sh.getRange(hr, cReel).setValue('Cout intervenants (reel)').setFontWeight('bold');
    journal.push(conf.onglet + ' : colonne Cout intervenants (reel)');
  }

  var cInt1 = ecoColonne_(sh, hr, 'Intervenant 1');
  if (!cInt1) {
    cInt1 = Math.max(sh.getLastColumn(), cReel) + 1;
    for (var k = 0; k < ECO_COLS_INTERV; k++) sh.getRange(hr, cInt1 + k).setValue('Intervenant ' + (k + 1)).setFontWeight('bold');
    journal.push(conf.onglet + ' : colonnes Intervenant 1 a ' + ECO_COLS_INTERV);
  }
  var cPers = cInt1 + ECO_COLS_INTERV;
  P.gens.forEach(function (g, k) { sh.getRange(hr, cPers + k).setValue('Part ' + g.nom).setFontWeight('bold'); });
  // les deux premiers de la liste sur la piste par defaut, uniquement si la case est vide
  for (var pr = r1; pr <= r2; pr++) {
    for (var pc = 0; pc < Math.min(2, P.gens.length); pc++) {
      var cell = sh.getRange(pr, cInt1 + pc);
      if (String(cell.getValue() || '').trim() === '') cell.setValue(P.gens[pc].nom);
    }
  }

  // sans ces deux colonnes les formules s ecriraient vides et casseraient le classeur en silence
  var cUnites = ecoColonne_(sh, hr, conf.colUnites);
  if (!cUnites) throw new Error(ecoManque_(sh, hr, conf.onglet, conf.colUnites));
  if (!cTarif) throw new Error(ecoManque_(sh, hr, conf.onglet, 'Tarif par intervenant'));
  var Lu = colLettre_(cUnites);
  var Lt = colLettre_(cTarif);
  var L1 = colLettre_(cInt1), L3 = colLettre_(cInt1 + ECO_COLS_INTERV - 1);
  var cRec = ecoColonne_(sh, hr, conf.libelleRecette);
  if (!cRec) throw new Error(ecoManque_(sh, hr, conf.onglet, conf.libelleRecette));
  var cMarge = ecoColonne_(sh, hr, conf.libelleMarge);
  var cPct = conf.libellePct ? ecoColonne_(sh, hr, conf.libellePct) : 0;

  for (var r = r1; r <= r2; r++) {
    var nommes = 'COUNTIF($' + L1 + r + ':$' + L3 + r + ',"<>")';
    var base = '$' + Lu + r + '*$' + Lt + r;
    sh.getRange(r, cTheo).setFormula('=' + base + '*IF(' + nommes + '=0,' + pRef_(P.nbInterv) + ',' + nommes + ')');
    // Part <nom> = ce que la personne toucherait pour cette ligne, qu elle soit payee ou non
    P.gens.forEach(function (g, k) {
      sh.getRange(r, cPers + k).setFormula('=IF(COUNTIF($' + L1 + r + ':$' + L3 + r + ',' + pRef_(g.row, IV_NOM) + ')=0,0,' + base + ')');
    });
    // cout reel = seulement les personnes activees sur l'activite de terrain
    var reel = P.gens.map(function (g, k) {
      return 'IF(' + pRef_(g.row, IV_ACTIVITE) + '="oui",' + colLettre_(cPers + k) + r + ',0)';
    }).join('+');
    sh.getRange(r, cReel).setFormula('=' + reel);
    if (cMarge) sh.getRange(r, cMarge).setFormula('=' + colLettre_(cRec) + r + '-' + colLettre_(cReel) + r);
    if (cPct) sh.getRange(r, cPct).setFormula('=IFERROR(' + colLettre_(cMarge) + r + '/' + colLettre_(cRec) + r + ',0)');
  }

  if (tot) {
    var cols = [cTheo, cReel, cMarge].concat(P.gens.map(function (g, k) { return cPers + k; }));
    cols.forEach(function (c) { if (c) sh.getRange(tot, c).setFormula('=SUM(' + colLettre_(c) + r1 + ':' + colLettre_(c) + r2 + ')'); });
    if (cPct) sh.getRange(tot, cPct).setFormula('=IFERROR(' + colLettre_(cMarge) + tot + '/' + colLettre_(cRec) + tot + ',0)');
  }

  try {
    var liste = P.sheet.getRange(P.entete + 1, 1, P.gens.length, 1);
    var dv = SpreadsheetApp.newDataValidation().requireValueInRange(liste, true).setAllowInvalid(false).build();
    sh.getRange(r1, cInt1, r2 - r1 + 1, ECO_COLS_INTERV).setDataValidation(dv);
  } catch (err) { journal.push(conf.onglet + ' : menu deroulant non pose (' + err.message + ')'); }

  return { sheet: sh, nom: conf.onglet, hr: hr, r1: r1, r2: r2, tot: tot, cTarif: cTarif, cTheo: cTheo, cReel: cReel, cPers: cPers, cInt1: cInt1 };
}

// ---------- 3. fiches par intervenant (lecture seule) ----------

/** Montant annuel admin/technique : heures par mois / heures par seance x tarif x mois factures */
function ecoFormuleHeures_(P, g, colHeures, tarifRow) {
  return '=' + pRef_(g.row, colHeures) + '/' + pRef_(P.heureParSeance) + '*' + pRef_(tarifRow) + '*' + pRef_(g.row, IV_MOIS);
}

function ecoFicheIntervenant_(ss, P, act, g, journal) {
  var sh = ss.getSheetByName(g.nom);
  if (!sh) { sh = ss.insertSheet(g.nom); journal.push('Onglet ' + g.nom + ' cree'); }
  var maxR = Math.max(sh.getLastRow(), 1), maxC = Math.max(sh.getLastColumn(), 1);
  sh.getRange(1, 1, maxR, maxC).clearContent();

  var parts = [act.cours, act.presta, act.stages];
  var k = P.gens.map(function (x) { return x.nom; }).indexOf(g.nom);
  var somme = function (a) { return "SUM('" + a.nom + "'!" + colLettre_(a.cPers + k) + a.r1 + ':' + colLettre_(a.cPers + k) + a.r2 + ')'; };
  var activiteOui = pRef_(g.row, IV_ACTIVITE) + '="oui"';
  var adminOui = pRef_(g.row, IV_ADMIN) + '="oui"';
  var techOui = pRef_(g.row, IV_TECH) + '="oui"';

  sh.getRange(1, 1).setValue(g.nom).setFontWeight('bold');
  sh.getRange(1, 2).setFormula('=IF(OR(' + activiteOui + ',' + adminOui + ',' + techOui + '),"PAYE","BENEVOLE POUR LE MOMENT")');
  sh.getRange(1, 3).setValue('Fiche calculee, ne rien saisir ici. Les reglages sont dans Parametres, tableau INTERVENANTS, ou dans l ERP.');
  sh.getRange(2, 1).setValue('RESUME DE LA REMUNERATION').setFontWeight('bold');

  sh.getRange(4, 1, 1, 5).setValues([['Nature', 'Paye ?', 'Volume', 'Annuel', 'Mensuel lisse']]).setFontWeight('bold');
  var L = [
    { lib: 'Cours annuels', paye: activiteOui, vol: '=IFERROR(' + somme(act.cours) + '/' + pRef_(P.tarifCours) + ',0)&" séances"', mt: '=' + somme(act.cours) },
    { lib: 'Prestations collectivites et ecoles', paye: activiteOui, vol: '=COUNTIF(\'' + act.presta.nom + '\'!' + colLettre_(act.presta.cPers + k) + act.presta.r1 + ':' + colLettre_(act.presta.cPers + k) + act.presta.r2 + ',">0")&" prestations"', mt: '=' + somme(act.presta) },
    { lib: 'Stages vacances', paye: activiteOui, vol: '=IFERROR(' + somme(act.stages) + '/' + pRef_(P.tarifStage) + ',0)&" jours"', mt: '=' + somme(act.stages) },
    { lib: 'Activite en plus (saisie)', paye: activiteOui, vol: '=' + pRef_(g.row, IV_H_ACTIVITE) + '&" h / an"', mt: '=' + pRef_(g.row, IV_H_ACTIVITE) + '/' + pRef_(P.heureParSeance) + '*' + pRef_(P.tarifCours) },
    { lib: 'Admin / com / dev', paye: adminOui, vol: '=' + pRef_(g.row, IV_H_ADMIN) + '&" h / mois"', mt: ecoFormuleHeures_(P, g, IV_H_ADMIN, P.tarifAdmin) },
    { lib: 'Technique (reparation, amenagement, site)', paye: techOui, vol: '=' + pRef_(g.row, IV_H_TECH) + '&" h / mois"', mt: ecoFormuleHeures_(P, g, IV_H_TECH, P.tarifTech) }
  ];
  L.forEach(function (l, i) {
    var r = 5 + i;
    sh.getRange(r, 1).setValue(l.lib);
    sh.getRange(r, 2).setFormula('=IF(' + l.paye + ',"oui","non")');
    sh.getRange(r, 3).setFormula(l.vol);
    sh.getRange(r, 4).setFormula(l.mt);
    sh.getRange(r, 5).setFormula('=IFERROR(D' + r + '/' + pRef_(g.row, IV_MOIS) + ',0)');
  });

  var r1 = 5, r2 = 4 + L.length;
  var aPayer = L.map(function (l, i) { return 'IF(' + l.paye + ',D' + (r1 + i) + ',0)'; }).join('+');
  sh.getRange(11, 1).setValue('TOTAL POTENTIEL (si tout est active)').setFontWeight('bold');
  sh.getRange(11, 4).setFormula('=SUM(D' + r1 + ':D' + r2 + ')');
  sh.getRange(11, 5).setFormula('=IFERROR(D11/' + pRef_(g.row, IV_MOIS) + ',0)');
  sh.getRange(12, 1).setValue('TOTAL A PAYER').setFontWeight('bold');
  sh.getRange(12, 4).setFormula('=' + aPayer);
  sh.getRange(12, 5).setFormula('=IFERROR(D12/' + pRef_(g.row, IV_MOIS) + ',0)');
  sh.getRange(13, 1).setValue('Objectif mensuel');
  sh.getRange(13, 5).setFormula('=' + pRef_(g.row, IV_OBJECTIF));
  sh.getRange(14, 1).setValue('Ecart a l objectif (par mois)');
  sh.getRange(14, 5).setFormula('=E12-E13');
  sh.getRange(16, 1).setValue('Tarifs : ' + 'cours et seance collectivite au forfait par seance, demi-journee et journee au forfait, stage a la journee. Admin et technique : heures / ' +
    '(1 seance = 2 h) x tarif seance x mois de lissage.');
  return { sheet: sh, nom: g.nom };
}

// ---------- 4. Synthese ----------

function ecoSynthese_(ss, P, act, fiches, journal) {
  var sh = ss.getSheetByName(getConfig_().ONGLET_SYNTHESE || 'Synthese');
  if (!sh) throw new Error('Onglet Synthese introuvable.');

  var poles = [
    { label: 'Cours annuels', a: act.cours },
    { label: 'Prestations collectivites et ecoles', a: act.presta },
    { label: 'Stages vacances', a: act.stages }
  ];
  var premiere = 0, derniere = 0;
  poles.forEach(function (p) {
    var r = ecoLigne_(sh, p.label);
    if (!r) { journal.push('Synthese : ligne « ' + p.label + ' » introuvable'); return; }
    if (!premiere) premiere = r;
    derniere = r;
    sh.getRange(r, 3).setFormula("='" + p.a.nom + "'!" + colLettre_(p.a.cReel) + p.a.tot);
    sh.getRange(r, 4).setFormula('=B' + r + '-C' + r);
    sh.getRange(r, 5).setFormula('=IFERROR(D' + r + '/B' + r + ',0)');
  });
  // le libellé doit rester lisible par l'ERP : « TOTAL ANNEE » en A, la précision en C
  var rTotal = ecoLigne_(sh, 'TOTAL ANNEE');
  if (rTotal && premiere) {
    sh.getRange(rTotal, 1).setValue('TOTAL ANNEE');
    sh.getRange(rTotal, 6).setValue('recettes de la saison, cout des intervenants reellement payes');
    [2, 3, 4].forEach(function (c) { sh.getRange(rTotal, c).setFormula('=SUM(' + colLettre_(c) + premiere + ':' + colLettre_(c) + derniere + ')'); });
    sh.getRange(rTotal, 5).setFormula('=IFERROR(D' + rTotal + '/B' + rTotal + ',0)');
    journal.push('Synthese : poles sur le cout reel, libelle TOTAL ANNEE restaure');
  }

  // bloc rémunération : on le réécrit en entier à partir de son titre
  var rBloc = ecoLigne_(sh, 'REMUNERATION REEL PAR INTERVENANT') || ecoLigne_(sh, 'REMUNERATION PAR INTERVENANT') || ecoLigne_(sh, 'REVENU PAR INTERVENANT');
  if (!rBloc) rBloc = sh.getLastRow() + 2;
  var place = 0;
  for (var r = rBloc + 1; r <= sh.getLastRow(); r++) {
    var a = norm_(sh.getRange(r, 1).getValue());
    if (a === '' || a.indexOf('charges de structure') === 0 || a.indexOf('resultat net') === 0 || a.indexOf('leviers') === 0) break;
    place++;
  }
  var besoin = P.gens.length + 3;   // en-tete + personnes + total a payer + total potentiel
  if (place < besoin) { sh.insertRowsAfter(rBloc + Math.max(place, 1), besoin - place); journal.push('Synthese : ' + (besoin - place) + ' ligne(s) ajoutee(s)'); }
  else if (place > besoin) sh.getRange(rBloc + besoin + 1, 1, place - besoin, 8).clearContent();

  // on nettoie tout le bloc avant d'ecrire : les anciennes formules laissees en colonne B
  // seraient lues par l'ERP a la place des bons chiffres
  sh.getRange(rBloc, 1, besoin + 1, 9).clearContent();
  sh.getRange(rBloc, 1).setValue('REMUNERATION PAR INTERVENANT').setFontWeight('bold');
  sh.getRange(rBloc, 3).setValue('Le montant en face de chaque nom est ce qui est reellement du. Le potentiel est ce que chacun toucherait si tout etait active.');
  var hr = rBloc + 1;
  sh.getRange(hr, 1, 1, 9).setValues([['Intervenant', 'Total annuel', 'Par mois', 'Paye ?', 'Activite', 'Admin', 'Technique', 'Objectif', 'Ecart']]).setFontWeight('bold');

  P.gens.forEach(function (g, k) {
    var r = hr + 1 + k;
    var f = "'" + g.nom + "'!";
    sh.getRange(r, 1).setValue(g.nom);
    sh.getRange(r, 2).setFormula('=SUM(E' + r + ':G' + r + ')');
    sh.getRange(r, 3).setFormula('=IFERROR(B' + r + '/' + pRef_(g.row, IV_MOIS) + ',0)');
    sh.getRange(r, 4).setFormula('=' + f + 'B1');
    sh.getRange(r, 5).setFormula('=IF(' + pRef_(g.row, IV_ACTIVITE) + '<>"oui",0,' + f + 'D5+' + f + 'D6+' + f + 'D7+' + f + 'D8)');
    sh.getRange(r, 6).setFormula('=IF(' + pRef_(g.row, IV_ADMIN) + '<>"oui",0,' + f + 'D9)');
    sh.getRange(r, 7).setFormula('=IF(' + pRef_(g.row, IV_TECH) + '<>"oui",0,' + f + 'D10)');
    sh.getRange(r, 8).setFormula('=' + pRef_(g.row, IV_OBJECTIF));
    sh.getRange(r, 9).setFormula('=C' + r + '-H' + r);
  });
  var rT = hr + 1 + P.gens.length;
  sh.getRange(rT, 1).setValue('TOTAL A PAYER AUX INTERVENANTS').setFontWeight('bold');
  [2, 3, 5, 6, 7].forEach(function (c) { sh.getRange(rT, c).setFormula('=SUM(' + colLettre_(c) + (hr + 1) + ':' + colLettre_(c) + (rT - 1) + ')'); });
  var rP = rT + 1;
  sh.getRange(rP, 1).setValue('TOTAL POTENTIEL (si tout le monde est paye)');
  sh.getRange(rP, 2).setFormula('=' + P.gens.map(function (g) { return "'" + g.nom + "'!D11"; }).join('+'));
  sh.getRange(rP, 3).setFormula('=' + P.gens.map(function (g) { return "'" + g.nom + "'!E11"; }).join('+'));
  journal.push('Synthese : tableau par intervenant (' + P.gens.length + ' personnes) + total potentiel');

  // charges et résultats
  var rCharges = ecoLigne_(sh, 'CHARGES DE STRUCTURE - PREVI') || ecoLigne_(sh, 'CHARGES DE STRUCTURE');
  var rCaisse = ecoLigneExacte_(sh, 'CAISSE');
  var rReelTot = ecoLigne_(sh, 'TOTAL ANNEE REEL');

  var rNet = ecoLigne_(sh, 'RESULTAT NET PREVI') || ecoLigne_(sh, 'RESULTAT NET DE L ASSOCIATION');
  if (rNet && rTotal && rCharges) {
    sh.getRange(rNet, 1).setValue('RESULTAT NET DE L ASSOCIATION');
    sh.getRange(rNet, 2).setFormula('=B' + rTotal + '-B' + rT + '-B' + rCharges);
    sh.getRange(rNet, 3).setValue('PREVI : recettes de la saison - remuneration a payer - charges de structure prevues');
    journal.push('Synthese : resultat net previsionnel');
  }
  var rNetReel = ecoLigne_(sh, 'RESULTAT NET REEL');
  if (rNetReel && rCaisse) {
    sh.getRange(rNetReel, 1).setValue('RESULTAT NET REEL DE L ASSOCIATION');
    sh.getRange(rNetReel, 2).setFormula('=B' + rCaisse + '-B' + rT);
    sh.getRange(rNetReel, 3).setValue('REEL : caisse a ce jour (encaisse moins depenses) - remuneration a payer sur la saison');
    journal.push('Synthese : resultat net reel corrige (la caisse est deja nette des depenses)');
  }

  // réel contre théorique
  var rTheo = ecoLigne_(sh, 'COUT DES INTERVENANTS') || ecoLigne_(sh, 'COUT THEORIQUE');
  if (!rTheo) rTheo = sh.getLastRow() + 2;
  sh.getRange(rTheo, 1).setValue('COUT DES INTERVENANTS : PAYE ET THEORIQUE').setFontWeight('bold');
  sh.getRange(rTheo + 1, 1, 1, 3).setValues([['Pole', 'Paye', 'Theorique (2 intervenants payes)']]).setFontWeight('bold');
  [['Cours annuels', act.cours], ['Prestations', act.presta], ['Stages', act.stages]].forEach(function (p, i) {
    var r = rTheo + 2 + i;
    sh.getRange(r, 1).setValue(p[0]);
    sh.getRange(r, 2).setFormula("='" + p[1].nom + "'!" + colLettre_(p[1].cReel) + p[1].tot);
    sh.getRange(r, 3).setFormula("='" + p[1].nom + "'!" + colLettre_(p[1].cTheo) + p[1].tot);
  });
  var rTT = rTheo + 5;
  sh.getRange(rTT, 1).setValue('TOTAL ACTIVITE').setFontWeight('bold');
  sh.getRange(rTT, 2).setFormula('=SUM(B' + (rTheo + 2) + ':B' + (rTheo + 4) + ')');
  sh.getRange(rTT, 3).setFormula('=SUM(C' + (rTheo + 2) + ':C' + (rTheo + 4) + ')');
  journal.push('Synthese : bloc paye / theorique');

  return { blocRemu: rBloc, entete: hr, totalRemu: rT, potentiel: rP, theo: rTheo };
}

// ---------- point d'entrée ----------

function majModeleEco() {
  proprietaireSeulement_();
  var journal = [];
  var ss = ecoSS_();

  var sauvegarde = '';
  try {
    var f = DriveApp.getFileById(ss.getId());
    var copie = f.makeCopy('SAUVEGARDE ' + Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH-mm') + ' - ' + f.getName());
    sauvegarde = copie.getUrl();
    journal.push('Sauvegarde : ' + copie.getName());
  } catch (e) { journal.push('Sauvegarde impossible (' + e.message + '), on continue'); }

  var P = ecoParametres_(ss, journal);

  var act = {
    cours: ecoOngletActivite_(ss, P, {
      onglet: 'Cours_annuels', entete: 'Cours', libelleCout: 'Cout intervenants',
      libelleRecette: 'Recette annuelle', libelleMarge: 'Resultat',
      colUnites: 'Seances/an', tarifDefaut: P.tarifCours
    }, journal),
    stages: ecoOngletActivite_(ss, P, {
      onglet: 'Stages', entete: 'Stage', libelleCout: 'Cout intervenants',
      libelleRecette: 'Recette', libelleMarge: 'Marge', libellePct: '% marge',
      colUnites: 'Jours', tarifDefaut: P.tarifStage
    }, journal),
    presta: ecoOngletActivite_(ss, P, {
      onglet: 'Prestations', entete: 'Client', libelleCout: 'Cout total',
      libelleRecette: 'Recette', libelleMarge: 'Marge', libellePct: '% marge',
      colUnites: 'Quantite', colTarifExistante: 'Cout unitaire (2 interv.)', tarifDefaut: P.tarifCours
    }, journal)
  };

  var fiches = P.gens.map(function (g) { return ecoFicheIntervenant_(ss, P, act, g, journal); });
  journal.push('Fiches intervenants refaites : ' + P.gens.map(function (g) { return g.nom; }).join(', '));

  var S = ecoSynthese_(ss, P, act, fiches, journal);
  SpreadsheetApp.flush();
  logAction_('modele eco', journal.join(' ; '));
  try { Logger.log(journal.join('\n')); } catch (e) { }
  return { journal: journal, sauvegarde: sauvegarde, synthese: S, intervenants: P.gens.map(function (g) { return g.nom; }) };
}

// ---------- lecture / écriture depuis l'ERP ----------

/** Le tableau INTERVENANTS tel qu'il est, plus les montants calculés */
function apiRemunerations_() {
  var ss = ecoSS_();
  var shP = ss.getSheetByName('Parametres');
  if (!shP) return { lignes: [], total: 0, potentiel: 0, pret: false };
  var rTab = ecoLigne_(shP, 'INTERVENANTS (qui');
  if (!rTab) return { lignes: [], total: 0, potentiel: 0, pret: false };

  var gens = [];
  for (var r = rTab + 2; r <= shP.getLastRow(); r++) {
    var nom = String(shP.getRange(r, 1).getValue() || '').trim();
    if (!nom || norm_(nom).indexOf('paye = oui') === 0) break;
    var v = shP.getRange(r, 1, 1, ECO_INTERV_ENTETE.length).getValues()[0];
    var o = {
      nom: nom, row: r,
      payeActivite: norm_(v[IV_ACTIVITE - 1]) === 'oui',
      payeAdmin: norm_(v[IV_ADMIN - 1]) === 'oui',
      payeTech: norm_(v[IV_TECH - 1]) === 'oui',
      hAdmin: num_(v[IV_H_ADMIN - 1]), hTech: num_(v[IV_H_TECH - 1]), hActivite: num_(v[IV_H_ACTIVITE - 1]),
      objectif: num_(v[IV_OBJECTIF - 1]), mois: num_(v[IV_MOIS - 1]) || 12
    };
    o.cours = 0; o.presta = 0; o.stages = 0; o.activitePlus = 0; o.admin = 0; o.technique = 0;
    o.potentiel = 0; o.total = 0; o.activite = 0; o.parMois = 0; o.ecart = -o.objectif;
    o.natures = [];
    var fiche = ss.getSheetByName(nom);
    if (fiche) {
      var f = fiche.getRange(5, 1, 8, 5).getValues();      // A5:E12
      var d = f.map(function (x) { return [x[3]]; });      // colonne D (annuel)
      o.cours = num_(d[0][0]); o.presta = num_(d[1][0]); o.stages = num_(d[2][0]);
      o.activitePlus = num_(d[3][0]); o.admin = num_(d[4][0]); o.technique = num_(d[5][0]);
      o.potentiel = num_(d[6][0]); o.total = num_(d[7][0]);
      o.activite = o.payeActivite ? o.cours + o.presta + o.stages + o.activitePlus : 0;
      o.parMois = o.mois ? Math.round(o.total / o.mois * 100) / 100 : 0;
      o.ecart = Math.round((o.parMois - o.objectif) * 100) / 100;
      // les 6 natures de la fiche, telles qu'elles s'y affichent : sert à générer la facture
      for (var n = 0; n < 6; n++) {
        var lib = String(f[n][0] || '').trim();
        if (!lib) continue;
        o.natures.push({
          cle: ECO_NATURE_CLES[n], libelle: lib,
          paye: norm_(f[n][1]) === 'oui',
          volume: String(f[n][2] || '').trim(),
          annuel: Math.round(num_(f[n][3]) * 100) / 100,
          mensuel: Math.round(num_(f[n][4]) * 100) / 100
        });
      }
    }
    gens.push(o);
  }
  return {
    lignes: gens, pret: true,
    total: gens.reduce(function (t, g) { return t + (g.total || 0); }, 0),
    potentiel: gens.reduce(function (t, g) { return t + (g.potentiel || 0); }, 0),
    parMois: gens.reduce(function (t, g) { return t + (g.parMois || 0); }, 0)
  };
}

/** Écrit une ligne du tableau INTERVENANTS. p = { nom, payeActivite, payeAdmin, payeTech, hAdmin, hTech, hActivite, objectif, mois } */
function apiSaveRemuneration_(p) {
  var ss = ecoSS_();
  var shP = ss.getSheetByName('Parametres');
  var rTab = ecoLigne_(shP, 'INTERVENANTS (qui');
  if (!rTab) throw new Error('Tableau INTERVENANTS absent. Lance d\'abord majModeleEco depuis l\'editeur Apps Script.');
  var row = 0;
  for (var r = rTab + 2; r <= shP.getLastRow(); r++) {
    var nom = String(shP.getRange(r, 1).getValue() || '').trim();
    if (!nom || norm_(nom).indexOf('paye = oui') === 0) break;
    if (norm_(nom) === norm_(p.nom) || (typeof memeNom_ === 'function' && memeNom_(nom, p.nom))) { row = r; break; }
  }
  if (!row) throw new Error('Intervenant « ' + p.nom + ' » introuvable dans le tableau.');
  var ON = function (v) { return v ? 'oui' : 'non'; };
  var champs = [
    [IV_ACTIVITE, ON(p.payeActivite)], [IV_ADMIN, ON(p.payeAdmin)], [IV_TECH, ON(p.payeTech)],
    [IV_H_ADMIN, num_(p.hAdmin)], [IV_H_TECH, num_(p.hTech)], [IV_H_ACTIVITE, num_(p.hActivite)],
    [IV_OBJECTIF, num_(p.objectif)], [IV_MOIS, num_(p.mois) || 12]
  ];
  champs.forEach(function (c) { shP.getRange(row, c[0]).setValue(c[1]); });
  SpreadsheetApp.flush();
  logAction_('remuneration', p.nom + ' activite=' + ON(p.payeActivite) + ' admin=' + ON(p.payeAdmin) + ' tech=' + ON(p.payeTech) +
    ' hAdmin=' + num_(p.hAdmin) + ' hTech=' + num_(p.hTech) + ' hActivite=' + num_(p.hActivite));
  return true;
}

// ---------- affectations, pour la facturation ----------

function ecoTarifs_() {
  var sh = ecoSS_().getSheetByName('Parametres');
  if (!sh) return {};
  var val = function (lib, def) {
    var r = ecoLigneExacte_(sh, lib);
    return r ? num_(sh.getRange(r, 2).getValue()) : def;
  };
  return {
    cours: val('Cours annuel - forfait par seance et par intervenant', 30),
    collectivite: val('Seance collectivite sur site - par intervenant', 30),
    demiJournee: val('Demi-journee exterieure 3h - par intervenant', 90),
    journee: val('Journee exterieure 6h - par intervenant', 180),
    stage: val('Stage - par jour et par intervenant', 150),
    admin: val('Remuneration Admin / IT', 30),
    technique: val('Remuneration Technique', 30),
    heureParSeance: val('1 seance = nbre heure', 2),
    nbInterv: val('Nombre d intervenants par seance', 2)
  };
}

function ecoLignesOnglet_(ss, onglet, entete, colUnites, colLibelle) {
  var sh = ss.getSheetByName(onglet);
  if (!sh) return [];
  var e = ecoEntete_(sh, entete);
  if (!e) return [];
  var hr = e.row, bornes = ecoFinTableau_(sh, hr);
  var cInt1 = ecoColonne_(sh, hr, 'Intervenant 1');
  if (!cInt1) return [];
  var cU = ecoColonne_(sh, hr, colUnites);
  var cT = ecoColonne_(sh, hr, 'Tarif par intervenant') || ecoColonne_(sh, hr, 'Cout unitaire par intervenant');
  var cL = colLibelle ? ecoColonne_(sh, hr, colLibelle) : 1;
  var out = [];
  for (var r = hr + 1; r <= bornes.last; r++) {
    var gens = [];
    for (var k = 0; k < ECO_COLS_INTERV; k++) {
      var n = String(sh.getRange(r, cInt1 + k).getValue() || '').trim();
      if (n) gens.push(n);
    }
    if (!gens.length) continue;
    out.push({
      libelle: String(sh.getRange(r, cL).getValue() || '').trim(),
      detail: String(sh.getRange(r, 2).getValue() || '').trim(),
      unites: num_(sh.getRange(r, cU).getValue()),
      tarif: cT ? num_(sh.getRange(r, cT).getValue()) : 0,
      gens: gens
    });
  }
  return out;
}

function ecoAffectations_() {
  var ss = ecoSS_();
  return {
    tarifs: ecoTarifs_(),
    cours: ecoLignesOnglet_(ss, 'Cours_annuels', 'Cours', 'Seances/an'),
    stages: ecoLignesOnglet_(ss, 'Stages', 'Stage', 'Jours'),
    presta: ecoLignesOnglet_(ss, 'Prestations', 'Client', 'Quantite')
  };
}

/* ===== FacturesInterv.gs ===== */
/**
 * FacturesInterv.gs : les intervenants (auto-entrepreneurs) facturent l'association.
 *
 * L'intervenant remplit sa fiche UNE fois (onglet Intervenants), puis chaque facture :
 *  - propose les lignes à partir du modèle économique (cours, stages, prestations, admin, technique)
 *    sur la période choisie, toutes modifiables,
 *  - génère un Doc + un PDF nommés FAC001-NOM-Periode, rangés dans « Factures intervenants »,
 *  - crée la ligne « à payer » correspondante dans l'onglet Achats.
 *
 * L'émetteur est l'intervenant, le destinataire est l'association.
 */

// ---------- fiches intervenants ----------

function intervSheet_() {
  var ss = erpSS_();
  var sh = ss.getSheetByName('Intervenants');
  if (!sh) { sh = ss.insertSheet('Intervenants'); sh.getRange(1, 1, 1, TABLES.Intervenants.length).setValues([TABLES.Intervenants]); }
  return sh;
}

function apiIntervenants_() {
  intervSheet_();
  var l = readTable_('Intervenants').filter(function (r) { return String(r.Nom || '').trim(); });
  l.forEach(function (r) {
    r.complet = !!(String(r.Adresse || '').trim() && String(r.Ville || '').trim() && String(r.SIRET || '').trim() && String(r.Email || '').trim());
    r.prochain = 'FAC' + numInterv_(prochainNumInterv_(r));
  });
  return { liste: l, ecole: enteteEcole_() };
}

function numInterv_(n) { var s = String(Math.max(1, Math.round(n))); while (s.length < 3) s = '0' + s; return s; }

/**
 * Prochain numero de facture d une personne.
 * On regarde d abord les factures qui existent vraiment : le plus grand numero deja emis + 1,
 * et 1 s il n y en a aucune. Comme ca, apres des essais supprimes, la numerotation repart
 * d elle-meme au bon endroit au lieu de garder le compteur gonfle.
 * Le compteur enregistre (ProchainNum) n est plus incremente tout seul : il ne sert que de
 * plancher, quand on force un depart depuis le bouton de renumerotation (numeros deja
 * utilises hors ERP par exemple). Un plancher plus bas qu un numero deja emis est ignore.
 */
function prochainNumInterv_(iv) {
  var maxi = 0;
  try {
    readTable_('FacturesInterv').forEach(function (r) {
      if (!memeNom_(r.Intervenant, iv.Nom)) return;
      var m = /(\d+)\s*$/.exec(String(r.Numero || ''));
      if (m) maxi = Math.max(maxi, parseInt(m[1], 10));
    });
  } catch (e) { }
  return Math.max(maxi + 1, num_(iv.ProchainNum) || 1);
}

/** Remet la numerotation d une personne sur les factures reellement presentes. p = { id, depart } */
function apiReinitNumInterv_(p) {
  var iv = readTable_('Intervenants').filter(function (r) { return r.ID === p.id; })[0];
  if (!iv) throw new Error('Profil de facturation introuvable.');
  var force = num_(p.depart);
  var n = force > 0 ? Math.round(force) : prochainNumInterv_({ Nom: iv.Nom, ProchainNum: 0 });
  upsertRow_('Intervenants', { ID: iv.ID, ProchainNum: n }, 'ID');
  logAction_('numerotation facture', iv.Nom + ' -> FAC' + numInterv_(n));
  return { id: iv.ID, prochain: 'FAC' + numInterv_(n), numero: n };
}

function enteteEcole_() {
  var c = getConfig_();
  return {
    nom: c.NOM_STRUCTURE || "",
    structure: c.NOM_ASSO || 'Association',
    adresse: c.ADRESSE_STRUCTURE || '',
    siret: c.SIRET_ASSO || '',
    email: c.EMAIL_STRUCTURE || '',
    tel: c.TEL_STRUCTURE || ''
  };
}

function apiSaveIntervenant_(p) {
  intervSheet_();
  if (!String(p.Nom || '').trim()) throw new Error('Nom obligatoire.');
  var ex = readTable_('Intervenants').filter(function (r) { return String(r.ID) === String(p.ID) || norm_(r.Nom) === norm_(p.Nom); })[0] || {};
  var rec = {
    ID: p.ID || ex.ID || uid_(),
    Nom: String(p.Nom).trim(),
    Statut: p.Statut || ex.Statut || 'Auto-entrepreneur',
    Adresse: p.Adresse !== undefined ? p.Adresse : (ex.Adresse || ''),
    CP: p.CP !== undefined ? p.CP : (ex.CP || ''),
    Ville: p.Ville !== undefined ? p.Ville : (ex.Ville || ''),
    SIRET: p.SIRET !== undefined ? p.SIRET : (ex.SIRET || ''),
    Email: p.Email !== undefined ? p.Email : (ex.Email || ''),
    Tel: p.Tel !== undefined ? p.Tel : (ex.Tel || ''),
    IBAN: p.IBAN !== undefined ? p.IBAN : (ex.IBAN || ''),
    BIC: p.BIC !== undefined ? p.BIC : (ex.BIC || ''),
    MentionTVA: p.MentionTVA || ex.MentionTVA || (getConfig_().MENTION_TVA || 'TVA non applicable, article 293 B du CGI'),
    ProchainNum: num_(p.ProchainNum) || num_(ex.ProchainNum) || 1,
    Actif: p.Actif || ex.Actif || 'oui',
    MajPar: currentUserLabel_(), MajLe: new Date()
  };
  upsertRow_('Intervenants', rec, 'ID');
  logAction_('fiche intervenant', rec.Nom);
  return rec;
}

// ---------- proposition de lignes à partir du modèle économique ----------

function moisFr_(d) {
  return ['janvier', 'fevrier', 'mars', 'avril', 'mai', 'juin', 'juillet', 'aout', 'septembre', 'octobre', 'novembre', 'decembre'][d.getMonth()];
}
function majuscule_(s) { return s ? s.charAt(0).toUpperCase() + s.slice(1) : s; }

/** « Octobre 2026 » ou « Septembre a decembre 2026 » */
function libellePeriode_(d1, d2) {
  if (!d1) return '';
  if (!d2 || (d1.getMonth() === d2.getMonth() && d1.getFullYear() === d2.getFullYear())) return majuscule_(moisFr_(d1)) + ' ' + d1.getFullYear();
  if (d1.getFullYear() === d2.getFullYear()) return majuscule_(moisFr_(d1)) + ' a ' + moisFr_(d2) + ' ' + d2.getFullYear();
  return majuscule_(moisFr_(d1)) + ' ' + d1.getFullYear() + ' a ' + moisFr_(d2) + ' ' + d2.getFullYear();
}

function nbMoisPeriode_(d1, d2) {
  if (!d1) return 1;
  if (!d2) return 1;
  return Math.max(1, (d2.getFullYear() - d1.getFullYear()) * 12 + (d2.getMonth() - d1.getMonth()) + 1);
}

/** Durée d'un créneau « 17h00-18h30 » en heures */
function dureeHoraire_(h) {
  var m = String(h || '').match(/(\d{1,2})\s*h\s*(\d{2})?\s*[-–à]\s*(\d{1,2})\s*h\s*(\d{2})?/i);
  if (!m) return 0;
  return ((+m[3] * 60 + (+m[4] || 0)) - (+m[1] * 60 + (+m[2] || 0))) / 60;
}

/**
 * p = { nom, debut, fin, mode: 'reel' | 'lisse' }
 * Renvoie des lignes { designation, detail, qte, pu } modifiables avant génération.
 * C'est une AIDE A LA SAISIE : quoi qu'il arrive elle rend une proposition utilisable.
 * Si le modele economique est illisible ou si la personne n'y est pas, on renvoie une ligne
 * vide et un avertissement affiche dans le tiroir, au lieu d'une erreur rouge qui bloque
 * la creation de la facture.
 */
function apiPropositionFacture_(p) {
  try {
    return propositionFacture_(p);
  } catch (e) {
    var d = parseDate_(p && p.debut), f = parseDate_(p && p.fin) || d;
    return {
      lignes: [{ designation: '', detail: '', qte: 1, pu: 0 }],
      detail: [], compteur: null, fiche: null,
      avertissement: 'Proposition automatique impossible : ' + ((e && e.message) || e) + '. A saisir a la main.',
      periode: d ? libellePeriode_(new Date(d.getFullYear(), d.getMonth(), 1), f) : '',
      debut: d ? fmtDate_(d) : '', fin: f ? fmtDate_(f) : '', nbMois: 1
    };
  }
}

function propositionFacture_(p) {
  var d1 = parseDate_(p.debut), d2 = parseDate_(p.fin) || d1;
  if (!d1) throw new Error('Choisis au moins un mois.');
  var fin = new Date(d2.getFullYear(), d2.getMonth() + 1, 0);   // fin du mois de d2
  var deb = new Date(d1.getFullYear(), d1.getMonth(), 1);
  var nom = String(p.nom || '').trim();
  var A = {};
  try { A = ecoAffectations_(); } catch (e) { A = { tarifs: {}, cours: [], stages: [], presta: [] }; }
  var T = A.tarifs || {};
  var lignes = [];
  var cours = coursList_(), stages = stagesList_();
  var nbMois = nbMoisPeriode_(deb, d2);

  var vol = volumesIntervenant_(nom);
  var payeActivite = vol.payeActivite !== false;

  // cours annuels : nombre de séances réellement prévues dans la période
  var totSeances = 0, totHeures = 0, tarifCours = T.cours || 30, detailCours = [];
  (A.cours || []).forEach(function (l) {
    if (!l.gens.some(function (g) { return memeNom_(g, nom); })) return;
    var lab = norm_(l.libelle);
    var k = cours.filter(function (x) { return lab.indexOf(norm_(x.cle)) === 0; })[0];
    if (!k) return;
    var dates = datesSeances_(k.jour, Math.max(num_(l.unites), 0));
    var n = dates.filter(function (d) { return d >= deb && d <= fin; }).length;
    if (!n) return;
    totSeances += n;
    totHeures += n * (dureeHoraire_(k.horaire) || 0);
    if (l.tarif) tarifCours = l.tarif;
    detailCours.push(k.jour + ' ' + k.horaire + ' x' + n);
  });
  if (totSeances && payeActivite) lignes.push(ligneNature_('cours',
    detailCours.length + ' créneau(x), ' + totSeances + ' séances' + (totHeures ? ', ' + arrondi_(totHeures) + ' h' : '') + '.',
    totSeances, tarifCours));

  // stages : les semaines dont la date de début tombe dans la période
  var totJours = 0, tarifStage = T.stage || 150, detailStages = [];
  (A.stages || []).forEach(function (l) {
    if (!l.gens.some(function (g) { return memeNom_(g, nom); })) return;
    var st = stages.filter(function (s) { return norm_(s.nom) === norm_(l.libelle); })[0];
    var d = st ? parseDate_(st.debut) : null;
    if (!d || d < deb || d > fin) return;
    var j = num_(l.unites);
    if (!j) return;
    totJours += j;
    if (l.tarif) tarifStage = l.tarif;
    detailStages.push(l.libelle + ' (' + j + ' j)');
  });
  if (totJours && payeActivite) lignes.push(ligneNature_('stages', detailStages.join(', ') + '.', totJours, tarifStage));

  // prestations : pas de date dans le modèle, on propose la ligne à 0, à compléter
  var presta = (A.presta || []).filter(function (l) { return l.gens.some(function (g) { return memeNom_(g, nom); }) && num_(l.unites) > 0; });
  if (presta.length && payeActivite) lignes.push(ligneNature_('presta',
    'Quantité à ajuster. Prestations concernées : ' + presta.map(function (l) { return l.libelle; }).join(', ') + '.',
    0, presta[0].tarif || T.collectivite || 30));

  // admin et technique : séances par mois x nombre de mois de la période
  if (vol.activitePlus) lignes.push(ligneNature_('activitePlus',
    vol.activitePlus + ' séance(s) de 2 h sur l\'année.', vol.activitePlus, tarifCours));
  if (vol.admin) lignes.push(ligneNature_('admin',
    vol.admin + ' séance(s) de 2 h par mois sur ' + nbMois + ' mois.', vol.admin * nbMois, T.admin || 30));
  if (vol.technique) lignes.push(ligneNature_('technique',
    vol.technique + ' séance(s) de 2 h par mois sur ' + nbMois + ' mois.', vol.technique * nbMois, T.technique || 30));

  // Mensualite lissee : ce qui a ete fait reste affiche en detail, mais le montant facture
  // est la mensualite du forfait annuel. Un compteur suit l'ecart realise / facture.
  var detail = [], compteur = null, fiche = null, avertissement = '';
  if (p.mode !== 'reel') {
    var moi = null, dispo = [];
    try {
      var R = apiRemunerations_();
      dispo = (R.lignes || []).map(function (x) { return x.nom; });
      moi = (R.lignes || []).filter(function (x) { return memeNom_(x.nom, nom); })[0] || null;
    } catch (e) { avertissement = 'Modele economique illisible : ' + e.message; }
    // sans correspondance on produirait une facture a 0 EUR sans rien dire : on le signale
    if (!moi && !avertissement) {
      avertissement = 'Aucune fiche « ' + nom + ' » dans le modele economique' +
        (dispo.length ? ' (il contient : ' + dispo.join(', ') + ')' : '') +
        '. Mets le meme nom dans le tableau INTERVENANTS de l onglet Parametres, ou dans la fiche de facturation.';
    } else if (moi && !(moi.natures || []).some(function (n) { return n.paye && num_(n.mensuel); })) {
      avertissement = 'La fiche « ' + moi.nom + ' » n a aucune nature activee et payee : rien a facturer. ' +
        'Active activite, admin ou technique dans Compta > Remuneration.';
    }
    var annuel = moi ? num_(moi.total) : 0;
    var mois = (moi && moi.mois) || vol.moisFactures || 12;
    var mensualite = mois ? arrondi_(annuel / mois) : 0;
    var rang = rangMensualite_(deb, mois);
    var libRang = (nbMois > 1 ? rang + ' à ' + Math.min(rang + nbMois - 1, mois) : String(rang)) + '/' + mois;
    detail = lignes;
    compteur = compteurIntervenant_(nom, fin, annuel, mois);
    compteur.facture = arrondi_(compteur.facture + mensualite * nbMois);
    compteur.ecart = arrondi_(compteur.realise - compteur.facture);
    lignes = lignesLissees_(moi, mois, nbMois, libRang, annuel, mensualite);
    fiche = moi ? {
      nom: moi.nom, mois: mois, annuel: arrondi_(annuel), mensuel: mensualite, rang: libRang,
      natures: (moi.natures || []).map(function (n) {
        return { libelle: n.libelle, paye: n.paye, volume: n.volume, annuel: n.annuel, mensuel: n.mensuel };
      })
    } : null;
  }

  return { lignes: lignes, detail: detail, compteur: compteur, fiche: fiche, avertissement: avertissement,
    periode: libellePeriode_(deb, d2), debut: fmtDate_(deb), fin: fmtDate_(fin), nbMois: nbMois };
}

/**
 * Une ligne de facture, descriptif officiel de la nature en tete puis le detail chiffre.
 * Les deux modes (reel et lisse) emploient les memes designations, pour que deux factures
 * de la meme personne se lisent pareil.
 */
function ligneNature_(cle, detail, qte, pu) {
  var m = ECO_NATURE_FACTURE[cle] || { designation: cle, detail: '' };
  // « chiffre » garde la partie quantitative seule : l annexe la reprend sans redire le descriptif
  return { designation: m.designation, detail: (m.detail + ' ' + (detail || '')).trim(), chiffre: detail || '', qte: qte, pu: pu };
}

/**
 * Une ligne de facture par nature active de la fiche de l intervenant dans le modele eco.
 * Chaque ligne porte sa mensualite lissee, la somme vaut exactement la mensualite globale
 * (le dernier centime d arrondi est repris sur la derniere ligne).
 */
function lignesLissees_(moi, mois, nbMois, libRang, annuel, mensualite) {
  var nat = ((moi && moi.natures) || []).filter(function (n) { return n.paye && num_(n.mensuel); });
  if (!nat.length) {
    return [{
      designation: 'Prestation de services',
      detail: 'Prestations de terrain, administratives et techniques réalisées pour le compte de l\'association.',
      qte: nbMois, pu: mensualite
    }];
  }
  var out = nat.map(function (n) {
    var m = ECO_NATURE_FACTURE[n.cle] || { designation: n.libelle, detail: '' };
    // la facture dit ce qui est vendu. Le forfait annuel, le lissage et le rang de mensualité
    // sont des mécaniques internes : leur place est dans la convention, pas sur la facture.
    return { designation: m.designation, detail: m.detail, qte: nbMois, pu: arrondi_(n.mensuel) };
  });
  var cible = arrondi_(mensualite * nbMois);
  var somme = arrondi_(out.reduce(function (t, l) { return t + l.qte * l.pu; }, 0));
  if (somme !== cible) {
    var der = out[out.length - 1];
    der.pu = arrondi_(der.pu + (cible - somme) / der.qte);
  }
  return out;
}

function arrondi_(n) { return Math.round(num_(n) * 100) / 100; }

/** « Camille » et « Dupont Camille » sont la même personne (fiche ERP vs modele eco) */
function memeNom_(a, b) {
  var x = norm_(a), y = norm_(b);
  if (!x || !y) return false;
  if (x === y || x.indexOf(y) === 0 || y.indexOf(x) === 0) return true;
  // « Camille » (modele eco) et « DUPONT Camille » (fiche ERP) : l ordre nom / prenom
  // change d un endroit a l autre. On compare mot a mot : le nom le plus court doit etre
  // entierement contenu dans l autre, et au moins un mot doit faire 3 lettres ou plus.
  var mots = function (s) { return s.split(/[^a-z0-9]+/).filter(function (m) { return m.length > 1; }); };
  var mx = mots(x), my = mots(y);
  if (!mx.length || !my.length) return false;
  var court = mx.length <= my.length ? mx : my, long = mx.length <= my.length ? my : mx;
  var solide = false;
  for (var i = 0; i < court.length; i++) {
    if (long.indexOf(court[i]) < 0) return false;
    if (court[i].length >= 3) solide = true;
  }
  return solide;
}

/** Rang de la mensualite dans la saison : septembre = 1 */
function rangMensualite_(debutPeriode, mois) {
  var s = parseDate_(getConfig_().DEBUT_SAISON_COURS);
  if (!s || !debutPeriode) return 1;
  var n = (debutPeriode.getFullYear() - s.getFullYear()) * 12 + (debutPeriode.getMonth() - s.getMonth()) + 1;
  return Math.min(Math.max(n, 1), mois || 12);
}

/**
 * Compteur de regularisation : ce qui a ete reellement fait depuis le debut de saison,
 * face a ce qui a deja ete facture. L'ecart se solde en fin d'annee.
 */
function compteurIntervenant_(nom, fin, annuel, mois) {
  var realise = 0;
  try {
    var s = parseDate_(getConfig_().DEBUT_SAISON_COURS) || new Date();
    var p = propositionFacture_({ nom: nom, debut: fmtDate_(s), fin: fmtDate_(fin), mode: 'reel' });
    realise = p.lignes.reduce(function (t, l) { return t + num_(l.qte) * num_(l.pu); }, 0);
  } catch (e) { }
  var facture = 0;
  try {
    readTable_('FacturesInterv').forEach(function (r) { if (memeNom_(r.Intervenant, nom)) facture += num_(r.Total); });
  } catch (e) { }
  return { realise: arrondi_(realise), facture: arrondi_(facture), ecart: arrondi_(realise - facture), annuel: arrondi_(annuel || 0), mois: mois || 12 };
}

/**
 * Volumes d'une personne, lus dans le tableau INTERVENANTS de Parametres.
 * Le tableau est en HEURES par mois, la facture raisonne en seances (1 seance = 2 h par defaut).
 * Une nature non activee renvoie 0 : on ne facture pas ce qui n'est pas paye.
 */
function volumesIntervenant_(nom) {
  try {
    var R = apiRemunerations_();
    var g = (R.lignes || []).filter(function (x) { return memeNom_(x.nom, nom); })[0];
    if (!g) return {};
    var h = num_((ecoTarifs_() || {}).heureParSeance) || 2;
    return {
      admin: g.payeAdmin ? Math.round(g.hAdmin / h * 100) / 100 : 0,
      technique: g.payeTech ? Math.round(g.hTech / h * 100) / 100 : 0,
      activitePlus: g.payeActivite ? Math.round(g.hActivite / h * 100) / 100 : 0,
      payeActivite: g.payeActivite,
      moisFactures: g.mois || 12
    };
  } catch (e) { }
  return {};
}

// ---------- génération du PDF ----------

function dossierFactInterv_() {
  var c = getConfig_();
  var root = c.ID_DOSSIER_DEVIS ? DriveApp.getFolderById(c.ID_DOSSIER_DEVIS) : DriveApp.getRootFolder();
  var dossier = getOrCreateFolder_(root, 'Factures intervenants');
  return { root: dossier, docs: getOrCreateFolder_(dossier, 'Modifiables (Docs)') };
}

function sansAccent_(s) {
  return String(s || '').normalize('NFD').replace(/[̀-ͯ]/g, '').replace(/[^A-Za-z0-9]+/g, '-').replace(/^-+|-+$/g, '');
}

/** p = { intervenantId, debut, fin, periode, lignes:[{designation,detail,qte,pu}], notes } */
function apiCreerFactureInterv_(p) {
  var lock = LockService.getScriptLock(); lock.waitLock(30000);
  try {
    intervSheet_();
    var iv = readTable_('Intervenants').filter(function (r) { return String(r.ID) === String(p.intervenantId); })[0];
    if (!iv) throw new Error('Fiche intervenant introuvable.');
    ['Adresse', 'Ville', 'SIRET', 'Email'].forEach(function (k) {
      if (!String(iv[k] || '').trim()) throw new Error('Complete d\'abord ta fiche : il manque ' + k.toLowerCase() + '.');
    });
    var lignes = (p.lignes || []).map(function (l) {
      return { designation: String(l.designation || '').trim(), detail: String(l.detail || '').trim(), qte: num_(l.qte), pu: num_(l.pu) };
    }).filter(function (l) { return l.designation && (l.qte * l.pu) !== 0; });
    if (!lignes.length) throw new Error('Aucune ligne a facturer.');
    var total = arrondi_(lignes.reduce(function (t, l) { return t + l.qte * l.pu; }, 0));

    var c = getConfig_();
    var rang = prochainNumInterv_(iv);
    var num = 'FAC' + numInterv_(rang);
    var periode = String(p.periode || libellePeriode_(parseDate_(p.debut), parseDate_(p.fin))).trim();
    var dateFact = parseDate_(p.date) || new Date();
    var fichier = num + '-' + sansAccent_(iv.Nom).toUpperCase() + '-' + sansAccent_(periode);

    var doss = dossierFactInterv_();
    var doc = DocumentApp.create(fichier);
    var detail = (p.detail || []).map(function (l) {
      return { designation: String(l.designation || '').trim(), detail: String(l.detail || '').trim(),
        chiffre: String(l.chiffre || '').trim(), qte: num_(l.qte), pu: num_(l.pu) };
    }).filter(function (l) { return l.designation; });
    var delai = num_(c.DELAI_PAIEMENT_INTERV) || 90;
    ecrireFactureInterv_(doc, { iv: iv, num: num, date: dateFact, periode: periode, lignes: lignes, total: total, notes: p.notes || '',
      convention: c.DATE_CONVENTION || '',
      echeance: new Date(dateFact.getFullYear(), dateFact.getMonth(), dateFact.getDate() + delai),
      ecole: enteteEcole_(), delai: num_(c.DELAI_PAIEMENT_INTERV) || 90, penalites: c.PENALITES_INTERV || '' });
    doc.saveAndClose();

    var fDoc = DriveApp.getFileById(doc.getId());
    doss.docs.addFile(fDoc);
    try { DriveApp.getRootFolder().removeFile(fDoc); } catch (e) { }
    var pdf = doss.root.createFile(fDoc.getAs('application/pdf').setName(fichier + '.pdf'));

    var rec = {
      ID: uid_(), Numero: num, Date: dateFact, Intervenant: iv.Nom,
      PeriodeDebut: parseDate_(p.debut) || '', PeriodeFin: parseDate_(p.fin) || '', Periode: periode,
      Lignes: JSON.stringify(lignes).slice(0, 45000), Total: total,
      Fichier: fichier + '.pdf', LienPdf: pdf.getUrl(), LienDoc: fDoc.getUrl(),
      Statut: 'Emise', AchatID: '', MajPar: currentUserLabel_(), MajLe: new Date()
    };

    // la facture arrive tout de suite dans les factures à payer de la compta
    try {
      var ech = new Date(dateFact.getFullYear(), dateFact.getMonth(), dateFact.getDate() + (num_(c.DELAI_PAIEMENT_INTERV) || 90));
      rec.AchatID = apiSaveAchat_({
        Beneficiaire: iv.Nom, Type: norm_(iv.Statut).indexOf('auto') === 0 ? 'Auto-entreprise' : 'Intervenant',
        Objet: num + ' - ' + periode, Montant: total, Date: dateFact, Echeance: ech, Lien: pdf.getUrl(), Notes: 'Facture générée depuis l\'ERP'
      });
    } catch (e) { rec.Notes = 'Ligne a payer non creee : ' + e.message; }

    upsertRow_('FacturesInterv', rec, 'ID');
    logAction_('facture intervenant', num + ' ' + iv.Nom + ' ' + total + ' EUR');
    return { id: rec.ID, numero: num, fichier: rec.Fichier, pdf: rec.LienPdf, doc: rec.LienDoc, total: total, achat: !!rec.AchatID };
  } finally { lock.releaseLock(); }
}

/**
 * Rendu du document : facture d'auto-entrepreneur, modele standard.
 * Emetteur a gauche, client a droite, bloc FACTURE avec les trois dates, un tableau
 * Designation / Prix unitaire / Quantite / Prix total, le total detache en bas a droite,
 * puis les mentions legales. Tout tient sur une page.
 */
var FI_LARG = [262, 86, 66, 86];              // points, A4 avec 42 pt de marges
var FI_GRIS = '#6b6259', FI_TRAIT = '#d9d2c7', FI_FOND = '#f4c20d';

/** Style d'une cellule : alignement, fonds, et descriptif en petit sous la designation */
function fiCellule_(cell, opt) {
  opt = opt || {};
  try {
    var A = DocumentApp.Attribute, attr = {};
    attr[A.PADDING_TOP] = 5; attr[A.PADDING_BOTTOM] = 5; attr[A.PADDING_LEFT] = 7; attr[A.PADDING_RIGHT] = 7;
    if (opt.fond) attr[A.BACKGROUND_COLOR] = opt.fond;
    if (opt.droite) attr[A.HORIZONTAL_ALIGNMENT] = DocumentApp.HorizontalAlignment.RIGHT;
    else if (opt.centre) attr[A.HORIZONTAL_ALIGNMENT] = DocumentApp.HorizontalAlignment.CENTER;
    cell.setAttributes(attr);
  } catch (e) { }
  try {
    for (var i = 0; i < cell.getNumChildren(); i++) {
      var ch = cell.getChild(i);
      if (!ch.editAsText) continue;
      var petit = i > 0;
      ch.editAsText().setFontSize(petit ? 7.5 : (opt.taille || 9))
        .setBold(!petit && !!opt.gras)
        .setForegroundColor(petit ? FI_GRIS : '#000000');
    }
  } catch (e) { }
}

function ecrireFactureInterv_(doc, o) {
  var body = doc.getBody();
  var iv = o.iv;
  var st = function (p, taille, gras, couleur) {
    try {
      var t = p.editAsText().setFontSize(taille || 9).setBold(!!gras);
      if (couleur) t.setForegroundColor(couleur);
    } catch (e) { }
    return p;
  };
  try { body.setMarginTop(44).setMarginBottom(36).setMarginLeft(42).setMarginRight(42); } catch (e) { }

  // ----- émetteur / client, côte à côte, sans bordures -----
  var emet = [iv.Nom, iv.Statut || 'Auto-entrepreneur', iv.Adresse, [iv.CP, iv.Ville].filter(String).join(' '),
    iv.SIRET ? 'SIRET ' + iv.SIRET : '', iv.Tel, iv.Email].filter(String).join('\n');
  var dest = [o.structure.structure, o.structure.nom, o.structure.adresse,
    o.structure.siret ? 'SIRET ' + o.structure.siret : '', o.structure.email].filter(String).join('\n');
  var tEnt = body.appendTable([['ÉMETTEUR\n' + emet, 'FACTURÉ À\n' + dest]]);
  try {
    tEnt.setBorderWidth(0);
    tEnt.setColumnWidth(0, 250); tEnt.setColumnWidth(1, 250);
    [0, 1].forEach(function (c) {
      var cell = tEnt.getRow(0).getCell(c);
      cell.editAsText().setFontSize(8);
      cell.getChild(0).asParagraph().editAsText().setFontSize(7).setBold(true).setForegroundColor(FI_GRIS);
      if (cell.getNumChildren() > 1) cell.getChild(1).asParagraph().editAsText().setFontSize(9.5).setBold(true);
    });
  } catch (e) { }

  // ----- bloc facture : numéro et les trois dates -----
  body.appendParagraph('');
  st(body.appendParagraph('FACTURE N° ' + o.num), 15, true);
  var dates = ['Date d\'émission : ' + fmtDate_(o.date), 'Période d\'exécution : ' + o.periode];
  if (o.echeance) dates.push('Date d\'échéance : ' + fmtDate_(o.echeance));
  st(body.appendParagraph(dates.join('\n')), 8.5, false, FI_GRIS);
  body.appendParagraph('');

  // ----- prestations -----
  var t = [['Désignation', 'Prix unitaire', 'Quantité', 'Prix total']];
  o.lignes.forEach(function (l) {
    t.push([l.designation + (l.detail ? '\n' + l.detail : ''), eur_(l.pu), String(l.qte), eur_(arrondi_(l.qte * l.pu))]);
  });
  t.push(['', '', 'Total', eur_(o.total)]);
  var table = body.appendTable(t);
  try {
    table.setBorderColor(FI_TRAIT);
    for (var c = 0; c < FI_LARG.length; c++) table.setColumnWidth(c, FI_LARG[c]);
    for (var r = 0; r < t.length; r++) {
      var row = table.getRow(r), entete = r === 0, total = r === t.length - 1;
      var fond = entete ? FI_FOND : null;
      fiCellule_(row.getCell(0), { gras: entete, fond: fond });
      fiCellule_(row.getCell(1), { droite: true, gras: entete, fond: fond });
      fiCellule_(row.getCell(2), { centre: !total, droite: total, gras: entete || total, fond: total ? FI_FOND : fond });
      fiCellule_(row.getCell(3), { droite: true, gras: entete || total, taille: total ? 10.5 : 9, fond: total ? FI_FOND : fond });
    }
  } catch (e) { }

  // ----- mentions légales -----
  body.appendParagraph('');
  st(body.appendParagraph(iv.MentionTVA || 'TVA non applicable, article 293 B du code général des impôts'), 9, true);
  var bas = ['Paiement à ' + o.delai + ' jours à compter de la date de facture, par virement bancaire.'];
  if (iv.IBAN) bas.push('IBAN ' + iv.IBAN + (iv.BIC ? '   ·   BIC ' + iv.BIC : ''));
  st(body.appendParagraph(bas.join('\n')), 8.5);
  var fin = [];
  if (o.penalites) fin.push(o.penalites);
  if (iv.Assurance) fin.push('Assurance responsabilité civile professionnelle : ' + iv.Assurance + '.');
  if (o.convention) fin.push('Établie au titre de la convention annuelle de prestation de services du ' + o.convention + '.');
  if (fin.length) st(body.appendParagraph(fin.join('\n')), 7, false, FI_GRIS);
  if (o.notes) st(body.appendParagraph(o.notes), 8.5);
}

// ---------- liste, envoi, suppression ----------

function apiFacturesInterv_(p) {
  var ss = erpSS_();
  if (!ss.getSheetByName('FacturesInterv')) { var s = ss.insertSheet('FacturesInterv'); s.getRange(1, 1, 1, TABLES.FacturesInterv.length).setValues([TABLES.FacturesInterv]); }
  var l = readTable_('FacturesInterv');
  if (p && p.nom) l = l.filter(function (r) { return norm_(r.Intervenant) === norm_(p.nom); });
  return l.sort(function (a, b) { return String(b.Numero).localeCompare(String(a.Numero)); });
}

/** Envoi du PDF par mail à l'intervenant (et à l'école si demandé) */
function apiEnvoyerFactureInterv_(p) {
  var f = readTable_('FacturesInterv').filter(function (r) { return String(r.ID) === String(p.id); })[0];
  if (!f) throw new Error('Facture introuvable.');
  var iv = readTable_('Intervenants').filter(function (r) { return norm_(r.Nom) === norm_(f.Intervenant); })[0] || {};
  var to = String(p.to || iv.Email || '').trim();
  if (!to) throw new Error('Pas d\'adresse mail pour ' + f.Intervenant + '.');
  var c = getConfig_();
  var pdf = null;
  try {
    var id = String(f.LienPdf).match(/[-\w]{25,}/);
    if (id) pdf = DriveApp.getFileById(id[0]).getBlob();
  } catch (e) { }
  var corps = 'Bonjour,\n\nCi-joint la facture ' + f.Numero + ' (' + f.Periode + '), ' + eur_(num_(f.Total)) + '.\n\n' +
    (f.LienPdf ? f.LienPdf + '\n\n' : '') + 'Envoyee depuis l\'ERP de ' + (c.NOM_STRUCTURE || 'l\'ecole') + '.';
  var opts = { name: c.NOM_STRUCTURE || 'ERP', body: corps };
  if (pdf) opts.attachments = [pdf];
  var dest = [to].concat(p.copieEcole && c.EMAIL_STRUCTURE ? [c.EMAIL_STRUCTURE] : []).join(',');
  MailApp.sendEmail(Object.assign({ to: dest, subject: f.Numero + ' - ' + f.Intervenant + ' - ' + f.Periode }, opts));
  upsertRow_('FacturesInterv', { ID: f.ID, Statut: 'Envoyee' }, 'ID');
  logAction_('facture intervenant envoyee', f.Numero + ' -> ' + dest);
  return true;
}

function apiSupprimerFactureInterv_(p) {
  var f = readTable_('FacturesInterv').filter(function (r) { return String(r.ID) === String(p.id); })[0];
  if (!f) throw new Error('Facture introuvable.');
  if (f.AchatID) { try { apiSupprimerAchat_({ id: f.AchatID }); } catch (e) { } }
  deleteRowById_('FacturesInterv', p.id);
  logAction_('facture intervenant supprimee', String(f.Numero));
  return true;
}

/* ===== Convention.gs ===== */
/**
 * Convention.gs : convention cadre annuelle de prestation de services, un seul document
 * pour toute l'equipe, avec les annexes chiffrees.
 *
 * A quoi ca sert : les factures des intervenants sont au forfait mensuel lisse. Une facture
 * qui dit « mensualite 2/12 du forfait annuel de 8 160 EUR » ne se suffit pas a elle-meme :
 * il faut une piece qui dise d ou vient ce forfait. C est ce document. Il est genere depuis
 * le modele economique, donc les annexes disent toujours la meme chose que les factures.
 */

/** Convention cadre + annexes, en Doc modifiable et en PDF. p = { date } */
function apiConventionCollective_(p) {
  proprietaireSeulement_();
  var c = getConfig_();
  var dateSig = parseDate_(p && p.date) || parseDate_(c.DATE_CONVENTION) || new Date();
  var ecole = enteteEcole_();
  var T = {};
  try { T = ecoTarifs_() || {}; } catch (e) { T = {}; }
  var R = { lignes: [] };
  try { R = apiRemunerations_(); } catch (e) { }

  // profils de facturation saisis dans l ERP, pour les coordonnees et les SIRET
  var fiches = {};
  try { readTable_('Intervenants').forEach(function (r) { if (r.Nom) fiches[norm_(r.Nom)] = r; }); } catch (e) { }
  var ficheDe = function (nom) {
    var k = Object.keys(fiches).filter(function (x) { return memeNom_(x, nom); })[0];
    return k ? fiches[k] : {};
  };

  var gens = (R.lignes || []).map(function (g) {
    var f = ficheDe(g.nom);
    return {
      nom: f.Nom || g.nom,
      statut: f.Statut || 'Auto-entrepreneur',
      siret: f.SIRET || '',
      adresse: [f.Adresse, [f.CP, f.Ville].filter(String).join(' ')].filter(String).join(', '),
      contact: [f.Tel, f.Email].filter(String).join(' · '),
      mois: g.mois || 12,
      annuel: g.total || 0,
      mensuel: g.mois ? Math.round(g.total / g.mois * 100) / 100 : 0,
      natures: (g.natures || []).filter(function (n) { return n.paye && num_(n.annuel); }),
      aucune: !(g.natures || []).some(function (n) { return n.paye && num_(n.annuel); })
    };
  });

  var doc = DocumentApp.create('Convention annuelle de prestation de services ' + Utilities.formatDate(dateSig, TZ, 'yyyy'));
  ecrireConvention_(doc, { date: dateSig, ecole: ecole, gens: gens, T: T,
    delai: num_(c.DELAI_PAIEMENT_INTERV) || 90, saison: c.SAISON || '2026-2027' });
  doc.saveAndClose();

  var fDoc = DriveApp.getFileById(doc.getId());
  var D = dossierFactInterv_();
  var nomFichier = 'Convention annuelle ' + Utilities.formatDate(dateSig, TZ, 'yyyy-MM-dd');
  var pdf = D.root.createFile(fDoc.getAs('application/pdf').setName(nomFichier + '.pdf'));
  try { D.docs.addFile(fDoc); DriveApp.getRootFolder().removeFile(fDoc); } catch (e) { }

  // la date de signature sert de reference sur chaque facture lissee
  try { setConfig_('DATE_CONVENTION', fmtDate_(dateSig)); } catch (e) { }
  logAction_('convention', nomFichier + ' (' + gens.length + ' intervenants)');
  return { fichier: nomFichier + '.pdf', pdf: pdf.getUrl(), doc: fDoc.getUrl(), date: fmtDate_(dateSig),
    intervenants: gens.map(function (g) { return g.nom; }) };
}

function ecrireConvention_(doc, o) {
  var body = doc.getBody();
  var st = function (p, taille, gras) { try { p.editAsText().setFontSize(taille || 9).setBold(!!gras); } catch (e) { } return p; };
  var titre = function (t) { return st(body.appendParagraph(t), 11, true); };
  var art = function (n, t) { return st(body.appendParagraph('Article ' + n + ' - ' + t), 10, true); };
  var par = function (t) { return st(body.appendParagraph(t), 9); };
  var eur = function (n) { return eur_(n); };
  try { body.setMarginTop(42).setMarginBottom(36).setMarginLeft(48).setMarginRight(48); } catch (e) { }

  st(body.appendParagraph('CONVENTION ANNUELLE DE PRESTATION DE SERVICES'), 16, true);
  st(body.appendParagraph('Saison ' + o.saison + '   ·   établie le ' + fmtDate_(o.date)), 9);
  body.appendParagraph('');

  titre('ENTRE LES SOUSSIGNÉS');
  par(o.structure.structure + (o.structure.nom && o.structure.nom !== o.structure.structure ? ', exploitant ' + o.structure.nom : '') + ',\n' +
    (o.structure.adresse ? 'dont le siège est situé ' + o.structure.adresse + ',\n' : '') +
    (o.structure.siret ? 'immatriculée sous le SIRET ' + o.structure.siret + ',\n' : '') +
    'ci-après « l\'association »,');
  par('d\'une part,');
  body.appendParagraph('');
  par('et les prestataires indépendants désignés ci-après, chacun pour ce qui le concerne :');
  var tp = [['Prestataire', 'Statut', 'SIRET', 'Adresse', 'Contact']];
  o.gens.forEach(function (g) { tp.push([g.nom, g.statut, g.siret || 'à compléter', g.adresse || 'à compléter', g.contact || '']); });
  var tabP = body.appendTable(tp);
  try { tabP.editAsText().setFontSize(8); tabP.getRow(0).editAsText().setBold(true); } catch (e) { }
  par('ci-après « les prestataires », d\'autre part.');
  body.appendParagraph('');

  art(1, 'Objet');
  par('La présente convention définit les conditions dans lesquelles chaque prestataire réalise, en toute indépendance, ' +
    'des prestations pour le compte de l\'association, et les modalités de leur rémunération. Elle est commune à l\'ensemble ' +
    'des prestataires, chacun n\'étant engagé que pour les prestations et les montants qui le concernent en annexe 2.');

  art(2, 'Nature des prestations');
  par('Les prestations relèvent de trois domaines, cumulables :');
  var tn = [['Domaine', 'Prestations couvertes']];
  [['Activite', 'Encadrement des cours annuels, des stages de vacances et des interventions pour des collectivités, écoles et structures partenaires. Préparation, animation et remise en état du matériel comprises.'],
   ['Administratif, communication et développement', 'Gestion des inscriptions et des dossiers, suivi des familles, facturation, communication et réseaux sociaux, site internet, maintenance et développement des outils de gestion.'],
   ['Technique', 'Entretien, réparation, aménagement et mise en conformité du site, du matériel, des installations électriques et des équipements de sécurité.']
  ].forEach(function (x) { tn.push(x); });
  var tabN = body.appendTable(tn);
  try { tabN.editAsText().setFontSize(8); tabN.getRow(0).editAsText().setBold(true); } catch (e) { }

  art(3, 'Indépendance des parties');
  par('Les prestataires exercent à titre indépendant. Ils organisent librement leur travail, ne sont soumis à aucun lien ' +
    'de subordination et conservent la responsabilité de leurs obligations sociales et fiscales. La présente convention ' +
    'ne crée aucun contrat de travail ni aucune exclusivité.');

  art(4, 'Rémunération et grille tarifaire');
  par('Les prestations sont rémunérées aux tarifs suivants, communs à tous les prestataires. Une séance correspond à ' +
    (num_(o.T.heureParSeance) || 2) + ' heures de travail effectif.');
  var tt = [['Prestation', 'Tarif par intervenant']];
  [['Cours annuel, la séance', o.T.cours], ['Séance collectivité sur site', o.T.collectivite],
   ['Demi-journée extérieure 3 h', o.T.demiJournee], ['Journée extérieure 6 h', o.T.journee],
   ['Stage, la journée', o.T.stage], ['Administratif, communication et développement, la séance', o.T.admin],
   ['Technique, la séance', o.T.technique]
  ].forEach(function (x) { if (num_(x[1])) tt.push([x[0], eur(x[1])]); });
  var tabT = body.appendTable(tt);
  try { tabT.editAsText().setFontSize(8); tabT.getRow(0).editAsText().setBold(true); } catch (e) { }

  art(5, 'Lissage de la rémunération');
  par('Le volume annuel de prestations de chaque prestataire est arrêté en annexe 2. Le montant correspondant est lissé ' +
    'en mensualités égales sur le nombre de mois indiqué, afin d\'assurer une régularité des versements et de la trésorerie ' +
    'de l\'association. Chaque facture mensuelle porte donc le montant de la mensualité, et non le volume réellement réalisé ' +
    'sur le mois, qui y figure à titre informatif. L\'écart entre le réalisé et le facturé est suivi en continu et régularisé ' +
    'à la dernière facture de la saison.');

  art(6, 'Facturation et paiement');
  par('Chaque prestataire adresse à l\'association une facture mensuelle, numérotée par ses soins et conforme aux mentions ' +
    'légales applicables à son statut. Le paiement intervient à ' + o.delai + ' jours à compter de la date de facture, par ' +
    'virement bancaire. Toute prestation non prévue en annexe 2 fait l\'objet d\'un accord écrit préalable et d\'une ligne ' +
    'distincte sur la facture.');

  art(7, 'Modification du volume de prestations');
  par('L\'activation, la désactivation ou la modification du volume d\'un domaine pour un prestataire donne lieu à la mise ' +
    'à jour de l\'annexe 2, datée et portée à la connaissance de l\'intéressé. La mensualité est recalculée à compter du mois ' +
    'suivant. Une annexe 2 mise à jour remplace la précédente sans qu\'il soit besoin de refaire la présente convention.');

  art(8, 'Assurance et sécurité');
  par('Chaque prestataire atteste être couvert par une assurance de responsabilité civile professionnelle en cours de ' +
    'validité et s\'engage à en fournir l\'attestation. Il respecte les règles de sécurité du lieu et l\'encadrement des ' +
    'mineurs accueillis par l\'école.');

  art(9, 'Confidentialité et données personnelles');
  par('Les prestataires sont tenus à la confidentialité sur les informations relatives aux familles, et notamment sur ' +
    'les données de santé des enfants, auxquelles ils n\'accèdent que pour les besoins de l\'encadrement. Ces données ne ' +
    'sont ni diffusées ni conservées en dehors des outils de l\'association.');

  art(10, 'Durée, résiliation et litiges');
  par('La présente convention est conclue pour la saison ' + o.saison + '. Elle peut être résiliée par l\'une ou l\'autre ' +
    'des parties, pour ce qui la concerne, par écrit et moyennant un préavis d\'un mois, les prestations réalisées restant ' +
    'dues. À défaut d\'accord amiable, les tribunaux compétents sont ceux du ressort du siège de l\'association.');

  body.appendPageBreak();
  st(body.appendParagraph('ANNEXE 1 - Justification des montants'), 13, true);
  par('Les montants de l\'annexe 2 sont calculés à partir du volume d\'activité de la saison et de la grille de l\'article 4. ' +
    'Ils sont issus du modèle économique de l\'association, qui est la source unique des chiffres portés sur les factures. ' +
    'Cette annexe est la pièce justificative des paiements mensuels : elle montre, pour chaque prestataire, quelles prestations ' +
    'composent son forfait annuel et comment on passe de ce forfait à la mensualité facturée.');
  body.appendParagraph('');

  o.gens.forEach(function (g) {
    st(body.appendParagraph(g.nom), 11, true);
    if (g.aucune) {
      par('Intervient bénévolement à ce jour. Aucun montant n\'est dû. Les prestations réalisées sont néanmoins recensées ' +
        'par l\'association, de sorte qu\'une rémunération puisse être activée en cours de saison par mise à jour de cette annexe.');
      body.appendParagraph('');
      return;
    }
    var tg = [['Domaine', 'Volume sur la saison', 'Montant annuel', 'Par mois']];
    g.natures.forEach(function (n) { tg.push([n.libelle, n.volume || '', eur(n.annuel), eur(n.mensuel)]); });
    tg.push(['TOTAL', 'lissé sur ' + g.mois + ' mois', eur(g.annuel), eur(g.mensuel)]);
    var tab = body.appendTable(tg);
    try {
      tab.editAsText().setFontSize(8);
      tab.getRow(0).editAsText().setBold(true);
      tab.getRow(tg.length - 1).editAsText().setBold(true);
    } catch (e) { }
    par('Mensualité facturée : ' + eur(g.mensuel) + ', soit ' + eur(g.annuel) + ' répartis sur ' + g.mois + ' mois.');
    body.appendParagraph('');
  });

  st(body.appendParagraph('ANNEXE 2 - Engagement par prestataire'), 13, true);
  par('Tableau récapitulatif arrêté au ' + fmtDate_(o.date) + '. Toute modification fait l\'objet d\'une annexe datée qui remplace celle-ci.');
  var ta = [['Prestataire', 'Domaines rémunérés', 'Montant annuel', 'Mois', 'Mensualité']];
  var totA = 0, totM = 0;
  o.gens.forEach(function (g) {
    totA += g.annuel; totM += g.mensuel;
    ta.push([g.nom, g.aucune ? 'bénévole, aucun montant dû' : g.natures.map(function (n) { return n.libelle; }).join(', '),
      eur(g.annuel), String(g.mois), eur(g.mensuel)]);
  });
  ta.push(['TOTAL', '', eur(totA), '', eur(totM)]);
  var tabA = body.appendTable(ta);
  try {
    tabA.editAsText().setFontSize(8);
    tabA.getRow(0).editAsText().setBold(true);
    tabA.getRow(ta.length - 1).editAsText().setBold(true);
  } catch (e) { }
  body.appendParagraph('');

  st(body.appendParagraph('SIGNATURES'), 11, true);
  par('Fait à ' + (o.structure.adresse ? String(o.structure.adresse).split(',').slice(-1)[0].trim() : 'Villeneuve') +
    ', le ' + fmtDate_(o.date) + ', en autant d\'exemplaires que de parties.');
  var ts = [['Pour l\'association', 'Signature']];
  ts.push([o.structure.structure + '\n\n\n', '\n\n\n']);
  o.gens.forEach(function (g) { ts.push([g.nom + ', ' + g.statut + '\n\n\n', '\n\n\n']); });
  var tabS = body.appendTable(ts);
  try { tabS.editAsText().setFontSize(8); tabS.getRow(0).editAsText().setBold(true); } catch (e) { }
  par('Chaque partie ne signe que pour ce qui la concerne. La signature vaut acceptation de la présente convention et de ses annexes.');
}

/* ===== Cache.gs ===== */
/**
 * Cache.gs : les ecrans doivent s'afficher sans rien recalculer.
 *
 * Le probleme : chaque lecture (accueil, cours, appels, compta...) relit des classeurs Google
 * et refait ses calculs, a chaque connexion de chaque personne. C'est ce temps-la qu'on attend.
 *
 * La reponse ici : on ne calcule plus pendant que l'utilisateur attend.
 * 1. chaque lecture est mise en cache, reponse complete, cote serveur ;
 * 2. un declencheur toutes les 10 minutes recalcule tout en arriere-plan et remplit ce cache,
 *    de sorte qu'un ecran ouvert lit une valeur deja prete ;
 * 3. une ecriture n'invalide que ce qu'elle touche (table INVALIDE), le reste reste chaud ;
 * 4. le cache est partage par tout le monde : si un second utilisateur ouvre l ERP apres le premier,
 *    il profite de ce qui a deja ete calcule.
 */

// duree de vie par action, en secondes. Large : c'est le rechauffage et les invalidations
// qui gardent les donnees justes, pas l'expiration.
var CACHE_LECTURES = {
  accueil: 1800, cours: 1800, stages: 1800, urgencesTous: 1800, appelsTous: 1800,
  compta: 1800, achats: 1800, remunerations: 1800, factures: 1800,
  intervenants: 1800, facturesInterv: 1800,
  mails: 600, aRelancer: 1800, rappels: 1800, agenda: 900, notesSite: 1800,
  docs: 3600, groupes: 3600,
  appel: 1800              // une fiche d'appel par groupe, voir CACHE_PARAM
};

/**
 * Lectures qui dependent d'un parametre et qu'on cache quand meme, une entree par valeur.
 * La fiche d'appel d'un groupe est la meme pour tout le monde : la mettre en cache evite
 * de refaire la synchronisation de la fiche de presence a chaque ouverture.
 */
var CACHE_PARAM = {
  appel: function (p) {
    var src = norm_(p && p.source), cle = norm_(p && p.cle);
    return src && cle ? src + '|' + cle : '';
  }
};
var CACHE_INDEX_APPEL = 'rep_appel_index';

// ce qu'une ecriture rend faux. Meme logique que le cote navigateur, mais cote serveur.
var CACHE_ELEVES = ['cours', 'stages', 'accueil', 'urgencesTous', 'aRelancer', 'appelsTous', 'groupes'];
var CACHE_COMPTA = ['compta', 'achats', 'remunerations', 'factures'];
var INVALIDE = {
  suivi: CACHE_ELEVES.concat(['compta']),
  ajouterInscription: CACHE_ELEVES, modifierInscription: CACHE_ELEVES,
  rattacherCours: CACHE_ELEVES, supprimerInscription: CACHE_ELEVES, deplacer: CACHE_ELEVES,
  saveAttente: CACHE_ELEVES, statutAttente: CACHE_ELEVES, supprimerAttente: CACHE_ELEVES,
  reglagesGroupe: CACHE_ELEVES, completerSuivi: CACHE_ELEVES,
  saveAppel: ['appelsTous'],
  brouillon: ['mails', 'aRelancer'], envoyerMail: ['mails', 'aRelancer', 'accueil'],
  brouillonsRelance: ['mails', 'aRelancer', 'accueil'], noterRelance: ['aRelancer', 'accueil'], brouillonEnvoyer: ['mails', 'accueil'],
  brouillonModifier: ['mails'], brouillonSupprimer: ['mails', 'accueil'], mailsAction: ['mails', 'accueil'],
  ajouterDepense: CACHE_COMPTA, ajouterRecette: CACHE_COMPTA, modifierCaisse: CACHE_COMPTA,
  modifierLigneCompta: CACHE_COMPTA, supprimerLigneCompta: CACHE_COMPTA,
  saveAchat: CACHE_COMPTA, payerAchat: CACHE_COMPTA, supprimerAchat: CACHE_COMPTA,
  saveRemuneration: CACHE_COMPTA, syncModeleEco: CACHE_COMPTA,
  creerFacture: CACHE_COMPTA, statutFacture: CACHE_COMPTA,
  saveIntervenant: ['intervenants'], reinitNumInterv: ['intervenants'],
  creerFactureInterv: ['facturesInterv', 'intervenants'].concat(CACHE_COMPTA),
  envoyerFactureInterv: ['facturesInterv'], supprimerFactureInterv: ['facturesInterv'].concat(CACHE_COMPTA),
  saveRappel: ['rappels', 'accueil'], deleteRappel: ['rappels', 'accueil'],
  saveNoteSite: ['notesSite'], statutNoteSite: ['notesSite'], deleteNoteSite: ['notesSite'],
  creerEvenement: ['agenda', 'accueil'], conventionCollective: []
};

function cacheCle_(action, payload) {
  if (CACHE_PARAM[action]) {
    var k = CACHE_PARAM[action](payload);
    return k ? 'rep_' + action + '_' + k : '';
  }
  return 'rep_' + action;
}

/** Les cles de fiches d'appel en cache, pour savoir quoi vider (CacheService ne sait pas lister) */
function appelIndexLire_() {
  try { return JSON.parse(CacheService.getScriptCache().get(CACHE_INDEX_APPEL) || '[]'); } catch (e) { return []; }
}
function appelIndexAjouter_(cle) {
  try {
    var l = appelIndexLire_();
    if (l.indexOf(cle) < 0) { l.push(cle); CacheService.getScriptCache().put(CACHE_INDEX_APPEL, JSON.stringify(l.slice(-200)), 21600); }
  } catch (e) { }
}
function appelsVider_() {
  appelIndexLire_().forEach(function (c) { cacheRemoveBig_(c); });
  try { CacheService.getScriptCache().remove(CACHE_INDEX_APPEL); } catch (e) { }
}

/** Lecture passee par le cache. fn() n'est appelee que si le cache est vide. */
function cacheLecture_(action, fn, payload) {
  var ttl = CACHE_LECTURES[action];
  if (!ttl) return fn();
  var cle = cacheCle_(action, payload);
  if (!cle) return fn();
  var hit = cacheGetBig_(cle);
  if (hit) { try { return JSON.parse(hit); } catch (e) { } }
  var data = fn();
  try {
    cachePutBig_(cle, JSON.stringify(data), ttl);
    if (CACHE_PARAM[action]) appelIndexAjouter_(cle);
  } catch (e) { }
  return data;
}

/** Vide les reponses que cette ecriture rend fausses */
function cacheInvalider_(action) {
  var l = INVALIDE[action];
  // action d'ecriture non listee : on ne prend pas de risque, on vide tout
  if (l === undefined) l = Object.keys(CACHE_LECTURES);
  l.forEach(function (a) { cacheRemoveBig_(cacheCle_(a)); });
  // les fiches d'appel suivent le sort de appelsTous
  if (l.indexOf('appelsTous') >= 0) appelsVider_();
}

function cacheToutVider_() {
  Object.keys(CACHE_LECTURES).forEach(function (a) { if (!CACHE_PARAM[a]) cacheRemoveBig_(cacheCle_(a)); });
  appelsVider_();
}

/**
 * Rechauffage : recalcule les lectures et remplit le cache, hors du temps d'attente des gens.
 * Appele par un declencheur toutes les 10 min et par le bouton Actualiser.
 * Les lectures sont faites dans l'ordre du plus utile au moins utile : si le temps d'execution
 * manque, ce qui est en haut est deja pret.
 */
var CACHE_ORDRE = ['accueil', 'cours', 'stages', 'urgencesTous', 'appelsTous', 'agenda', 'rappels',
  'notesSite', 'compta', 'achats', 'remunerations', 'factures', 'intervenants', 'facturesInterv',
  'mails', 'aRelancer', 'groupes', 'docs'];

function rechaufferCaches() {
  var debut = Date.now();
  var R = routes_();
  var fait = [], rates = [];
  CACHE_ORDRE.forEach(function (a) {
    if (Date.now() - debut > 240000) return;          // on s'arrete avant la limite d'execution
    if (!R[a] || !CACHE_LECTURES[a]) return;
    try {
      var data = R[a]((typeof REQ_SERVEUR !== 'undefined' && REQ_SERVEUR[a]) || {});
      cachePutBig_(cacheCle_(a), JSON.stringify(data), CACHE_LECTURES[a]);
      fait.push(a);
      // appelsTous a deja tout calcule : on en profite pour remplir la fiche de chaque groupe,
      // pour que « Faire l'appel » n'ait plus rien a faire que lire
      if (a === 'appelsTous') {
        Object.keys(data || {}).forEach(function (k) {
          var x = k.split('|');
          var cle = cacheCle_('appel', { source: x[0], cle: x.slice(1).join('|') });
          if (!cle) return;
          cachePutBig_(cle, JSON.stringify(data[k]), CACHE_LECTURES.appel);
          appelIndexAjouter_(cle);
        });
      }
    } catch (e) { rates.push(a + ' (' + (e && e.message) + ')'); }
  });
  var msg = fait.length + ' prêtes en ' + Math.round((Date.now() - debut) / 1000) + ' s' +
    (rates.length ? ' · échecs : ' + rates.join(', ') : '');
  try { logAction_('rechauffage cache', msg); } catch (e) { }
  return msg;
}

/** Parametres par defaut des lectures qui en attendent (l'agenda veut une fenetre de jours) */
var REQ_SERVEUR = { agenda: { jours: 21 } };

/** Pose le declencheur de rechauffage (appele par setup) */
function poserRechauffage_() {
  try {
    ScriptApp.getProjectTriggers().forEach(function (t) {
      if (t.getHandlerFunction() === 'rechaufferCaches') ScriptApp.deleteTrigger(t);
    });
    ScriptApp.newTrigger('rechaufferCaches').timeBased().everyMinutes(10).create();
    return true;
  } catch (e) { return false; }
}

/* ===== Api.gs ===== */
/**
 * Api.gs : point d'entrée web + routeur unique protégé par le code PIN (onglet Config).
 */

function doGet() {
  var t = HtmlService.createTemplateFromFile('Index');
  // infos publiques (nom, logo, PIN configuré ?) glissées dans la page : un aller-retour serveur de moins au démarrage
  var info = {};
  try { info = apiInfo_(); } catch (e) { info = {}; }
  t.info = JSON.stringify(info).replace(/</g, '\\u003c');
  // le nom est aussi rendu cote serveur : il s'affiche avant le premier appel
  t.nom = String(info.nom || 'ERP');
  return t.evaluate()
    .setTitle("ERP · " + (info.nom || "Association"))
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, viewport-fit=cover')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.DEFAULT);
}

function include(name) {
  return HtmlService.createHtmlOutputFromFile(name).getContent();
}

// construit à l'appel (les fichiers .gs sont chargés dans l'ordre, pas de référence au chargement)
function routes_() {
  return {
  accueil: apiAccueil_,
  bloc: apiBloc_,
  conventionCollective: apiConventionCollective_,
  reinitNumInterv: apiReinitNumInterv_,
  cours: apiCours_,
  stages: apiStages_,
  recherche: apiRecherche_,
  urgencesTous: apiUrgencesTous_,
  eleve: apiEleve_,
  champs: apiChamps_,
  ajouterInscription: apiAjouterInscription_,
  modifierInscription: apiModifierInscription_,
  rattacherCours: apiRattacherCours_,
  supprimerInscription: apiSupprimerInscription_,
  suivi: apiSuivi_,
  urgencesGroupe: apiUrgencesGroupe_,
  syncModeleEco: apiSyncModeleEco_,
  completerSuivi: completerSuivi_,

  compta: apiCompta_,
  ajouterDepense: apiAjouterDepense_,
  ajouterRecette: apiAjouterRecette_,
  modifierCaisse: apiModifierCaisse_,
  modifierLigneCompta: apiModifierLigneCompta_,
  supprimerLigneCompta: apiSupprimerLigneCompta_,
  uploadJustificatif: apiUploadJustificatif_,
  achats: apiAchats_,
  saveAchat: apiSaveAchat_,
  payerAchat: apiPayerAchat_,
  supprimerAchat: apiSupprimerAchat_,
  intervenants: apiIntervenants_,
  saveIntervenant: apiSaveIntervenant_,
  propositionFacture: apiPropositionFacture_,
  creerFactureInterv: apiCreerFactureInterv_,
  facturesInterv: apiFacturesInterv_,
  envoyerFactureInterv: apiEnvoyerFactureInterv_,
  supprimerFactureInterv: apiSupprimerFactureInterv_,
  remunerations: apiRemunerations_,
  saveRemuneration: apiSaveRemuneration_,

  factures: apiFactures_,
  dossierDevis: apiDossierDevis_,
  creerFacture: apiCreerFacture_,
  statutFacture: apiStatutFacture_,

  apercuMail: apiApercuMail_,
  brouillon: apiBrouillon_,
  envoyerMail: apiEnvoyerMail_,
  brouillonVoir: apiBrouillonVoir_,
  brouillonEnvoyer: apiBrouillonEnvoyer_,
  brouillonModifier: apiBrouillonModifier_,
  brouillonSupprimer: apiBrouillonSupprimer_,
  aRelancer: apiARelancer_,
  brouillonsRelance: apiBrouillonsRelance_,
  noterRelance: apiNoterRelance_,
  mails: apiMails_,
  mailsAction: apiMailsAction_,
  mailVoir: apiMailVoir_,

  rappels: apiRappels_,
  saveRappel: apiSaveRappel_,
  deleteRappel: apiDeleteRappel_,
  notesSite: apiNotesSite_,
  saveNoteSite: apiSaveNoteSite_,
  statutNoteSite: apiStatutNoteSite_,
  deleteNoteSite: apiDeleteNoteSite_,
  agenda: apiAgenda_,
  creerEvenement: apiCreerEvenement_,
  medias: apiMedias_,
  rafraichirMedias: apiRafraichirMedias_,
  mediasDossiers: apiMediasDossiers_,
  uploadMedia: apiUploadMedia_,
  miniature: apiMiniature_,
  texteDoc: apiTexteDoc_,
  docs: apiDocs_,
  rapports: apiRapports_,
  info: apiInfo_,
  ping: function () { return true; },
  groupes: apiGroupes_,
  deplacer: apiDeplacer_,
  reglagesGroupe: apiReglagesGroupe_,
  attente: apiAttente_,
  saveAttente: apiSaveAttente_,
  statutAttente: apiStatutAttente_,
  supprimerAttente: apiSupprimerAttente_,
  appel: apiAppel_,
  appelsTous: apiAppelsTous_,
  saveAppel: apiSaveAppel_
  };
}

function apiInfo_() {
  var c = getConfig_();
  return {
    nom: c.NOM_STRUCTURE,
    pinConfigure: String(c.PIN || '').length >= 4,
    tel: c.TEL_STRUCTURE,
    logo: logo_(c.ID_LOGO),
    // le vocabulaire part avec le bootstrap : l'interface s'ecrit dans les mots
    // de la structure sans aller-retour serveur supplementaire
    voc: vocabulaire_()
  };
}

function hash_(s) {
  var d = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, String(s || ''), Utilities.Charset.UTF_8);
  return d.map(function (b) { return ('0' + (b & 0xff).toString(16)).slice(-2); }).join('');
}

/**
 * Anti force brute : 5 essais ratés -> blocage 15 min pour ce compte Google.
 */
var MAX_ESSAIS = 5, BLOCAGE_S = 900;

function userKey_() {
  try { return Session.getTemporaryActiveUserKey() || 'anon'; } catch (e) { return 'anon'; }
}

function checkPin_(pin) {
  var expected = String(getConfig_().PIN || '');
  if (expected.length < 4) return { ok: false, code: 'NOPIN', message: 'Aucun code PIN (4 caractères minimum) dans l\'onglet Config du classeur ERP. Ajoute-le puis recharge la page.' };
  var cache = CacheService.getScriptCache();
  var kUser = 'pin_fail_' + userKey_();
  var nUser = +(cache.get(kUser) || 0);
  if (nUser >= MAX_ESSAIS) return { ok: false, code: 'LOCK', message: 'Trop d\'essais. Réessaie dans 15 minutes.' };
  if (hash_(pin) === hash_(expected)) { if (nUser) cache.remove(kUser); return { ok: true }; }
  cache.put(kUser, String(nUser + 1), BLOCAGE_S);
  Utilities.sleep(1000);
  return { ok: false, code: 'PIN', message: 'Code incorrect (' + (MAX_ESSAIS - nUser - 1) + ' essai(s) restant(s)).' };
}

/**
 * Plusieurs lectures en un seul aller-retour.
 * Apps Script execute les appels d'un meme utilisateur a la file, un par un : seize appels
 * « en parallele » depuis le navigateur s'empilent en realite. Les regrouper ici fait seize
 * fois moins de latence reseau et d'initialisation, et les caches internes (eleves, parametres,
 * classeurs ouverts) sont partages par les lectures du meme bloc.
 * p = { actions: ['accueil','cours',...], params: { cours: {...} } }
 * Renvoie { cours: { v: ... }, mails: { err: '...' } } : une lecture qui echoue n'emporte pas les autres.
 */
function apiBloc_(p) {
  var R = routes_();
  var out = {};
  (p && p.actions || []).slice(0, 24).forEach(function (a) {
    if (out[a]) return;
    if (!LECTURES[a] || a === 'bloc') { out[a] = { err: 'Action non groupable : ' + a }; return; }
    var par = (p.params && p.params[a]) || {};
    try {
      out[a] = { v: estCachable_(a, par) ? cacheLecture_(a, function () { return R[a](par); }, par) : R[a](par) };
    } catch (e) { out[a] = { err: (e && e.message) || String(e) }; }
  });
  return out;
}

/**
 * Une reponse n'est mise en cache que si elle ne depend pas de parametres :
 * une fiche eleve ou un appel d'un groupe precis change avec le parametre, pas l'accueil.
 * appelsTous filtre par jour : seule la version complete est cachee.
 */
function estCachable_(action, payload) {
  if (!CACHE_LECTURES[action]) return false;
  // lecture cachee par parametre (la fiche d'appel d'un groupe) : une entree par valeur
  if (CACHE_PARAM[action]) return !!cacheCle_(action, payload);
  var ref = (typeof REQ_SERVEUR !== 'undefined' && REQ_SERVEUR[action]) || {};
  var cles = Object.keys(payload || {}).filter(function (k) { return k !== 'fresh'; });
  if (!cles.length) return true;
  return cles.every(function (k) { return String(payload[k]) === String(ref[k]); });
}

/** Actions en lecture seule (ne vident pas le cache) */
var LECTURES = { bloc: 1, attente: 1, mailVoir: 1, dossierDevis: 1, mediasDossiers: 1, appelsTous: 1, urgencesTous: 1, brouillonVoir: 1, accueil: 1, cours: 1, stages: 1, recherche: 1, eleve: 1, champs: 1, urgencesGroupe: 1, compta: 1, factures: 1,
  apercuMail: 1, aRelancer: 1, mails: 1, rappels: 1, notesSite: 1, agenda: 1, medias: 1, miniature: 1, texteDoc: 1, docs: 1,
  rapports: 1, ping: 1, groupes: 1, appel: 1, achats: 1, intervenants: 1, facturesInterv: 1, propositionFacture: 1, remunerations: 1 };

var TOUCHE_ELEVES = { saveAttente: 1, statutAttente: 1, supprimerAttente: 1, reglagesGroupe: 1, suivi: 1, ajouterInscription: 1, modifierInscription: 1, rattacherCours: 1, supprimerInscription: 1, deplacer: 1, brouillon: 1, envoyerMail: 1, brouillonsRelance: 1, noterRelance: 1 };
var TOUCHE_ACCUEIL = { mailsAction: 1, brouillon: 1, envoyerMail: 1, brouillonsRelance: 1, brouillonEnvoyer: 1, brouillonModifier: 1, brouillonSupprimer: 1, creerEvenement: 1 };

/** Bouton Actualiser : relit tout depuis les Sheets / Gmail / Agenda */
function viderCaches_() {
  invalideEleves_();
  cacheToutVider_();
  try { CacheService.getScriptCache().removeAll(['accueil_agenda', 'accueil_gmail', 'choixForm_cours', 'choixForm_stages']); } catch (e) { }
}

/** Routeur appelé par google.script.run.api(action, payload, pin) */
function api(action, payload, pin) {
  if (action === 'info') return JSON.stringify({ ok: true, data: apiInfo_() });
  var chk = checkPin_(pin);
  if (!chk.ok) return JSON.stringify({ ok: false, error: chk.code, message: chk.message });
  var fn = routes_()[action];
  if (!fn) return JSON.stringify({ ok: false, message: 'Action inconnue : ' + action });
  try {
    payload = payload || {};
    if (payload.fresh) viderCaches_();
    // une lecture repond depuis le cache serveur quand il est chaud : rien n'est recalculé
    var data = (LECTURES[action] && !payload.fresh && estCachable_(action, payload))
      ? cacheLecture_(action, function () { return fn(payload); }, payload)
      : fn(payload);
    // toute action qui écrit invalide le cache élèves (les pages se rechargent à jour)
    // seules les actions qui touchent les élèves vident le cache élèves (le reste reste rapide)
    if (!LECTURES[action]) cacheInvalider_(action);
    if (TOUCHE_ELEVES[action]) invalideEleves_();
    if (TOUCHE_ACCUEIL[action]) { try { CacheService.getScriptCache().removeAll(['accueil_agenda', 'accueil_gmail']); } catch (e) { } }
    return JSON.stringify({ ok: true, data: data });
  } catch (e) {
    logAction_('ERREUR ' + action, e && e.message);
    return JSON.stringify({ ok: false, message: (e && e.message) || String(e) });
  }
}

/** Logo en data URI (miniature Drive ~220 px, mise en cache 6 h) */
function logo_(id) {
  if (!id) return '';
  var cache = CacheService.getScriptCache();
  var hit = cache.get('logo_' + id);
  if (hit) return hit;
  try {
    var blob = DriveApp.getFileById(id).getThumbnail();
    if (!blob) return '';
    var data = 'data:' + (blob.getContentType() || 'image/png') + ';base64,' + Utilities.base64Encode(blob.getBytes());
    if (data.length < 95000) cache.put('logo_' + id, data, 21600);
    return data;
  } catch (e) { return ''; }
}
