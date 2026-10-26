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
