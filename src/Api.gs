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
