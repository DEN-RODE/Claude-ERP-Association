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
