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
