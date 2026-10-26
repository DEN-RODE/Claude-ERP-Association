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
