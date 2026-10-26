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
