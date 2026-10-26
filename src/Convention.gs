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
