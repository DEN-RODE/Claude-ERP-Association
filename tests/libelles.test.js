/**
 * Couche vocabulaire : c'est elle qui rend l'ERP reutilisable. Le modele de
 * donnees garde des noms stables en interne ; les mots affiches viennent de la
 * config. Une ecole de musique, un club sportif et une MJC font tourner le meme
 * code avec trois vocabulaires differents.
 */
const { charger } = require('./harness');
const { erpDeTest } = require('./fixtures');
const { groupe, test, egal, contient, profond } = global.T;

groupe('valeurs par defaut, quand rien n\'est configure', () => {
  const { erp } = charger(erpDeTest());
  test('un inscrit est un adherent', () => egal(erp.L_('membre'), 'adherent'));
  test('le pluriel suit', () => egal(erp.L_('membres'), 'adherents'));
  test('une cle inconnue se renvoie elle-meme au lieu de casser l\'affichage', () => {
    egal(erp.L_('cleQuiNExistePas'), 'cleQuiNExistePas');
  });
});

groupe('vocabulaire d\'une ecole de musique', () => {
  const { erp } = charger(erpDeTest({
    config: { VOC_MEMBRE: 'eleve', VOC_MEMBRES: 'eleves', VOC_COURS: 'cours', VOC_STAGE: 'stage', VOC_LIEU: 'au conservatoire' }
  }));
  test('les mots configures remplacent les defauts', () => {
    egal(erp.L_('membre'), 'eleve');
    egal(erp.L_('lieu'), 'au conservatoire');
  });
  test('ils arrivent dans les mails', () => {
    contient(erp.infosPaiement_('cheque'), 'au conservatoire');
  });
});

groupe('vocabulaire d\'un club sportif', () => {
  const { erp } = charger(erpDeTest({
    config: { VOC_MEMBRE: 'licencie', VOC_MEMBRES: 'licencies', VOC_COURS: 'entrainement', VOC_COURS_PL: 'entrainements', VOC_STAGE: 'stage', VOC_LIEU: 'au gymnase' }
  }));
  test('la reference a indiquer sur un virement parle de licencie, avec le bon article', () => {
    contient(erp.infosPaiement_('virement'), 'le nom du licencie en référence');
  });
  test('le vocabulaire complet est expose au navigateur par apiInfo_', () => {
    const v = erp.apiInfo_().voc;
    egal(v.membre, 'licencie');
    egal(v.coursPl, 'entrainements');
    egal(v.lieu, 'au gymnase');
  });
});

groupe('Lmaj_ et Lle_ : les formes grammaticales', () => {
  const { erp } = charger(erpDeTest({ config: { VOC_COURS: 'atelier', VOC_MEMBRE: 'adherent' } }));
  test('Lmaj_ met une majuscule pour un debut de phrase', () => egal(erp.Lmaj_('cours'), 'Atelier'));
  test('Lle_ elide devant une voyelle', () => egal(erp.Lle_('cours'), "l'atelier"));
  test('Lle_ n\'elide pas devant une consonne', () => {
    const { erp: e } = charger(erpDeTest({ config: { VOC_COURS: 'cours' } }));
    egal(e.Lle_('cours'), 'le cours');
  });
  test('Lun_ construit l\'article indefini', () => egal(erp.Lun_('membre'), 'un adherent'));
  test('Ldu_ elide le complement du nom devant une voyelle', () => egal(erp.Ldu_('membre'), "de l'adherent"));
  test('et contracte en "du" devant une consonne', () => {
    const { erp: e } = charger(erpDeTest({ config: { VOC_MEMBRE: 'licencie' } }));
    egal(e.Ldu_('membre'), 'du licencie');
  });
});

groupe('Ltxt_ : gabarits a trous', () => {
  const { erp } = charger(erpDeTest({ config: { VOC_MEMBRE: 'licencie', VOC_MEMBRES: 'licencies', VOC_LIEU: 'au gymnase' } }));
  test('remplace les marqueurs', () => {
    egal(erp.Ltxt_('Chaque {{membre}} vient {{lieu}}'), 'Chaque licencie vient au gymnase');
  });
  test('un marqueur capitalise ressort capitalise', () => {
    egal(erp.Ltxt_('{{Membres}} inscrits'), 'Licencies inscrits');
  });
  test('un marqueur inconnu reste visible au lieu de disparaitre silencieusement', () => {
    egal(erp.Ltxt_('{{inconnu}} ici'), '{{inconnu}} ici');
  });
  test('un texte vide ou null ne leve pas', () => {
    egal(erp.Ltxt_(''), '');
    egal(erp.Ltxt_(null), '');
  });
});

groupe('une cle de vocabulaire vide retombe sur le defaut', () => {
  const { erp } = charger(erpDeTest({ config: { VOC_MEMBRE: '', VOC_MEMBRES: '   ' } }));
  test('chaine vide : defaut', () => egal(erp.L_('membre'), 'adherent'));
  test('espaces seuls : defaut', () => egal(erp.L_('membres'), 'adherents'));
});
