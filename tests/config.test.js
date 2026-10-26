/**
 * Installation et configuration : ce que fait setup() sur un classeur vierge,
 * et ce qu'il NE doit pas faire sur un classeur deja en service. C'est le
 * premier geste de quiconque reprend ce code, et le plus facile a casser : une
 * reinstallation qui ecrase les parametres d'une structure en pleine saison
 * serait une perte de donnees.
 */
const { charger, config } = require('./harness');
const { erpDeTest, CONFIG_MIN } = require('./fixtures');
const { groupe, test, egal, vrai, faux, contient, profond, leve } = global.T;

groupe('setup() sur un classeur vierge', () => {
  const { erp, classeur } = charger({ onglets: {} });
  erp.setup();
  const noms = classeur.getSheets().map((s) => s.getName()).sort();
  test('cree les 13 onglets du modele', () => egal(noms.length, 13));
  test('dont les onglets structurants', () => {
    ['Config', 'Cours', 'Stages', 'Suivi', 'Factures', 'Intervenants', 'Presences', 'Log'].forEach((n) => {
      vrai(noms.indexOf(n) >= 0, 'onglet manquant : ' + n);
    });
  });
  test('pose une ligne d\'en-tete figee sur chaque onglet', () => {
    const suivi = classeur.getSheetByName('Suivi');
    egal(suivi.getRange(1, 1).getValue(), 'ID');
    egal(suivi.getRange(1, 3).getValue(), 'Nom');
  });
  test('pre-remplit la config avec toutes les cles, valeurs d\'identite vides', () => {
    const cfg = classeur.getSheetByName('Config');
    const lignes = cfg.getRange(2, 1, cfg.getLastRow() - 1, 2).getValues();
    const map = {};
    lignes.forEach(([k, v]) => { map[k] = v; });
    egal(map.NOM_STRUCTURE, '', 'aucun nom ne doit etre code en dur');
    egal(map.IBAN, '', 'aucun IBAN ne doit etre code en dur');
    egal(map.SIRET_ASSO, '', 'aucun SIRET ne doit etre code en dur');
    egal(map.EMAIL_STRUCTURE, '', 'aucune adresse mail ne doit etre codee en dur');
    egal(map.ID_COURS_REPONSES, '', 'aucun identifiant Google ne doit etre code en dur');
  });
  test('pose quand meme des valeurs de fonctionnement sensees', () => {
    const cfg = classeur.getSheetByName('Config');
    const map = {};
    cfg.getRange(2, 1, cfg.getLastRow() - 1, 2).getValues().forEach(([k, v]) => { map[k] = v; });
    egal(map.RELANCE_APRES_JOURS, '10');
    egal(map.VOC_MEMBRE, 'adherent');
    egal(map.MENTION_TVA.indexOf('293 B') >= 0, true);
  });
  test('ecrit les creneaux d\'exemple pour qu\'un ecran ne soit pas vide au premier lancement', () => {
    vrai(classeur.getSheetByName('Cours').getLastRow() > 1);
    vrai(classeur.getSheetByName('Stages').getLastRow() > 1);
  });
  test('installe les declencheurs', () => {
    const { erp: e, etat } = charger({ onglets: {} });
    e.setup();
    vrai(etat.triggers.length > 0, 'au moins la tache quotidienne');
    vrai(etat.triggers.some((t) => t.getHandlerFunction() === 'dailyJob'));
  });
});

groupe('setup() relance sur un classeur en service : rien n\'est ecrase', () => {
  const { erp, classeur } = charger(erpDeTest());
  const avantCours = classeur.getSheetByName('Cours').getLastRow();
  erp.setup();
  test('les valeurs de config saisies par la structure survivent', () => {
    const cfg = classeur.getSheetByName('Config');
    const map = {};
    cfg.getRange(2, 1, cfg.getLastRow() - 1, 2).getValues().forEach(([k, v]) => { map[k] = v; });
    egal(map.NOM_STRUCTURE, 'Association Test');
    egal(map.IBAN, CONFIG_MIN.IBAN);
  });
  test('les creneaux saisis ne sont pas remplaces par les exemples', () => {
    egal(classeur.getSheetByName('Cours').getLastRow(), avantCours);
    egal(classeur.getSheetByName('Cours').getRange(2, 2).getValue(), 'Groupe 6-9 ans');
  });
  test('les cles de config ajoutees par une nouvelle version sont ajoutees', () => {
    const cfg = classeur.getSheetByName('Config');
    const cles = cfg.getRange(2, 1, cfg.getLastRow() - 1, 1).getValues().map(([k]) => k);
    vrai(cles.indexOf('VOC_MEMBRE') >= 0, 'une cle absente du classeur doit etre ajoutee');
    vrai(cles.indexOf('REDUC_MOTIFS') >= 0);
  });
  test('une seule occurrence par cle, meme apres plusieurs setup()', () => {
    erp.setup();
    const cfg = classeur.getSheetByName('Config');
    const cles = cfg.getRange(2, 1, cfg.getLastRow() - 1, 1).getValues().map(([k]) => String(k));
    const doublons = cles.filter((k, i) => k && cles.indexOf(k) !== i);
    profond(doublons, [], 'cles en double');
  });
});

groupe('getConfig_ : lecture des parametres', () => {
  const { erp } = charger(erpDeTest());
  test('lit la valeur du classeur', () => egal(erp.getConfig_().NOM_STRUCTURE, 'Association Test'));
  test('une cle absente du classeur retombe sur le defaut du code', () => {
    egal(erp.getConfig_().ONGLET_SYNTHESE, 'Synthese');
  });
  test('les valeurs sont des chaines, meme quand le tableur a devine un nombre', () => {
    egal(typeof erp.getConfig_().ADHESION, 'string');
  });
  test('setConfig_ ecrit et la lecture suivante voit la nouvelle valeur', () => {
    erp.setConfig_('NOM_STRUCTURE', 'Autre Association');
    egal(erp.getConfig_(true).NOM_STRUCTURE, 'Autre Association');
  });
  test('setConfig_ sur une cle inconnue l\'ajoute au lieu d\'echouer', () => {
    erp.setConfig_('CLE_NOUVELLE', 'valeur');
    egal(erp.getConfig_(true).CLE_NOUVELLE, 'valeur');
  });
});

groupe('verifierConfig : dit ce qui manque avant la mise en service', () => {
  test('sur un classeur vierge, nomme les manques bloquants', () => {
    const { erp } = charger({ onglets: { Config: config({}) } });
    const msg = erp.verifierConfig();
    contient(msg, 'A remplir avant utilisation');
    contient(msg, 'NOM_STRUCTURE');
    contient(msg, 'ID_COURS_REPONSES');
  });
  test('sur un classeur renseigne, confirme que le minimum est la', () => {
    const { erp } = charger(erpDeTest());
    contient(erp.verifierConfig(), 'Configuration minimale complete');
  });
  test('distingue le bloquant de l\'optionnel', () => {
    const { erp } = charger(erpDeTest({ config: { ID_MODELE_ECO: '', PIN: '' } }));
    const msg = erp.verifierConfig();
    contient(msg, 'Optionnel');
    contient(msg, 'ID_MODELE_ECO');
  });
});

groupe('readTable_ et upsertRow_ : acces generique aux onglets', () => {
  const { erp, classeur } = charger(erpDeTest({
    onglets: { Rappels: [['ID', 'Date', 'Heure', 'Type', 'Titre', 'Detail', 'Fait']] }
  }));
  test('readTable_ renvoie des objets clefs par en-tete', () => {
    const t = erp.readTable_('Cours');
    egal(t[0].Libelle, 'Groupe 6-9 ans');
    egal(t[0].Tarif, 300);
  });
  test('chaque ligne porte son numero de ligne pour l\'ecriture', () => {
    egal(erp.readTable_('Cours')[0]._row, 2);
  });
  test('upsertRow_ insere une ligne nouvelle', () => {
    erp.upsertRow_('Rappels', { ID: 'R1', Titre: 'Appeler la mairie', Fait: 'non' }, 'ID');
    const t = erp.readTable_('Rappels');
    egal(t.length, 1);
    egal(t[0].Titre, 'Appeler la mairie');
  });
  test('upsertRow_ met a jour au lieu de dupliquer quand l\'ID existe', () => {
    erp.upsertRow_('Rappels', { ID: 'R1', Titre: 'Rappeler la mairie' }, 'ID');
    const t = erp.readTable_('Rappels');
    egal(t.length, 1, 'pas de doublon');
    egal(t[0].Titre, 'Rappeler la mairie');
  });
  test('deleteRowById_ supprime la bonne ligne', () => {
    erp.upsertRow_('Rappels', { ID: 'R2', Titre: 'Autre' }, 'ID');
    erp.deleteRowById_('Rappels', 'R1', 'ID');
    const t = erp.readTable_('Rappels');
    egal(t.length, 1);
    egal(t[0].ID, 'R2');
  });
  test('un onglet inexistant leve une erreur explicite', () => {
    leve(() => erp.erpSheet_('OngletInvente'), 'Onglet');
  });
});
