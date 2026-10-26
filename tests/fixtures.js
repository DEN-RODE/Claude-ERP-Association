/**
 * fixtures.js : donnees de test. Inventees de bout en bout, sans aucun lien avec
 * une structure reelle. Les noms viennent d'une liste de prenoms neutres.
 */
const { config } = require('./harness');

const CONFIG_MIN = {
  NOM_STRUCTURE: 'Association Test',
  NOM_ASSO: 'Association Test',
  EMAIL_STRUCTURE: 'contact@example.org',
  TEL_STRUCTURE: '01 02 03 04 05',
  ADRESSE_STRUCTURE: '1 rue des Tests, 00000 Villeneuve',
  SITE_WEB: 'example.org',
  SIRET_ASSO: '000 000 000 00000',
  IBAN_TITULAIRE: 'ASSOCIATION TEST',
  IBAN: 'FR00 0000 0000 0000 0000 0000 000',
  BIC: 'TESTFRPP',
  ID_COURS_REPONSES: 'REP_COURS',
  ID_STAGES_REPONSES: 'REP_STAGES',
  ADHESION: '10',
  PROCHAIN_NUM_FACTURE: '2026-001',
  PROCHAIN_NUM_DEVIS: 'DEV-2026-001',
  ECHEANCE_FACTURE_JOURS: '30',
  VALIDITE_DEVIS_JOURS: '30',
  RELANCE_APRES_JOURS: '10',
  RELANCE_INTERVALLE_JOURS: '7',
  SEUIL_MIN_COURS: '4',
  SEUIL_MIN_STAGE: '6',
  DEBUT_SAISON_COURS: '2026-09-07',
  VACANCES_COURS: '2026-10-17>2026-11-01; 2026-12-19>2027-01-03'
};

const EN_TETE_COURS = ['Cle', 'Libelle', 'Jour', 'Horaire', 'Public', 'Tarif', 'Max', 'Code'];
const COURS = [
  EN_TETE_COURS,
  ['Lundi 17h00', 'Groupe 6-9 ans', 'Lundi', '17h00-18h30', 'Enfants 6-9 ans', 300, 12, 'C1'],
  ['Mardi 17h00', 'Groupe 9-13 ans', 'Mardi', '17h00-19h00', 'Enfants 9-13 ans', 350, 12, 'C2'],
  ['Mercredi 11h00', 'Groupe 4-6 ans', 'Mercredi', '11h00-12h00', 'Enfants 4-6 ans', 250, 12, 'C3'],
  ['Jeudi 19h15', 'Groupe adultes', 'Jeudi', '19h15-21h15', 'Adultes 18+', 350, 8, 'C4']
];

const EN_TETE_STAGES = ['Cle', 'Groupe', 'Nom', 'Dates', 'Public', 'Tarif', 'Max', 'Debut'];
const STAGES = [
  EN_TETE_STAGES,
  ['26 au 30 octobre 2026', 'matin', 'Session 1', 'Du 26 au 30 octobre 2026', 'Matin 10h-12h, 4-6 ans', 200, 12, '2026-10-26'],
  ['26 au 30 octobre 2026', 'apres-midi', 'Session 1', 'Du 26 au 30 octobre 2026', 'Apres-midi 13h30-16h30, 7 ans et +', 250, 12, '2026-10-26'],
  ['15 au 19 fevrier 2027', 'matin', 'Session 2', 'Du 15 au 19 fevrier 2027', 'Matin 10h-12h, 4-6 ans', 200, 12, '2027-02-15']
];

const SUIVI_VIDE = [[
  'ID', 'Source', 'Nom', 'Prenom', 'Paye', 'MontantRecu', 'ModeRecu', 'Statut', 'Notes',
  'DerniereRelance', 'Confirmation', 'MajPar', 'MajLe'
]];

const TABLES_VIDES = {
  Factures: [['ID', 'Numero', 'Date', 'Client', 'Objet', 'Montant', 'Echeance', 'Statut', 'PayeeLe', 'LienDoc', 'LienPdf', 'Type', 'Details']],
  Achats: [['ID', 'Date', 'Beneficiaire', 'Type', 'Objet', 'Montant', 'Echeance', 'Statut', 'PayeLe', 'Mode', 'Lien', 'Pieces', 'Notes', 'MajPar', 'MajLe']],
  Intervenants: [['ID', 'Nom', 'Statut', 'Adresse', 'CP', 'Ville', 'SIRET', 'Email', 'Tel', 'IBAN', 'BIC', 'MentionTVA', 'Assurance', 'ProchainNum', 'Actif', 'MajPar', 'MajLe']],
  Attente: [['ID', 'Date', 'Source', 'CleGroupe', 'Nom', 'Prenom', 'Naissance', 'Responsable', 'Tel', 'Email', 'Notes', 'Statut', 'MajPar', 'MajLe']],
  Log: [['Date', 'Qui', 'Action', 'Detail']]
};

/** ERP complet pret a l'emploi, surchargeable. */
function erpDeTest(surcharge = {}) {
  const cfg = Object.assign({}, CONFIG_MIN, surcharge.config || {});
  return {
    onglets: Object.assign({
      Config: config(cfg),
      Cours: COURS,
      Stages: STAGES,
      Suivi: SUIVI_VIDE
    }, TABLES_VIDES, surcharge.onglets || {}),
    autres: Object.assign({
      REP_COURS: { 'Reponses': EN_TETE_REPONSES_COURS },
      REP_STAGES: { 'Reponses': EN_TETE_REPONSES_STAGES }
    }, surcharge.autres || {})
  };
}

const EN_TETE_REPONSES_COURS = [[
  'Horodateur', 'Nom de l\'enfant', 'Prenom de l\'enfant', 'Date de naissance',
  'Creneau souhaite', 'Nom du responsable legal', 'Telephone', 'Adresse e-mail',
  'Personne a contacter en cas d\'urgence', 'Allergies ou informations utiles',
  'Reductions eventuelles', 'Commentaire'
]];

const EN_TETE_REPONSES_STAGES = [[
  'Horodateur', 'Nom de l\'enfant', 'Prenom de l\'enfant', 'Date de naissance',
  'Stage souhaite', 'Nom du responsable legal', 'Telephone', 'Adresse e-mail',
  'Personne a contacter en cas d\'urgence', 'Allergies ou informations utiles',
  'Reductions eventuelles', 'Commentaire'
]];

/** Un inscrit, dans la forme que produit toEleve_ et que consomment les mails. */
function inscrit(surcharge = {}) {
  return Object.assign({
    id: '1700000000|martin|lea',
    source: 'cours',
    nom: 'Martin',
    prenom: 'Lea',
    naissance: '12/04/2017',
    creneaux: 'Groupe 6-9 ans - Lundi 17h00-18h30 [C1]',
    stage: '',
    responsable: 'Martin Claude',
    tel: '01 02 03 04 05',
    email: 'famille.martin@example.org',
    urgence: '',
    sante: '',
    reductions: '',
    notes: '',
    commentaire: '',
    paye: 'Non',
    montantRecu: 0,
    derniereRelance: ''
  }, surcharge);
}

module.exports = {
  CONFIG_MIN, COURS, STAGES, SUIVI_VIDE, TABLES_VIDES,
  EN_TETE_REPONSES_COURS, EN_TETE_REPONSES_STAGES,
  erpDeTest, inscrit
};
