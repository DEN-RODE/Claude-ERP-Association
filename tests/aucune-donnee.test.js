/**
 * Garde-fou du depot : aucune donnee d'une structure reelle ne doit se trouver
 * dans le code publie.
 *
 * Ce fichier existe parce que ce code vient d'un ERP en production. La version
 * generique a ete obtenue en retirant les donnees de cette structure, et une
 * telle operation se refait a chaque contribution : un IBAN colle dans une
 * valeur par defaut "pour tester", un identifiant de Google Sheet laisse dans
 * un commentaire, et le depot fuit. Ce test echoue avant le commit.
 *
 * Il tourne sur les fichiers, pas sur les fonctions : c'est le seul test de la
 * suite qui lit le disque plutot que d'appeler du code.
 */
const fs = require('fs');
const path = require('path');
const { groupe, test, egal, profond, vrai } = global.T;

const RACINE = path.join(__dirname, '..');
const A_SCANNER = []
  .concat(fs.readdirSync(path.join(RACINE, 'src')).map((f) => 'src/' + f))
  .concat(fs.existsSync(path.join(RACINE, 'bundle')) ? fs.readdirSync(path.join(RACINE, 'bundle')).map((f) => 'bundle/' + f) : [])
  .filter((f) => /\.(gs|html|json)$/.test(f));

function lire(f) { return fs.readFileSync(path.join(RACINE, f), 'utf8'); }

/** Renvoie la liste des "fichier:ligne: extrait" qui matchent. */
function trouver(motif, filtre) {
  const out = [];
  for (const f of A_SCANNER) {
    lire(f).split('\n').forEach((ligne, i) => {
      const m = ligne.match(motif);
      if (m && (!filtre || !filtre(m[0], ligne))) out.push(f + ':' + (i + 1) + ': ' + m[0]);
    });
  }
  return out;
}

groupe('aucune coordonnee bancaire', () => {
  test('pas d\'IBAN', () => {
    profond(trouver(/\b[A-Z]{2}\d{2}[\s]?[0-9A-Z]{4}([\s]?[0-9A-Z]{4}){2,}/), []);
  });
  test('pas de BIC dans une chaine litterale', () => {
    // un BIC fait 8 ou 11 caracteres, dont un code pays ISO en positions 5 et 6.
    // On ne cherche que dans les chaines litterales : hors de la, une suite de
    // majuscules est un nom de constante, pas une coordonnee bancaire.
    profond(trouver(/'[A-Z]{4}(FR|BE|CH|LU|DE|ES|IT|GB|NL|PT)[0-9A-Z]{2}([0-9A-Z]{3})?'/), []);
  });
});

groupe('aucune coordonnee de contact', () => {
  test('pas d\'adresse mail, hors example.org reserve a la documentation', () => {
    profond(trouver(/[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}/, (v) => /example\.(org|com)$/.test(v)), []);
  });
  test('pas de numero de telephone francais', () => {
    // 0612345678 est cite dans un commentaire comme exemple de normalisation
    profond(trouver(/\b0[1-9]([\s.-]?\d{2}){4}\b/, (v) => v.replace(/\D/g, '') === '0612345678'), []);
  });
  test('pas de SIRET ni de numero RNA', () => {
    profond(trouver(/\b\d{3}[\s]?\d{3}[\s]?\d{3}[\s]?\d{5}\b/), []);
    profond(trouver(/\bW\d{9}\b/), []);
  });
});

groupe('aucun identifiant de fichier Google', () => {
  test('pas d\'identifiant Drive, Sheet ou Form en dur', () => {
    // un identifiant Google fait 25 caracteres ou plus, melangeant lettres,
    // chiffres, tirets et soulignes, et commence quasi toujours par 1
    const trouves = trouver(/['"][-_A-Za-z0-9]{25,60}['"]/, (v) => {
      const s = v.slice(1, -1);
      if (!/[0-9]/.test(s)) return true;                 // constante en texte
      if (/^[A-Z_]+$/.test(s)) return true;              // cle de config
      if (/^[a-z]+([A-Z][a-z]+)*$/.test(s)) return true; // identifiant camelCase
      return false;
    });
    profond(trouves, []);
  });
  test('pas d\'URL de classeur ou de formulaire', () => {
    profond(trouver(/docs\.google\.com\/(spreadsheets|forms|document)\/d\/[-_A-Za-z0-9]+/), []);
  });
});

groupe('aucune valeur d\'identite par defaut dans Config.gs', () => {
  const cfg = lire('src/Config.gs');
  const DOIVENT_ETRE_VIDES = [
    'NOM_STRUCTURE', 'EMAIL_STRUCTURE', 'TEL_STRUCTURE', 'ADRESSE_STRUCTURE', 'SITE_WEB',
    'NOM_ASSO', 'SIRET_ASSO', 'PIN', 'BASELINE', 'MOT_ACCENTUE',
    'IBAN', 'IBAN_TITULAIRE', 'BIC', 'SIGNATAIRE', 'DIGEST_EMAILS', 'INTERVENANTS',
    'CONTACT_PAIEMENTS', 'REDUC_MOTIFS', 'DEBUT_SAISON_COURS', 'VACANCES_COURS', 'COULEURS_AGENDA'
  ];
  DOIVENT_ETRE_VIDES.forEach((cle) => {
    test(cle + ' est vide par defaut', () => {
      const m = cfg.match(new RegExp("\\['" + cle + "',\\s*('')|\\['" + cle + "',\\s*'([^']*)'"));
      vrai(m, cle + ' doit figurer dans DEFAULT_CONFIG');
      egal(m[2] === undefined ? '' : m[2], '', cle + ' porte une valeur en dur');
    });
  });
  test('toutes les cles ID_* sont vides par defaut', () => {
    const ids = cfg.match(/\['ID_[A-Z_]+',\s*'[^']*'/g) || [];
    const remplies = ids.filter((l) => !/,\s*''/.test(l));
    profond(remplies, [], 'identifiants Google en dur');
  });
});

groupe('le vocabulaire metier reste neutre', () => {
  // Ce depot doit convenir a une ecole de musique comme a un club sportif :
  // aucune discipline ne doit etre inscrite dans le code.
  const DISCIPLINES = ['trapeze', 'trapèze', 'jonglage', 'acrobatie', 'solfege', 'solfège',
    'judo', 'tatami', 'piano', 'violon', 'chapiteau', 'agres', 'agrès'];
  DISCIPLINES.forEach((mot) => {
    test('aucune occurrence de "' + mot + '"', () => {
      profond(trouver(new RegExp(mot, 'i')), []);
    });
  });
});

groupe('couverture du scan', () => {
  test('les 17 fichiers source sont scannes', () => {
    vrai(A_SCANNER.filter((f) => f.indexOf('src/') === 0).length >= 17,
      'fichiers scannes : ' + A_SCANNER.length);
  });
  test('le bundle genere est scanne aussi', () => {
    vrai(A_SCANNER.indexOf('bundle/Code.gs') >= 0, 'bundle/Code.gs doit etre dans le scan');
  });
});
