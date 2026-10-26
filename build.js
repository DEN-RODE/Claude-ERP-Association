#!/usr/bin/env node
/**
 * build.js : concatene les modules src/*.gs en un seul bundle/Code.gs.
 *
 * Apps Script n'a pas de systeme de modules : tous les fichiers .gs d'un projet
 * partagent un espace de noms global et sont charges dans l'ordre alphabetique
 * de l'editeur, ce qui rend l'ordre de declaration fragile. On developpe donc en
 * modules lisibles, et on colle un fichier unique dans l'editeur, dans un ordre
 * maitrise : les utilitaires d'abord, le routeur en dernier.
 *
 *   node build.js            -> bundle/Code.gs + les .html, version de package.json
 *   node build.js v2.1      -> idem, en forcant la version dans l'en-tete
 */
const fs = require('fs');
const path = require('path');

// version deterministe : l'argument, sinon celle de package.json. L'en-tete du
// bundle doit etre reproductible, sinon la CI ne peut pas verifier que le bundle
// versionne correspond bien aux sources.
const VERSION = process.argv[2] || require('./package.json').version;
const ORDRE = [
  'Util.gs',          // helpers purs : dates, normalisation, euros, regex
  'Config.gs',        // parametres, setup(), acces au classeur
  'Libelles.gs',      // couche vocabulaire (depend de Config)
  'Inscriptions.gs',  // lecture des reponses de formulaire, rattachement aux creneaux
  'Compta.gs',        // depenses, recettes, justificatifs
  'Factures.gs',      // devis et factures emises (Google Docs -> PDF)
  'Mails.gs',         // brouillons Gmail, relances, recap quotidien
  'Divers.gs',        // agenda, rappels, medias, notes
  'Presences.gs',     // appel et feuilles de presence
  'ModeleEco.gs',     // previsionnel et remuneration des intervenants
  'FacturesInterv.gs',// factures que les prestataires adressent a la structure
  'Convention.gs',    // convention annuelle de prestation
  'Cache.gs',         // cache serveur + rechauffage
  'Api.gs'            // point d'entree web et routeur (doit rester en dernier)
];

const SRC = path.join(__dirname, 'src');
const OUT = path.join(__dirname, 'bundle');

const manquants = ORDRE.filter((f) => !fs.existsSync(path.join(SRC, f)));
if (manquants.length) {
  console.error('Fichiers absents de src/ : ' + manquants.join(', '));
  process.exit(1);
}
const inconnus = fs.readdirSync(SRC)
  .filter((f) => f.endsWith('.gs') && ORDRE.indexOf(f) < 0);
if (inconnus.length) {
  console.error('Modules non declares dans ORDRE : ' + inconnus.join(', '));
  process.exit(1);
}

let out = '/** ERP Association - ' + VERSION + ' - genere par build.js - ne pas editer a la main */\n';
let lignes = 0;
for (const f of ORDRE) {
  const txt = fs.readFileSync(path.join(SRC, f), 'utf8');
  lignes += txt.split('\n').length;
  out += '\n/* ===== ' + f + ' ===== */\n' + txt;
}

fs.mkdirSync(OUT, { recursive: true });
fs.writeFileSync(path.join(OUT, 'Code.gs'), out);
for (const f of ['Index.html', 'Styles.html', 'App.html', 'appsscript.json']) {
  if (fs.existsSync(path.join(SRC, f))) fs.copyFileSync(path.join(SRC, f), path.join(OUT, f));
}

console.log(VERSION + ' : ' + ORDRE.length + ' modules, ' + lignes + ' lignes -> bundle/Code.gs');
