#!/usr/bin/env node
/**
 * run.js : lanceur de tests minimal, sans dependance.
 *
 * `npm test` n'existe pas dans Apps Script et on ne veut pas d'un node_modules
 * de 200 Mo pour verifier des fonctions pures. Ce lanceur fait le strict
 * necessaire : des groupes, des assertions, un code de sortie non nul en cas
 * d'echec (ce qui suffit a une CI), et un rapport lisible.
 *
 *   node tests/run.js              tous les fichiers tests/*.test.js
 *   node tests/run.js mails        seulement ceux dont le nom contient "mails"
 */
const fs = require('fs');
const path = require('path');

const etat = { groupe: '', ok: 0, ko: 0, echecs: [], actuels: 0 };

function groupe(nom, fn) {
  etat.groupe = nom;
  const avant = etat.ok + etat.ko;
  try {
    fn();
  } catch (e) {
    etat.ko++;
    etat.echecs.push({ groupe: nom, test: '(le groupe a leve une exception)', message: e.stack || e.message });
  }
  const n = etat.ok + etat.ko - avant;
  console.log('  ' + nom + ' (' + n + ')');
}

function test(nom, fn) {
  try {
    fn();
    etat.ok++;
  } catch (e) {
    etat.ko++;
    etat.echecs.push({ groupe: etat.groupe, test: nom, message: e.message, pile: e.stack });
    console.log('    ECHEC  ' + nom);
  }
}

// ---- assertions ----
function fmt(v) {
  if (typeof v === 'string') return JSON.stringify(v);
  try { return JSON.stringify(v); } catch (e) { return String(v); }
}
function egal(obtenu, attendu, msg) {
  if (obtenu !== attendu) {
    throw new Error((msg ? msg + ' : ' : '') + 'attendu ' + fmt(attendu) + ', obtenu ' + fmt(obtenu));
  }
}
function profond(obtenu, attendu, msg) {
  const a = JSON.stringify(obtenu), b = JSON.stringify(attendu);
  if (a !== b) throw new Error((msg ? msg + ' : ' : '') + 'attendu ' + b + ', obtenu ' + a);
}
function vrai(v, msg) { if (!v) throw new Error(msg || 'attendu vrai, obtenu ' + fmt(v)); }
function faux(v, msg) { if (v) throw new Error(msg || 'attendu faux, obtenu ' + fmt(v)); }
function contient(txt, morceau, msg) {
  if (String(txt).indexOf(morceau) < 0) {
    throw new Error((msg ? msg + ' : ' : '') + fmt(morceau) + ' absent de ' + fmt(String(txt).slice(0, 220)));
  }
}
function absent(txt, morceau, msg) {
  if (String(txt).indexOf(morceau) >= 0) {
    throw new Error((msg ? msg + ' : ' : '') + fmt(morceau) + ' present alors qu\'il ne devait pas l\'etre');
  }
}
function leve(fn, motif, msg) {
  try {
    fn();
  } catch (e) {
    if (motif && !new RegExp(motif, 'i').test(e.message)) {
      throw new Error((msg ? msg + ' : ' : '') + 'exception levee mais message inattendu : ' + e.message);
    }
    return;
  }
  throw new Error(msg || 'aucune exception levee');
}

global.T = { groupe, test, egal, profond, vrai, faux, contient, absent, leve };

// ---- execution ----
const filtre = process.argv[2] || '';
const dir = __dirname;
const fichiers = fs.readdirSync(dir)
  .filter((f) => f.endsWith('.test.js'))
  .filter((f) => !filtre || f.indexOf(filtre) >= 0)
  .sort();

if (!fs.existsSync(path.join(dir, '..', 'bundle', 'Code.gs'))) {
  console.error('bundle/Code.gs absent. Lance d\'abord : node build.js');
  process.exit(1);
}

console.log('');
for (const f of fichiers) {
  console.log(f.replace('.test.js', ''));
  require(path.join(dir, f));
  console.log('');
}

if (etat.echecs.length) {
  console.log('--- echecs ---');
  for (const e of etat.echecs) console.log('\n' + e.groupe + ' > ' + e.test + '\n  ' + e.message);
  console.log('');
}
const total = etat.ok + etat.ko;
console.log(etat.ko === 0
  ? total + ' tests, tous au vert'
  : total + ' tests, ' + etat.ko + ' en echec');
process.exit(etat.ko === 0 ? 0 : 1);
