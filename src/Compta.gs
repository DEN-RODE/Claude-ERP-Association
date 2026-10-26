/**
 * Compta.gs : travaille sur le petit tableau de caisse existant de Modele_economique
 * (Affichage, Matos : boule, Portique, ... puis Cash / Chèque / Virement / Total).
 *
 * Règles :
 *  - le tableau est retrouvé par ses libellés (cellules "Cash" et "Total"), pas par adresse fixe
 *  - une DÉPENSE = une nouvelle ligne insérée juste au-dessus de "Cash" (montant négatif)
 *  - une RECETTE = ajoutée au total du mode (Cash / Chèque / Virement), la cellule garde
 *    l'historique sous forme de formule (=1000+300+250) + une note datée
 *  - 2 colonnes ajoutées à droite du montant : Date et Justificatif (lien Drive)
 *  - on insère des CELLULES (pas des lignes entières) pour ne pas décaler les autres tableaux
 */

var MODES = ['cash', 'cheque', 'virement'];

function comptaLocate_() {
  var c = getConfig_();
  var ss = SpreadsheetApp.openById(c.ID_MODELE_ECO);
  var sheets = ss.getSheets();
  // l'onglet "Compte aussi" (Config ONGLET_COMPTE) est cherché en premier
  var pref = norm_(c.ONGLET_COMPTE || '');
  if (pref) sheets = sheets.slice().sort(function (a, b) { return (norm_(b.getName()) === pref) - (norm_(a.getName()) === pref); });
  for (var s = 0; s < sheets.length; s++) {
    var sh = sheets[s];
    var v = sh.getDataRange().getValues();
    for (var r = 0; r < v.length; r++) {
      for (var col = 0; col < v[r].length; col++) {
        if (norm_(v[r][col]) !== 'cash') continue;
        // on a "Cash" : cherche Chèque, Virement, Total dans la même colonne en dessous
        var rows = { cash: r };
        var totalRow = -1;
        for (var k = r + 1; k < Math.min(r + 12, v.length); k++) {
          var lab = norm_(v[k][col]);
          if (lab === 'cheque' || lab === 'cheques') rows.cheque = k;
          else if (lab === 'virement' || lab === 'virements') rows.virement = k;
          else if (lab === 'total') { totalRow = k; break; }
        }
        if (totalRow < 0) continue;
        // début du bloc : on remonte jusqu'à 3 lignes vides d'affilée
        var scanStart = r, empties = 0;
        for (var u = r - 1; u >= 0 && empties < 3; u--) {
          var filled = String(v[u][col]).trim() !== '' || String(v[u][col + 1]).trim() !== '';
          if (filled) { scanStart = u; empties = 0; } else empties++;
        }
        var L = { ss: ss, sheet: sh, col: col + 1, cashRow: r + 1, totalRow: totalRow + 1, bloque: '',
          modeRows: { cash: rows.cash + 1, cheque: rows.cheque !== undefined ? rows.cheque + 1 : null, virement: rows.virement !== undefined ? rows.virement + 1 : null } };
        var tf = sh.getRange(totalRow + 1, col + 2).getFormula();
        var fm = tf && tf.match(/^=\s*SUM\(\s*\$?([A-Z]+)\$?(\d+)\s*:\s*\$?([A-Z]+)\$?(\d+)\s*\)\s*$/i);
        var amountCol = columnLetter_(col + 2);
        if (fm && fm[1].toUpperCase() === amountCol && fm[3].toUpperCase() === amountCol && Math.min(+fm[2], +fm[4]) < r + 1) {
          L.totalMode = 'sum'; L.start = Math.min(+fm[2], +fm[4]);
        } else if (!tf) {
          L.totalMode = 'value'; L.start = scanStart + 1;
          var somme = 0;
          for (var q = scanStart; q < totalRow; q++) somme += num_(v[q][col + 1]);
          if (Math.abs(somme - num_(v[totalRow][col + 1])) > 0.005) L.bloque = 'Le Total de la caisse est une valeur tapée à la main qui ne correspond pas à la somme des lignes. Mets une formule =SUM(...) sur la cellule Total, puis réessaie.';
        } else {
          L.totalMode = 'autre'; L.start = scanStart + 1;
          L.bloque = 'La cellule Total de la caisse contient une formule que l\'ERP ne sait pas étendre sans risque (' + tf + '). Remplace-la par =SUM(' + amountCol + (scanStart + 1) + ':' + amountCol + totalRow + ') puis réessaie.';
        }
        // colonnes Date / Justificatif (à droite du montant) : vides, dates ou liens uniquement
        if (!L.bloque) {
          for (var q2 = L.start - 1; q2 <= totalRow; q2++) {
            for (var dc = 2; dc <= 3; dc++) {
              var x = v[q2] ? v[q2][col + dc] : '';
              if (x === '' || x === null || x === undefined || isDate_(x)) continue;
              var xs = norm_(x);
              if (/^https?:\/\//.test(xs) || xs === 'date' || xs === 'justificatif') continue;
              L.bloque = 'La caisse a quelque chose à droite du montant (' + columnLetter_(col + 1 + dc) + (q2 + 1) + ' : "' + String(x).slice(0, 30) + '"). L\'ERP a besoin des 2 colonnes à droite du montant libres pour la date et le justificatif.';
              break;
            }
            if (L.bloque) break;
          }
        }
        // en dessous : un tableau à cheval sur ces colonnes serait décalé par l'insertion
        if (!L.bloque) {
          for (var q3 = totalRow + 1; q3 < v.length; q3++) {
            var inBlock = false;
            for (var cc = col; cc <= col + 3; cc++) if (String(v[q3][cc] === undefined ? '' : v[q3][cc]).trim() !== '') inBlock = true;
            var left = col > 0 && String(v[q3][col - 1]).trim() !== '';
            var right = col + 4 < v[q3].length && String(v[q3][col + 4]).trim() !== '';
            if (inBlock && (left || right)) {
              L.bloque = 'Sous la caisse (ligne ' + (q3 + 1) + '), un tableau déborde sur les colonnes de la caisse : insérer une ligne le décalerait. Déplace la caisse ou ce tableau.';
              break;
            }
          }
        }
        return L;
      }
    }
  }
  throw new Error('Tableau de caisse introuvable dans Modele_economique (il faut une cellule "Cash" puis "Total" en dessous).');
}

/** Lecture du tableau + synthèse */
function apiCompta_() {
  var L = comptaLocate_();
  var sh = L.sheet;
  var n = L.totalRow - L.start + 1;
  var rng = sh.getRange(L.start, L.col, n, 4);
  var vals = rng.getValues();
  var forms = rng.getFormulas();
  var notes = rng.getNotes();
  var lignes = [];
  var modes = [];
  for (var i = 0; i < n; i++) {
    var row = L.start + i;
    var label = String(vals[i][0]).trim();
    var montant = vals[i][1];
    if (!label && (montant === '' || montant === null)) continue;
    var key = norm_(label).replace(/s$/, '');
    if (row === L.totalRow) continue;
    var item = { row: row, label: label, montant: num_(montant), formule: forms[i][1] || '', date: plain_(vals[i][2]), justificatif: String(vals[i][3] || ''), note: notes[i][1] || '' };
    if (MODES.indexOf(key) >= 0) modes.push(item); else lignes.push(item);
  }
  var total = num_(sh.getRange(L.totalRow, L.col + 1).getValue());
  var depenses = lignes.filter(function (l) { return l.montant < 0; }).reduce(function (a, l) { return a + l.montant; }, 0);
  var recettesAutres = lignes.filter(function (l) { return l.montant > 0; }).reduce(function (a, l) { return a + l.montant; }, 0);
  var recettesModes = modes.reduce(function (a, m) { return a + m.montant; }, 0);
  return {
    lignes: lignes, modes: modes, total: total,
    depenses: depenses, recettes: recettesModes + recettesAutres,
    synthese: syntheseModeleEco_(L.ss),
    synthesePoles: synthesePoles_(L.ss),
    onglet: sh.getName(),
    factures: apiFactures_().filter(function (f) { return f.Type !== 'Devis' && f.Statut !== 'Payée' && f.Statut !== 'Annulée'; }),
    bloque: L.bloque,
    lien: 'https://docs.google.com/spreadsheets/d/' + L.ss.getId() + '/edit#gid=' + sh.getSheetId(),
    emplacement: sh.getName() + '!' + columnLetter_(L.col) + L.start + ':' + columnLetter_(L.col + 3) + L.totalRow
  };
}

function columnLetter_(c) {
  var s = '';
  while (c > 0) { var m = (c - 1) % 26; s = String.fromCharCode(65 + m) + s; c = Math.floor((c - 1) / 26); }
  return s;
}

/**
 * Le tableau des poles de la Synthese : recettes, cout intervenants, marge, et la colonne
 * de scenario optimal quand elle existe. On ne recopie plus l onglet entier : seul ce
 * tableau est utile dans l ERP, le reste se lit dans le Sheet.
 */
function synthesePoles_(ss) {
  var nom = getConfig_().ONGLET_SYNTHESE || 'Synthese';
  var sh = ss.getSheetByName(nom);
  if (!sh) sh = ss.getSheets().filter(function (x) { return norm_(x.getName()).indexOf('synth') === 0; })[0];
  if (!sh) return null;
  var disp = sh.getDataRange().getDisplayValues();
  var raw = sh.getDataRange().getValues();
  for (var r = 0; r < disp.length; r++) {
    if (norm_(disp[r][0]) !== 'pole') continue;
    var larg = disp[r].length;
    while (larg > 0 && String(disp[r][larg - 1]).trim() === '') larg--;
    if (larg < 2) continue;
    var nett = function (x) { return String(x).replace(/\s+/g, ' ').trim(); };
    var entetes = disp[r].slice(0, larg).map(nett);
    var lignes = [], total = null;
    for (var i = r + 1; i < disp.length; i++) {
      var lab = String(disp[i][0]).trim();
      if (!lab) break;
      var o = {
        cells: disp[i].slice(0, larg).map(nett),
        nums: raw[i].slice(0, larg).map(function (x) { return typeof x === 'number' ? x : null; })
      };
      if (norm_(lab).indexOf('total') === 0) { total = o; break; }
      lignes.push(o);
    }
    if (lignes.length) return { entetes: entetes, lignes: lignes, total: total };
  }
  return null;
}

/** Quelques chiffres clés de la synthèse, retrouvés par libellé */
function syntheseModeleEco_(ss) {
  // clé cherchée -> libellé affiché. La comparaison se fait sur le DEBUT du libellé,
  // pour survivre aux precisions ajoutees a la main (« TOTAL ANNEE (PAYE) »...).
  var wanted = [
    ['total annee reel', 'Total encaissé à ce jour'],
    ['total annee', 'Total année (recettes)'],
    ['caisse', 'Caisse'],
    ['total a payer aux intervenants', 'À payer aux intervenants'],
    ['total equipe', 'À payer aux intervenants'],
    ['total potentiel', 'Rémunération potentielle'],
    ['resultat optimal', 'Résultat si 7 élèves partout'],
    ['resultat previ net', 'Résultat net'],
    ['resultat net reel', 'Résultat net réel'],
    ['resultat net', 'Résultat net'],
    ['cours annuels', 'Recettes cours annuels'],
    ['prestations collectivites et ecoles', 'Recettes prestations'],
    ['stages vacances', 'Recettes stages'],
    ['charges de structure - previ', 'Charges de structure'],
    ['charges de structure - reel', 'Charges réelles'],
    ['total charges', 'Charges de structure']
  ];
  var out = [], seen = {};
  // la clé d'unicité est le libellé affiché : deux intitulés différents du Sheet qui
  // désignent la même chose (TOTAL EQUIPE / TOTAL A PAYER) ne font qu'une ligne
  var pose = function (cle, libelle, valeur) {
    if (seen[libelle] || typeof valeur !== 'number') return;
    seen[libelle] = true; out.push({ label: libelle, valeur: valeur });
  };
  ss.getSheets().forEach(function (sh) {
    var v = sh.getDataRange().getValues();
    if (norm_(v.map(function (r) { return r.join(' '); }).join(' ')).indexOf('synthese annuelle') < 0) return;
    for (var r = 0; r < v.length; r++) {
      for (var c = 0; c < v[r].length; c++) {
        var k = norm_(v[r][c]).replace(/'/g, ' ');
        if (!k) continue;
        for (var w = 0; w < wanted.length; w++) {
          if (k.indexOf(wanted[w][0]) !== 0) continue;
          // la valeur est juste a droite, sinon (cas « TOTAL CHARGES » en legende) juste a gauche
          var droite = c + 1 < v[r].length ? v[r][c + 1] : null;
          var gauche = c > 0 ? v[r][c - 1] : null;
          if (typeof droite === 'number') pose(wanted[w][0], wanted[w][1], droite);
          else if (typeof gauche === 'number') pose(wanted[w][0], wanted[w][1], gauche);
          break;
        }
      }
    }
  });
  return out;
}

function assertWritable_(L) {
  if (L.bloque) throw new Error(L.bloque);
}

/**
 * Après insertion / suppression : le Total doit couvrir tout le bloc (du début jusqu'à la ligne au-dessus de Total).
 * Seulement si le Total est un SUM simple (ou une valeur qui collait à la somme).
 */
function fixTotal_(L, insertedRow) {
  var sh = L.sheet;
  var cell = sh.getRange(L.totalRow, L.col + 1);
  var f = cell.getFormula();
  var colL = columnLetter_(L.col + 1);
  var fm = f && f.match(/^=\s*SUM\(\s*\$?([A-Z]+)\$?(\d+)\s*:\s*\$?([A-Z]+)\$?(\d+)\s*\)\s*$/i);
  var a = fm ? Math.min(+fm[2], +fm[4]) : L.start;
  if (insertedRow && insertedRow < a) a = insertedRow;
  var nf = '=SUM(' + colL + a + ':' + colL + (L.totalRow - 1) + ')';
  if (f.replace(/\$/g, '').replace(/\s/g, '').toUpperCase() !== nf.toUpperCase()) {
    cell.setFormula(nf);
    logAction_('compta', 'Formule Total : ' + (f || '(valeur)') + ' -> ' + nf);
  }
}

/** Ajoute une dépense (ligne au-dessus de Cash) */
function apiAjouterDepense_(p) {
  var lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    var montant = Math.abs(num_(p.montant));
    if (!p.libelle || !montant) throw new Error('Libellé et montant obligatoires.');
    var L = comptaLocate_();
    assertWritable_(L);
    var sh = L.sheet;
    var ins = L.cashRow;
    sh.getRange(ins, L.col, 1, 4).insertCells(SpreadsheetApp.Dimension.ROWS);
    var d = parseDate_(p.date) || new Date();
    var just = /^https:\/\//.test(String(p.justificatif || '')) ? p.justificatif : '';
    sh.getRange(ins, L.col, 1, 4).setValues([[safeCell_(String(p.libelle)), -montant, d, just]]);
    sh.getRange(L.cashRow, L.col + 2).setNumberFormat('dd/mm/yyyy');
    sh.getRange(L.cashRow, L.col + 1).setNumberFormat(sh.getRange(L.start, L.col + 1).getNumberFormat() || '#,##0 €');
    SpreadsheetApp.flush();
    fixTotal_(comptaLocate_(), ins);
    logAction_('dépense', p.libelle + ' -' + montant);
    return true;
  } finally { lock.releaseLock(); }
}

/**
 * Modifie une ligne de la caisse : p = { row, label (contrôle), libelle, sens: 'depense'|'entree', montant (positif), date }
 * Lignes Cash / Chèque / Virement : seul le montant change (libellé fixe). L'ancienne valeur est gardée en note.
 */
function apiModifierLigneCompta_(p) {
  var lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    var L = comptaLocate_();
    assertWritable_(L);
    var sh = L.sheet;
    var row = Number(p.row);
    if (!(row >= L.start && row < L.totalRow)) throw new Error('Cette ligne ne peut pas être modifiée ici.');
    var ancienLabel = String(sh.getRange(row, L.col).getValue());
    if (norm_(ancienLabel) !== norm_(p.label)) throw new Error('Le tableau a bougé, recharge la page.');
    var brut = String(p.montant === undefined ? '' : p.montant).replace(/\s/g, '').replace(',', '.');
    if (brut === '' || isNaN(Number(brut))) throw new Error('Montant invalide.');
    var estMode = MODES.indexOf(norm_(ancienLabel).replace(/s$/, '')) >= 0;
    var montant = Math.abs(Number(brut));
    if (!estMode && p.sens === 'depense') montant = -montant;
    var cell = sh.getRange(row, L.col + 1);
    var avant = num_(cell.getValue());
    var libelle = estMode ? ancienLabel : String(p.libelle || '').trim();
    if (!libelle) throw new Error('Libellé obligatoire.');
    if (!estMode && libelle !== ancienLabel) sh.getRange(row, L.col).setValue(safeCell_(libelle));
    if (Math.abs(avant - montant) > 0.001) {
      cell.setValue(montant);
      var note = cell.getNote();
      cell.setNote((note ? note + '\n' : '') + fmtDate_(new Date()) + ' : ' + avant + ' € -> ' + montant + ' € (modifié par ' + currentUserLabel_() + ')');
    }
    var d = parseDate_(p.date);
    if (d) sh.getRange(row, L.col + 2).setValue(d).setNumberFormat('dd/mm/yyyy');
    SpreadsheetApp.flush();
    fixTotal_(comptaLocate_());
    logAction_('ligne compta modifiée', ancienLabel + ' ' + avant + ' -> ' + libelle + ' ' + montant);
    return true;
  } finally { lock.releaseLock(); }
}

/** Supprime une ligne de dépense (jamais les lignes Cash / Chèque / Virement / Total) */
function apiSupprimerLigneCompta_(p) {
  var lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    var L = comptaLocate_();
    assertWritable_(L);
    var row = Number(p.row);
    var label = String(L.sheet.getRange(row, L.col).getValue());
    if (norm_(label) !== norm_(p.label)) throw new Error('Le tableau a bougé, recharge la page.');
    if (row < L.start || row >= L.cashRow) throw new Error('Cette ligne ne peut pas être supprimée ici.');
    L.sheet.getRange(row, L.col, 1, 4).deleteCells(SpreadsheetApp.Dimension.ROWS);
    SpreadsheetApp.flush();
    fixTotal_(comptaLocate_());
    logAction_('suppression ligne compta', label);
    return true;
  } finally { lock.releaseLock(); }
}

/** Ajoute (ou retire si négatif) une recette sur un mode de paiement */
function apiAjouterRecette_(p) {
  var lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    var mode = norm_(p.mode).replace(/s$/, '');
    if (mode === 'especes' || mode === 'espece' || mode === 'liquide') mode = 'cash';
    if (MODES.indexOf(mode) < 0) throw new Error('Mode inconnu : Cash, Chèque ou Virement.');
    var montant = num_(p.montant);
    if (!montant) throw new Error('Montant obligatoire.');
    var L = comptaLocate_();
    assertWritable_(L);
    var row = L.modeRows[mode];
    if (!row) throw new Error('Ligne "' + p.mode + '" introuvable dans le tableau.');
    var cell = L.sheet.getRange(row, L.col + 1);
    var f = cell.getFormula();
    var base = f ? f.replace(/^=/, '') : String(num_(cell.getValue()));
    var add = montant >= 0 ? '+' + montant : String(montant);
    cell.setFormula('=' + base + add);
    var d = parseDate_(p.date) || new Date();
    var note = cell.getNote();
    var line = fmtDate_(d) + ' ' + (montant >= 0 ? '+' : '') + montant + ' € ' + (p.libelle || '');
    cell.setNote((note ? note + '\n' : '') + line);
    L.sheet.getRange(row, L.col + 2).setValue(d).setNumberFormat('dd/mm/yyyy');
    SpreadsheetApp.flush();
    fixTotal_(comptaLocate_());
    logAction_('recette', line + ' (' + mode + ')');
    return true;
  } finally { lock.releaseLock(); }
}

/**
 * Total de la caisse corrigé à la main (l'info vient souvent de l'oral : "il y a 2 350 € en caisse").
 * On ne casse pas le =SUM : on met l'écart sur une ligne unique « Ajustement caisse » (au-dessus de Cash),
 * dont la cellule garde l'historique (=-120+45) et une note datée à chaque correction.
 */
function apiModifierCaisse_(p) {
  var lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    var voulu = num_(p.total);
    if (p.total === '' || p.total === null || p.total === undefined || isNaN(voulu)) throw new Error('Indique le nouveau total de la caisse.');
    var L = comptaLocate_();
    assertWritable_(L);
    var sh = L.sheet;
    var actuel = num_(sh.getRange(L.totalRow, L.col + 1).getValue());
    var ecart = Math.round((voulu - actuel) * 100) / 100;
    if (Math.abs(ecart) < 0.005) return { total: actuel, ecart: 0 };
    var d = new Date();
    var ligne = fmtDate_(d) + ' : total ' + actuel + ' € -> ' + voulu + ' € (' + (ecart > 0 ? '+' : '') + ecart + ' €)' + (p.raison ? ' ' + String(p.raison).slice(0, 120) : '') + ' par ' + currentUserLabel_();
    // ligne d'ajustement existante ?
    var n = L.cashRow - L.start;
    var labels = n > 0 ? sh.getRange(L.start, L.col, n, 1).getValues() : [];
    var row = -1;
    for (var i = 0; i < labels.length; i++) if (norm_(labels[i][0]) === 'ajustement caisse') { row = L.start + i; break; }
    if (row < 0) {
      row = L.cashRow;
      sh.getRange(row, L.col, 1, 4).insertCells(SpreadsheetApp.Dimension.ROWS);
      sh.getRange(row, L.col, 1, 3).setValues([['Ajustement caisse', ecart, d]]);
      sh.getRange(row, L.col + 1).setNumberFormat(sh.getRange(L.start, L.col + 1).getNumberFormat() || '#,##0 €');
    } else {
      var cell = sh.getRange(row, L.col + 1);
      var f = cell.getFormula();
      var base = f ? f.replace(/^=/, '') : String(num_(cell.getValue()));
      cell.setFormula('=' + base + (ecart >= 0 ? '+' + ecart : String(ecart)));
      sh.getRange(row, L.col + 2).setValue(d);
    }
    sh.getRange(row, L.col + 2).setNumberFormat('dd/mm/yyyy');
    var c2 = sh.getRange(row, L.col + 1);
    var note = c2.getNote();
    c2.setNote((note ? note + '\n' : '') + ligne);
    SpreadsheetApp.flush();
    fixTotal_(comptaLocate_(), row);
    logAction_('caisse', ligne);
    return { total: voulu, ecart: ecart };
  } finally { lock.releaseLock(); }
}

/** Upload d'un justificatif (photo / PDF) : renvoie le lien Drive */
/** Sous-dossier des justificatifs, créé au besoin. '' = racine des justificatifs. */
function dossierJustif_(nom) {
  var c = getConfig_();
  if (!c.ID_DOSSIER_JUSTIFICATIFS) throw new Error('Dossier justificatifs non configuré (relance setup).');
  var racine = DriveApp.getFolderById(c.ID_DOSSIER_JUSTIFICATIFS);
  if (!nom) return racine;
  var it = racine.getFoldersByName(nom);
  return it.hasNext() ? it.next() : racine.createFolder(nom);
}

var SOUS_DOSSIERS = { fournisseurs: 'Factures fournisseurs', intervenants: 'Factures intervenants', caisse: 'Justificatifs caisse' };

/**
 * Dépôt d'un ou plusieurs justificatifs.
 * p = { name, mime, data } pour un fichier, ou p = { fichiers: [{name, mime, data}], dossier: 'fournisseurs' }.
 * Renvoie { url, id, fichiers: [{url, id, nom}] } : url reste le premier fichier, pour l'existant.
 */
function apiUploadJustificatif_(p) {
  var liste = p.fichiers && p.fichiers.length ? p.fichiers : [{ name: p.name, mime: p.mime, data: p.data }];
  if (liste.length > 10) throw new Error('10 fichiers maximum à la fois.');
  var folder = dossierJustif_(SOUS_DOSSIERS[p.dossier] || '');
  var jour = Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd');
  var out = liste.map(function (f) {
    var bytes = Utilities.base64Decode(f.data);
    if (bytes.length > 15 * 1024 * 1024) throw new Error('Fichier trop lourd (15 Mo max) : ' + (f.name || ''));
    var name = jour + ' - ' + String(f.name || 'justificatif').replace(/[\\/:*?"<>|]/g, '_');
    var file = folder.createFile(Utilities.newBlob(bytes, f.mime || 'application/octet-stream', name));
    return { url: file.getUrl(), id: file.getId(), nom: name };
  });
  logAction_('justificatif', out.length + ' fichier(s) dans ' + folder.getName());
  return { url: out[0].url, id: out[0].id, fichiers: out };
}

// ---------- Factures à payer (intervenants, auto-entreprise, fournisseurs) ----------

/** Onglet Achats du classeur ERP (créé au besoin) */
function achatsSheet_() {
  var ss = erpSS_();
  var sh = ss.getSheetByName('Achats');
  if (!sh) { sh = ss.insertSheet('Achats'); sh.getRange(1, 1, 1, TABLES.Achats.length).setValues([TABLES.Achats]); }
  return sh;
}

/** Liste des personnes / sociétés qu'on paie (Config INTERVENANTS) */
function beneficiaires_() {
  var c = getConfig_();
  return String(c.INTERVENANTS || '').split(/[;\n]/).map(function (s) { return s.trim(); }).filter(String);
}

function apiAchats_() {
  achatsSheet_();
  var auj = new Date(); auj.setHours(0, 0, 0, 0);
  var liste = readTable_('Achats').map(function (r) {
    r.montant = num_(r.Montant);
    r.paye = /^pay/i.test(String(r.Statut || ''));
    var ech = parseDate_(r.Echeance);
    r.retard = !r.paye && !!ech && ech < auj;
    r.jours = ech ? Math.round((ech - auj) / 86400000) : null;
    r.pieces = [r.Lien].concat(String(r.Pieces || '').split('|')).map(function (x) { return String(x || '').trim(); })
      .filter(function (x) { return /^https:\/\//.test(x); });
    return r;
  }).sort(function (a, b) {
    if (a.paye !== b.paye) return a.paye ? 1 : -1;
    var da = parseDate_(a.Echeance) || parseDate_(a.Date), db = parseDate_(b.Echeance) || parseDate_(b.Date);
    return (da ? da.getTime() : 0) - (db ? db.getTime() : 0);
  });
  var aPayer = liste.filter(function (r) { return !r.paye; });
  return {
    liste: liste,
    beneficiaires: beneficiaires_(),
    aPayer: aPayer.reduce(function (a, r) { return a + Math.abs(r.montant); }, 0),
    nbAPayer: aPayer.length,
    retards: aPayer.filter(function (r) { return r.retard; }).length,
    payeTotal: liste.filter(function (r) { return r.paye; }).reduce(function (a, r) { return a + Math.abs(r.montant); }, 0)
  };
}

/** Enregistre (ou modifie) une facture à payer */
function apiSaveAchat_(p) {
  achatsSheet_();
  var montant = Math.abs(num_(p.Montant));
  if (!String(p.Beneficiaire || '').trim()) throw new Error('Choisis qui doit être payé.');
  if (!montant) throw new Error('Montant obligatoire.');
  var rec = {
    ID: p.ID || uid_(),
    Date: parseDate_(p.Date) || new Date(),
    Beneficiaire: String(p.Beneficiaire).trim(),
    Type: p.Type || 'Intervenant',
    Objet: p.Objet || '',
    Montant: montant,
    Echeance: parseDate_(p.Echeance) || '',
    Statut: p.Statut || 'À payer',
    PayeLe: parseDate_(p.PayeLe) || '',
    Mode: p.Mode || '',
    Lien: /^https:\/\//.test(String(p.Lien || '')) ? p.Lien : '',
    // pièces jointes supplémentaires, séparées par « | » (une facture fournisseur peut tenir en plusieurs pages)
    Pieces: String(p.Pieces || '').split('|').map(function (x) { return x.trim(); })
      .filter(function (x) { return /^https:\/\//.test(x); }).join(' | '),
    Notes: p.Notes || '',
    MajPar: currentUserLabel_(), MajLe: new Date()
  };
  upsertRow_('Achats', rec, 'ID');
  logAction_('facture à payer', rec.Beneficiaire + ' ' + montant + ' € (' + rec.Statut + ')');
  return rec.ID;
}

/**
 * Marque une facture payée. p = { id, mode, date, caisse: true/false }
 * caisse = true : la dépense est aussi écrite dans le tableau de caisse (Compte aussi).
 */
function apiPayerAchat_(p) {
  achatsSheet_();
  var a = readTable_('Achats').filter(function (r) { return String(r.ID) === String(p.id); })[0];
  if (!a) throw new Error('Facture introuvable.');
  if (/^pay/i.test(String(a.Statut || ''))) throw new Error('Cette facture est déjà notée payée.');
  var d = parseDate_(p.date) || new Date();
  var montant = Math.abs(num_(a.Montant));
  var libelle = String(a.Beneficiaire) + (a.Objet ? ' · ' + a.Objet : '');
  if (p.caisse) apiAjouterDepense_({ libelle: libelle, montant: montant, date: d, justificatif: a.Lien || '' });
  upsertRow_('Achats', { ID: a.ID, Statut: 'Payée', PayeLe: d, Mode: p.mode || a.Mode || '', MajPar: currentUserLabel_(), MajLe: new Date() }, 'ID');
  logAction_('paiement', libelle + ' -' + montant + (p.caisse ? ' (caisse)' : ' (hors caisse)'));
  return { montant: montant, caisse: !!p.caisse };
}

function apiSupprimerAchat_(p) {
  achatsSheet_();
  if (!deleteRowById_('Achats', p.id)) throw new Error('Facture introuvable.');
  logAction_('facture à payer supprimée', String(p.id));
  return true;
}
