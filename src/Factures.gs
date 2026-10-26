/**
 * Factures.gs : génère un DEVIS ou une FACTURE à partir de Modele_facture
 * (copie du Doc + remplissage + PDF), suivis dans l'onglet Factures.
 * Nomenclature (comme DEV-2026-001_Villeneuve.pdf) :
 *   DEV-2026-002_<Client>.pdf  /  FAC-2026-001_<Client>.pdf  à la racine du dossier Devis / Facture,
 *   le Google Doc modifiable dans le sous-dossier « Modifiables (Docs) ».
 */

function eur_(n) {
  var v = Math.round(num_(n) * 100) / 100;
  var s = v.toFixed(2).split('.');
  var ent = s[0].replace(/\B(?=(\d{3})+(?!\d))/g, '\u202f');
  return ent + ',' + s[1] + ' €';
}

function apiFactures_() {
  return readTable_('Factures').map(function (f) {
    f.Type = f.Type || (/^DEV/i.test(String(f.Numero)) ? 'Devis' : 'Facture');
    return f;
  }).sort(function (a, b) {
    var da = parseDate_(a.Date), db = parseDate_(b.Date);
    return ((db ? db.getTime() : 0) - (da ? da.getTime() : 0)) || String(b.Numero).localeCompare(String(a.Numero));
  });
}

function nextNumero_(n) {
  var m = String(n).match(/^(.*?)(\d+)$/);
  if (!m) return n + '-1';
  var next = String(+m[2] + 1);
  while (next.length < m[2].length) next = '0' + next;
  return m[1] + next;
}

function dossiersFactures_() {
  var c = getConfig_();
  var root = c.ID_DOSSIER_DEVIS ? DriveApp.getFolderById(c.ID_DOSSIER_DEVIS) : DriveApp.getRootFolder();
  return { root: root, docs: getOrCreateFolder_(root, 'Modifiables (Docs)') };
}

/** Numéro du prochain devis / de la prochaine facture (DEV-2026-002, FAC-2026-001) */
function numeroDoc_(type, c) {
  var an = Utilities.formatDate(new Date(), TZ, 'yyyy');
  if (type === 'devis') return String(c.PROCHAIN_NUM_DEVIS || ('DEV-' + an + '-001')).trim();
  var n = String(c.PROCHAIN_NUM_FACTURE || ('FAC-' + an + '-001')).trim();
  if (/^\d{4}-\d+$/.test(n)) n = 'FAC-' + n; // ancien format 2026-001
  return n;
}

/** "Mairie de Villeneuve" -> "Villeneuve" (pour le nom de fichier) */
function nomCourt_(client) {
  var t = String(client || '').trim()
    .replace(/^(mairie|commune|ville|ecole|école|groupe scolaire|college|collège|centre de loisirs|accueil de loisirs|association|communauté de communes)\s+(de\s+la\s+|de\s+l'|du\s+|des\s+|de\s+|d')?/i, '');
  t = t.replace(/[\\/:*?"<>|#%]/g, '').replace(/['’]/g, '-').replace(/\s+/g, '-').replace(/-+/g, '-').replace(/^-|-$/g, '');
  return (t || 'Client').slice(0, 40);
}

function ensureColumns_(name) {
  var sh = erpSheet_(name);
  var h = headersOf_(name);
  (TABLES[name] || []).forEach(function (col) {
    if (h.indexOf(col) < 0) { sh.getRange(1, h.length + 1).setValue(col); h.push(col); }
  });
}

/**
 * p = { type: 'devis' | 'facture', client, service, adresse, cpVille, siret, engagement, periode,
 *       date (yyyy-mm-dd), conditions?, depuisDevis? (ID), lignes: [{designation, detail, qte, pu}] }
 */
function apiCreerFacture_(p) {
  var lock = LockService.getScriptLock(); lock.waitLock(30000);
  try {
    var c = getConfig_(true);
    var type = p.type === 'devis' ? 'devis' : 'facture';
    if (!p.client) throw new Error('Client obligatoire.');
    var lignes = (p.lignes || []).filter(function (l) { return l.designation && num_(l.qte) && num_(l.pu); });
    if (!lignes.length) throw new Error('Au moins une ligne avec désignation, quantité et prix.');
    var numero = numeroDoc_(type, c);
    var dateF = parseDate_(p.date) || new Date();
    var jours = type === 'devis' ? (num_(c.VALIDITE_DEVIS_JOURS) || 30) : (num_(c.ECHEANCE_FACTURE_JOURS) || 30);
    var ech = new Date(dateF.getTime() + jours * 86400000);
    var total = lignes.reduce(function (a, l) { return a + num_(l.qte) * num_(l.pu); }, 0);
    var court = p.nomCourt ? nomCourt_(p.nomCourt) : nomCourt_(p.client);
    var nom = numero + '_' + court;

    var f = dossiersFactures_();
    var copy = DriveApp.getFileById(c.ID_MODELE_FACTURE).makeCopy(nom, f.docs);
    var doc = DocumentApp.openById(copy.getId());
    var body = doc.getBody();

    function rep(pattern, value) { body.replaceText(pattern, String(value || '').replace(/\$/g, '\\$')); }
    rep('N° ?\\d{4}-\\d{3}', 'N° ' + numero);
    rep('\\[Nom de la collectivité\\]', p.client);
    rep('\\[Service\\]', p.service || '');
    rep('\\[Adresse\\]', p.adresse || '');
    rep('\\[Code postal et ville\\]', p.cpVille || '');
    rep('\\[SIRET du client\\]', p.siret || '-');
    rep('IBAN \\[à compléter\\]', 'IBAN ' + (c.IBAN || ''));
    rep('BIC \\[à compléter\\]', 'BIC ' + (c.BIC || ''));
    fillTable_(body, lignes, total);

    if (type === 'facture') {
      rep('Émise le \\[JJ/MM/AAAA\\]', 'Émise le ' + fmtDate_(dateF));
      rep('Échéance le \\[JJ/MM/AAAA\\]', 'Échéance le ' + fmtDate_(ech));
      rep('\\[bon de commande\\]', p.engagement || '-');
      rep('\\[période\\]', p.periode || '');
    } else {
      rep('^\\s*FACTURE\\s*$', 'DEVIS');
      rep('Émise le \\[JJ/MM/AAAA\\]', 'Établi le ' + fmtDate_(dateF));
      rep('Échéance le \\[JJ/MM/AAAA\\]', 'Valable ' + jours + ' jours');
      rep('FACTURÉ À', 'DEVIS ÉTABLI POUR');
      rep("N° d'engagement \\[bon de commande\\]", p.engagement ? "N° d'engagement " + p.engagement : '');
      rep('Prestation réalisée \\[période\\]', p.periode ? 'Prestation : ' + p.periode : '');
      rep('\\[période\\]', p.periode || '');
      rep('TOTAL À PAYER', 'TOTAL DU DEVIS');
      blocsDevis_(body, c, p.conditions);
    }
    doc.saveAndClose();

    var pdf = f.root.createFile(copy.getAs('application/pdf').setName(nom + '.pdf'));
    ensureColumns_('Factures');
    var details = { service: p.service || '', adresse: p.adresse || '', cpVille: p.cpVille || '', siret: p.siret || '', engagement: p.engagement || '',
      periode: p.periode || '', nomCourt: court, lignes: lignes, depuisDevis: p.depuisDevis || '' };
    var rec = {
      ID: uid_(), Numero: numero, Date: dateF, Client: p.client, Objet: lignes.map(function (l) { return l.designation; }).join(' / ').slice(0, 200),
      Montant: total, Echeance: ech, Statut: type === 'devis' ? 'Envoyé' : 'Émise', PayeeLe: '', LienDoc: copy.getUrl(), LienPdf: pdf.getUrl(),
      Type: type === 'devis' ? 'Devis' : 'Facture', Details: JSON.stringify(details)
    };
    upsertRow_('Factures', rec, 'ID');
    setConfig_(type === 'devis' ? 'PROCHAIN_NUM_DEVIS' : 'PROCHAIN_NUM_FACTURE', nextNumero_(numero));
    if (p.depuisDevis) upsertRow_('Factures', { ID: p.depuisDevis, Statut: 'Facturé' }, 'ID');
    logAction_(type, numero + ' ' + p.client + ' ' + total);
    return { numero: numero, type: type, doc: copy.getUrl(), pdf: pdf.getUrl(), total: total, fichier: nom + '.pdf' };
  } finally { lock.releaseLock(); }
}

/** Devis : le cadre RÈGLEMENT devient CONDITIONS, MENTIONS LÉGALES devient BON POUR ACCORD */
function blocsDevis_(body, c, conditions) {
  var tables = body.getTables();
  for (var i = 0; i < tables.length; i++) {
    var t = tables[i];
    if (!t.getNumRows() || t.getRow(0).getNumCells() < 2) continue;
    if (norm_(t.getRow(0).getCell(0).getText()).indexOf('reglement') !== 0) continue;
    var cond = conditions || ('Devis valable ' + (num_(c.VALIDITE_DEVIS_JOURS) || 30) + ' jours. Lieu : sur place, ' + c.ADRESSE_STRUCTURE +
      '. Facturation à l\'issue de la prestation, règlement par virement (Chorus Pro pour les clients publics). Toute séance annulée moins de 48 heures à l\'avance est due.');
    setCellParas_(t.getRow(0).getCell(0), [{ t: 'CONDITIONS', b: true }, { t: cond }, { t: 'TVA non applicable, article 261-7-1° du Code général des impôts.', i: true }]);
    setCellParas_(t.getRow(0).getCell(1), [{ t: 'BON POUR ACCORD', b: true }, { t: 'Date, signature et cachet, précédés de la mention « Bon pour accord »' }, { t: ' ' }, { t: ' ' },
      { t: 'Pour L\'Ain\'stant Activite : ' + (c.SIGNATAIRE || '') }]);
    return;
  }
}

function setCellParas_(cell, paras) {
  for (var i = cell.getNumChildren() - 1; i >= 1; i--) cell.getChild(i).removeFromParent();
  var first = cell.getChild(0).asParagraph();
  paras.forEach(function (x, k) {
    var par = k === 0 ? first : cell.appendParagraph('');
    par.setText(x.t || ' ');
    var tx = par.editAsText();
    tx.setBold(!!x.b); tx.setItalic(!!x.i);
  });
}

function fillTable_(body, lignes, total) {
  var tables = body.getTables();
  var t = null;
  for (var i = 0; i < tables.length; i++) {
    if (norm_(tables[i].getRow(0).getText()).indexOf('designation') >= 0) { t = tables[i]; break; }
  }
  if (!t) throw new Error('Tableau DÉSIGNATION introuvable dans le modèle de facture.');
  // lignes "article" = entre l'en-tête et la ligne Sous-total
  var subRow = -1;
  for (var r = 1; r < t.getNumRows(); r++) if (norm_(t.getRow(r).getText()).indexOf('sous-total') >= 0) { subRow = r; break; }
  if (subRow < 0) throw new Error('Ligne Sous-total introuvable dans le modèle.');
  var slots = subRow - 1;
  // ajoute des lignes si besoin (copie de la dernière ligne vide)
  while (slots < lignes.length) { t.insertTableRow(subRow, t.getRow(subRow - 1).copy()); subRow++; slots++; }
  for (var k = 0; k < slots; k++) {
    var row = t.getRow(1 + k);
    var l = lignes[k];
    setCell_(row.getCell(0), l ? l.designation : '', l ? l.detail : '');
    setCell_(row.getCell(1), l ? String(num_(l.qte)) : '');
    setCell_(row.getCell(2), l ? eur_(l.pu) : '');
    setCell_(row.getCell(3), l ? eur_(num_(l.qte) * num_(l.pu)) : '');
  }
  // Sous-total et TOTAL : le montant est dans la cellule qui suit le libellé
  for (var r2 = subRow; r2 < t.getNumRows(); r2++) {
    var rw = t.getRow(r2);
    for (var cI = 0; cI < rw.getNumCells() - 1; cI++) {
      var tx = norm_(rw.getCell(cI).getText());
      if (tx.indexOf('sous-total') >= 0 || tx.indexOf('total a payer') >= 0 || tx.indexOf('total du devis') >= 0) {
        var cell = rw.getCell(cI + 1);
        var bold = tx.indexOf('total a payer') >= 0;
        setCell_(cell, eur_(total));
        if (bold) cell.editAsText().setBold(true);
        break;
      }
    }
  }
}

function setCell_(cell, main, detail) {
  var n = cell.getNumChildren();
  var p0 = cell.getChild(0).asParagraph();
  var wasBold = p0.getText().length ? p0.editAsText().isBold(0) : false;
  for (var i = n - 1; i >= 1; i--) cell.getChild(i).removeFromParent();
  p0.setText(main || ' ');
  if (wasBold && main) p0.editAsText().setBold(0, main.length - 1, true);
  if (detail) {
    var p1 = cell.appendParagraph(detail);
    p1.editAsText().setBold(false).setFontSize(9);
  }
}

/** PDF présents dans le dossier Devis / Facture du Drive (y compris ceux faits à la main) */
function apiDossierDevis_() {
  var c = getConfig_();
  if (!c.ID_DOSSIER_DEVIS) return [];
  var suivis = {};
  readTable_('Factures').forEach(function (f) { suivis[norm_(f.Numero)] = f.Statut || ''; });
  var out = [];
  try {
    var folder = DriveApp.getFolderById(c.ID_DOSSIER_DEVIS);
    var it = folder.getFiles();
    while (it.hasNext() && out.length < 60) {
      var f = it.next();
      var nom = f.getName();
      if (!/\.pdf$/i.test(nom)) continue;
      var num = (nom.match(/^(DEV|FAC)-\d{4}-\d+/i) || [''])[0];
      var type = /^DEV/i.test(num) || /devis/i.test(nom) ? 'Devis' : 'Facture';
      out.push({ nom: nom, numero: num, type: type, url: f.getUrl(), date: fmtDate_(f.getLastUpdated()),
        modif: f.getLastUpdated().getTime(), statut: num ? (suivis[norm_(num)] || '') : '' });
    }
  } catch (e) { return []; }
  out.sort(function (a, b) { return b.modif - a.modif; });
  return out;
}

/** Changement de statut ; si payée et option, ajoute la recette en Virement dans la caisse */
function apiStatutFacture_(p) {
  var f = readTable_('Factures').filter(function (x) { return String(x.ID) === String(p.id); })[0];
  if (!f) throw new Error('Facture introuvable.');
  var upd = { ID: p.id, Statut: p.statut };
  if (p.statut === 'Payée') upd.PayeeLe = new Date();
  upsertRow_('Factures', upd, 'ID');
  if (p.statut === 'Payée' && p.ajouterCaisse) {
    apiAjouterRecette_({ mode: p.mode || 'Virement', montant: num_(f.Montant), libelle: 'Facture ' + f.Numero + ' ' + f.Client });
  }
  logAction_('facture statut', f.Numero + ' -> ' + p.statut);
  return true;
}
