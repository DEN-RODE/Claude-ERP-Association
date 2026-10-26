/**
 * Mails.gs : brouillons Gmail (rien ne part sans validation humaine),
 * déclencheurs d'inscription, tâche quotidienne (relances optionnelles + récap équipe).
 */

/**
 * Signature commune à tous les mails sortants : même formule, même adresse,
 * téléphone + mail + site. Tout vient de l'onglet Config.
 */
function signature_() {
  var c = getConfig_();
  var adresse = String(c.ADRESSE_STRUCTURE || '').trim();
  var contact = [c.TEL_STRUCTURE, c.EMAIL_STRUCTURE, c.SITE_WEB].filter(String).join(' · ');
  return '\n\nMerci à vous, et à très vite ' + L_('lieu') + '.\n\nL\'équipe de ' + c.NOM_STRUCTURE + '\n' + adresse + '\n' + contact;
}

function infosPaiement_(mode) {
  var c = getConfig_();
  var m = norm_(mode);
  var lines = [];
  if (!m || m.indexOf('virement') >= 0) {
    lines.push('Par virement : titulaire ' + c.IBAN_TITULAIRE + ', IBAN ' + c.IBAN + ', BIC ' + c.BIC + ' (merci d\'indiquer le nom ' + Ldu_('membre') + ' en référence).');
  }
  if (!m || m.indexOf('cheque') >= 0) lines.push('Par chèque : à l\'ordre de ' + (c.IBAN_TITULAIRE || c.NOM_ASSO || c.NOM_STRUCTURE) + ', à remettre ' + L_('lieu') + '.');
  if (!m || m.indexOf('espece') >= 0 || m.indexOf('cash') >= 0) lines.push('En espèces : directement ' + L_('lieu') + ', un reçu vous sera remis.');
  return lines.join('\n');
}

function tarifsEleve_(e) {
  if (e.source === 'cours') {
    var m = matchCours_(e.creneaux, coursList_());
    return m.map(function (k) { return { label: k.libelle + ', ' + k.jour + ' ' + k.horaire, prix: k.tarif }; });
  }
  var s = matchStages_(e.stage, stagesList_());
  return s.map(function (k) { return { label: 'Stage ' + k.nom + ' (' + k.dates + '), ' + k.public, prix: k.tarif }; });
}

function corpsConfirmation_(e) {
  var c = getConfig_();
  var t = tarifsEleve_(e);
  var total = t.reduce(function (a, x) { return a + x.prix; }, 0);
  var bonjour = 'Bonjour' + (e.responsable ? ' ' + String(e.responsable).trim() : '') + ',';
  var qui = e.prenom + ' ' + e.nom;
  var lignes = t.length ? t.map(function (x) { return '  • ' + x.label + ' : ' + x.prix + ' €'; }).join('\n') : '  • ' + (e.creneaux || e.stage);
  var txt = bonjour + '\n\nMerci, nous avons bien reçu l\'inscription de ' + qui + ' à ' + c.NOM_STRUCTURE + '.\n\n' +
    'Récapitulatif :\n' + lignes + '\n';
  if (t.length > 1) txt += '  Total : ' + total + ' € (réduction fratrie éventuelle appliquée à la validation)\n';
  if (e.source === 'cours') {
    txt += '\nLe tarif annuel comprend l\'adhésion à l\'association.';
    if (norm_(e.essai).indexOf('oui') === 0) txt += ' Le premier cours sert de cours d\'essai gratuit : l\'inscription est confirmée après ce cours.';
    txt += '\n';
  } else {
    txt += '\nL\'adhésion à l\'association (' + c.ADHESION + ' € par an et par enfant) s\'ajoute au prix du stage pour les nouveaux inscrits.\n';
  }
  txt += '\nRèglement :\n' + infosPaiement_(e.modePaiement) + '\n';
  if (e.manquants && e.manquants.length) {
    txt += '\nIl nous manque encore : ' + e.manquants.join(', ') + '. Vous pouvez simplement répondre à ce mail avec ces informations.\n';
  }
  txt += '\nRendez-vous ' + c.ADRESSE_STRUCTURE + '. Prévoir une tenue souple et une gourde.';
  txt += signature_();
  var sujet = (e.source === 'cours' ? 'Inscription confirmée' : 'Inscription au stage confirmée') + ' : ' + e.prenom + ' · ' + c.NOM_STRUCTURE;
  return { sujet: sujet, corps: txt };
}

/**
 * Une inscription offerte, un tarif « à la carte », une réduction fratrie ou commune :
 * dans ces cas le tarif du créneau n'est PAS le montant dû. On ne met alors aucun chiffre
 * dans le mail, plutôt que d'en annoncer un faux. C'est la règle centrale de tout ce
 * fichier : un montant incertain ne sort jamais vers une famille.
 *
 * REDUC_A_VERIFIER se complète depuis l'onglet Config (clé REDUC_MOTIFS) pour ajouter
 * les motifs propres à la structure, par exemple le nom d'une commune conventionnée.
 */
var REDUC_A_VERIFIER = ['fratrie', 'demi ?annee', 'demi-annee', 'a la carte', 'preferentiel', 'convention', 'tarif reduit', 'quotient'];

function motifsReduc_() {
  var sup = String(getConfig_().REDUC_MOTIFS || '').split(/[;,]/)
    .map(function (s) { return norm_(s).trim(); }).filter(Boolean);
  return new RegExp(REDUC_A_VERIFIER.concat(sup.map(escapeRegex_)).join('|'));
}

function reducEleve_(e) {
  var r = norm_(e.reductions) + ' ' + norm_(e.notes) + ' ' + norm_(e.commentaire);
  if (/offert|gratuit|invit/.test(r)) return { rien: true, libelle: 'offert' };
  var m = r.match(motifsReduc_());
  return m ? { aVerifier: true, libelle: String(e.reductions || e.notes || '').trim() } : {};
}

/** Montant à réclamer, ou null quand il demande une vérification humaine */
function montantDu_(e) {
  var red = reducEleve_(e);
  if (red.rien || red.aVerifier) return null;
  var total = tarifsEleve_(e).reduce(function (a, x) { return a + x.prix; }, 0);
  if (!total) return null;
  var recu = num_(e.montantRecu);
  return { total: total, recu: recu, reste: Math.max(total - recu, 0) };
}

/**
 * Relance de PAIEMENT. Elle ne parle que du règlement : pas de demande de complément
 * d'informations, qui est un autre sujet et un autre message.
 * eleves = un ou plusieurs enfants de la même famille, pour n'envoyer qu'un mail par foyer.
 */
function corpsRelance_(eleves) {
  var c = getConfig_();
  var liste = [].concat(eleves);
  var e = liste[0];
  var bonjour = 'Bonjour' + (e.responsable ? ' ' + String(e.responsable).trim() : '') + ',';
  var lignes = [], total = 0, sur = true, modes = {};

  liste.forEach(function (x) {
    var m = montantDu_(x);
    var t = tarifsEleve_(x);
    var quoi = t.length ? t.map(function (y) { return y.label; }).join(' et ') : (x.creneaux || x.stage || '');
    if (m) { total += m.reste; lignes.push('  • ' + x.prenom + ' ' + x.nom + (quoi ? ' — ' + quoi : '') + ' : ' + m.reste + ' €' + (m.recu ? ' (reste à régler, ' + m.recu + ' € déjà reçus)' : '')); }
    else { sur = false; lignes.push('  • ' + x.prenom + ' ' + x.nom + (quoi ? ' — ' + quoi : '')); }
    if (x.modePaiement) modes[norm_(x.modePaiement)] = x.modePaiement;
  });

  var txt = bonjour + '\n\n';
  txt += liste.length > 1
    ? 'Nous revenons vers vous au sujet des inscriptions de vos enfants à ' + c.NOM_STRUCTURE + '.'
    : 'Nous revenons vers vous au sujet de l\'inscription de ' + e.prenom + ' ' + e.nom + ' à ' + c.NOM_STRUCTURE + '.';
  txt += '\n\nSauf erreur de notre part, le règlement ne nous est pas encore parvenu :\n' + lignes.join('\n');
  if (sur && liste.length > 1) txt += '\n  Total : ' + total + ' €';
  if (!sur) txt += '\n\nUn tarif particulier s\'applique à votre situation : nous vous confirmons le montant exact par retour de mail.';
  txt += '\n\n' + infosPaiement_(Object.keys(modes).length === 1 ? modes[Object.keys(modes)[0]] : '');
  // beaucoup de familles ont déjà convenu d'un échelonnement de vive voix : on le dit,
  // sinon elles reçoivent une relance qui a l'air de ne pas tenir compte de leur accord
  var suivi = String(c.CONTACT_PAIEMENTS || '').trim();
  txt += '\n\nSi le règlement a déjà été effectué' +
    (suivi ? ', ou si vous avez déjà convenu de modalités de paiement avec ' + suivi : '') +
    ', merci de ne pas tenir compte de ce message. Pour toute difficulté ou pour étaler le règlement en plusieurs fois, répondez simplement à ce mail, nous trouverons une solution.';
  txt += signature_();

  var sujet = liste.length > 1
    ? 'Règlement des inscriptions · ' + c.NOM_STRUCTURE
    : 'Règlement de l\'inscription de ' + e.prenom + ' · ' + c.NOM_STRUCTURE;
  return { sujet: sujet, corps: txt, montant: sur ? total : null };
}

// ---------- mise en page officielle des mails (bandeau + couleurs de la structure) ----------

function logoBlob_() {
  var d = logo_(getConfig_().ID_LOGO);
  var m = d && d.match(/^data:([^;]+);base64,(.*)$/);
  if (!m) return null;
  return Utilities.newBlob(Utilities.base64Decode(m[2]), m[1], 'logo.png');
}

/** Texte brut du mail -> HTML aux couleurs de la structure (le texte brut reste la version de secours) */
function mailHtml_(corps, logoSrc) {
  var c = getConfig_();
  var acc = c.COULEUR_ACCENT || '#9b1017';
  var blocs = String(corps || '').replace(/\r/g, '').split(/\n{2,}/);
  var puce = /^\s*[•\-]\s+/;
  var html = blocs.map(function (b) {
    // un bloc = paragraphe ; les lignes à puces consécutives deviennent une vraie liste
    var out = '', para = [], liste = [];
    function flushPara() { if (para.length) { out += '<p style="margin:0 0 14px">' + para.join('<br>') + '</p>'; para = []; } }
    function flushListe() { if (liste.length) { out += '<ul style="margin:0 0 14px;padding-left:20px">' + liste.join('') + '</ul>'; liste = []; } }
    b.split('\n').forEach(function (l, i) {
      if (puce.test(l)) { flushPara(); liste.push('<li style="margin:3px 0">' + esc_(l.replace(puce, '')) + '</li>'); return; }
      flushListe();
      var t = esc_(l);
      // titre de section : une ligne courte tout en majuscules, comme dans les envois groupés
      if (i === 0 && l.length < 70 && l.trim() && l === l.toUpperCase() && /[A-ZÀ-Ý]/.test(l)) {
        flushPara();
        out += '<div style="margin:22px 0 10px;color:' + acc + ';font-size:13px;font-weight:bold;letter-spacing:.6px">' + t + '</div>';
        return;
      }
      if (i === 0 && /:\s*$/.test(l) && l.length < 60) t = '<b style="color:' + acc + '">' + t + '</b>';
      para.push(t);
    });
    flushPara(); flushListe();
    return out;
  }).join('');
  // Bandeau : titre en serif, baseline et intertitres dans la couleur d'accent,
  // filets dans la couleur secondaire, corps en sans-serif. AUCUNE image : pas de
  // pièce jointe, pas de « télécharger les images » chez le destinataire, et le
  // même rendu dans Gmail, Outlook et sur mobile.
  //
  // MOT_ACCENTUE permet de teinter un mot du nom (par exemple le mot central d'un
  // nom composé) sans toucher au code : il est cherché tel quel dans NOM_STRUCTURE.
  var nom = String(c.NOM_STRUCTURE || '');
  var accent = acc;
  var mot = String(c.MOT_ACCENTUE || '').trim();
  var titre;
  if (mot && nom.indexOf(mot) >= 0) {
    var i = nom.indexOf(mot);
    titre = span_(nom.slice(0, i), '#141414') + span_(mot, accent) + span_(nom.slice(i + mot.length), '#141414');
  } else {
    titre = span_(nom, '#141414');
  }
  var site = String(c.SITE_WEB || '').replace(/^https?:\/\//, '').replace(/\/$/, '');
  var pied = [c.NOM_STRUCTURE, c.NOM_ASSO, String(c.ADRESSE_STRUCTURE || '').trim()].filter(String).map(esc_).join(' · ');
  return '<div data-erp="1" style="font-family:Calibri,Arial,sans-serif;font-size:15px;color:#2b2b2b;line-height:1.55;max-width:640px">' +
    '<div style="text-align:center;padding:6px 0 4px 0">' + titre +
    (c.BASELINE ? '<div style="font-size:13px;font-weight:bold;color:' + accent + ';padding-top:4px">' + esc_(c.BASELINE) + '</div>' : '') +
    '</div>' +
    filet_(c, '8px 0 18px 0') +
    html +
    filet_(c, '18px 0 0 0') +
    '<p style="text-align:center;color:#6b6b6b;font-size:12px;margin:8px 0 0 0">' + pied +
    (site ? ' · <a href="https://' + esc_(site) + '" style="color:' + accent + '">' + esc_(site) + '</a>' : '') +
    '</p></div>';
}

var MAIL_TITRE = 'font-family:Georgia,serif;font-size:30px;font-weight:bold';

function span_(txt, couleur) {
  return txt ? '<span style="' + MAIL_TITRE + ';color:' + couleur + '">' + esc_(txt) + '</span>' : '';
}

function filet_(c, marge) {
  return '<table cellpadding="0" cellspacing="0" border="0" width="100%" style="border-collapse:collapse;margin:' + marge + '">' +
    '<tr><td style="background-color:' + (c.COULEUR_FILET || '#ffd400') + ';height:4px;line-height:4px;font-size:0">&nbsp;</td></tr></table>';
}

/**
 * Le gabarit n'a plus d'image : plus de pièce jointe logo, plus de « télécharger les images »
 * chez le destinataire, et le même rendu partout (Gmail, Outlook, mobile).
 */
function optionsMail_(corps) {
  return { name: getConfig_().NOM_STRUCTURE, htmlBody: mailHtml_(corps) };
}

/** Aperçu dans l'ERP : exactement le HTML qui part */
function htmlApercu_(corps) {
  return mailHtml_(corps);
}

function verifDest_(to) {
  var t = String(to || '').trim();
  if (!t) throw new Error('Pas d\'adresse email.');
  t.split(/\s*,\s*/).forEach(function (a) {
    if (!/^[^\s@<>"]+@[^\s@<>"]+\.[^\s@<>"]+$/.test(a)) throw new Error('Adresse email invalide : ' + a);
  });
  return t;
}

function creerBrouillon_(to, sujet, corps) {
  if (!to) throw new Error('Pas d\'adresse email pour cette inscription.');
  var d = GmailApp.createDraft(verifDest_(to), sujet, corps, optionsMail_(corps));
  return d.getId();
}

function noterEnvoi_(e, type) {
  var s = suiviIndex_()[e.id] || { ID: e.id, Source: e.source, Nom: e.nom, Prenom: e.prenom, Statut: 'Actif', Paye: 'Non' };
  delete s._row;
  if (type === 'relance') s.DerniereRelance = new Date(); else s.Confirmation = new Date();
  saveSuivi_(s);
}

/**
 * Un mail de relance peut couvrir plusieurs enfants d'un même foyer : la date de relance
 * doit alors être notée sur chacun, sinon les frères et sœurs ressortent dans la liste.
 */
function noterEnvoiTous_(p, e, type) {
  var l = (p && p.items || []).filter(function (x) { return x && x.id; });
  if (!l.length) return noterEnvoi_(e, type);
  l.forEach(function (x) {
    try {
      var y = (x.source === e.source && String(x.id) === String(e.id)) ? e : apiEleve_({ source: x.source, id: x.id });
      noterEnvoi_(y, type);
    } catch (err) { }
  });
}

/** Bouton "Créer le brouillon" (texte éventuellement retouché dans l'ERP) */
function apiBrouillon_(p) {
  var e = apiEleve_({ source: p.source, id: p.id });
  var m = p.type === 'relance' ? corpsRelance_(e) : corpsConfirmation_(e);
  var to = p.to || e.emailBest, sujet = p.sujet || m.sujet, corps = p.corps || m.corps;
  creerBrouillon_(to, sujet, corps);
  noterEnvoiTous_(p, e, p.type);
  logAction_('brouillon ' + p.type, e.prenom + ' ' + e.nom);
  return { to: to, sujet: sujet };
}

/** Bouton "Envoyer maintenant" (après aperçu dans l'ERP) */
function apiEnvoyerMail_(p) {
  var e = apiEleve_({ source: p.source, id: p.id });
  var m = p.type === 'relance' ? corpsRelance_(e) : corpsConfirmation_(e);
  var to = verifDest_(p.to || e.emailBest), sujet = p.sujet || m.sujet, corps = p.corps || m.corps;
  GmailApp.sendEmail(to, sujet, corps, optionsMail_(corps));
  noterEnvoiTous_(p, e, p.type);
  logAction_('mail envoyé ' + p.type, e.prenom + ' ' + e.nom + ' -> ' + to);
  return { to: to, sujet: sujet };
}

/** Aperçu du mail avant création / envoi */
function apiApercuMail_(p) {
  var e = apiEleve_({ source: p.source, id: p.id });
  var m = p.type === 'relance' ? corpsRelance_(e) : corpsConfirmation_(e);
  return { to: e.emailBest, sujet: m.sujet, corps: m.corps, html: htmlApercu_(m.corps) };
}

// ---------- brouillons Gmail gérés depuis l'ERP ----------

/**
 * Les brouillons préparés par Claude ont une étiquette dans l'objet ([À COMPLÉTER], [À RELIRE], [PRÊT])
 * pour dire s'il faut encore les retoucher. Elle est retirée avant l'envoi : le destinataire ne la voit jamais.
 */
var RE_ETIQUETTE = /^\s*\[\s*(a\s*completer|à\s*compléter|a\s*relire|à\s*relire|pret|prêt)[^\]]*\]\s*/i;
function etiquette_(sujet) {
  var m = String(sujet || '').match(RE_ETIQUETTE);
  return m ? m[0].trim() : '';
}
function sansEtiquette_(sujet) {
  return String(sujet || '').replace(RE_ETIQUETTE, '').trim();
}

function brouillon_(id) {
  var d = GmailApp.getDraft(String(id || ''));
  if (!d) throw new Error('Brouillon introuvable (déjà envoyé ou supprimé ?).');
  return d;
}

function apiBrouillonVoir_(p) {
  var d = brouillon_(p.id), m = d.getMessage();
  var sujet = m.getSubject() || '';
  var html = m.getBody() || '';
  var logo = '';
  try { logo = logo_(getConfig_().ID_LOGO); } catch (e) { logo = ''; }
  if (logo) html = html.replace(/(src=["'])cid:[^"']+(["'])/gi, '$1' + logo + '$2');
  var pj = 0;
  try { pj = m.getAttachments({ includeInlineImages: false }).length; } catch (e) { pj = 0; }
  return { id: p.id, a: m.getTo(), cc: m.getCc(), sujet: sansEtiquette_(sujet), etiquette: etiquette_(sujet), corps: m.getPlainBody(), html: html, piecesJointes: pj, date: fmtDateTime_(m.getDate()) };
}

function apiBrouillonEnvoyer_(p) {
  var d = brouillon_(p.id), m = d.getMessage();
  if (!String(m.getTo() || '').trim()) throw new Error('Ce brouillon n\'a pas de destinataire.');
  // l'étiquette de relecture ne part jamais chez le destinataire
  if (etiquette_(m.getSubject())) {
    var pj = [];
    try { pj = m.getAttachments({ includeInlineImages: false }); } catch (e) { pj = []; }
    var opt = { htmlBody: m.getBody(), name: getConfig_().NOM_STRUCTURE };
    if (pj.length) opt.attachments = pj;
    d.update(m.getTo(), sansEtiquette_(m.getSubject()), m.getPlainBody(), opt);
    d = brouillon_(p.id); m = d.getMessage();
  }
  d.send();
  logAction_('brouillon envoyé', m.getSubject() + ' -> ' + m.getTo());
  return true;
}

function apiBrouillonModifier_(p) {
  var d = brouillon_(p.id);
  var to = verifDest_(p.a);
  var sujet = sansEtiquette_(p.sujet);
  if (!sujet) throw new Error('Objet vide.');
  d.update(to, sujet, p.corps || '', optionsMail_(p.corps || ''));
  logAction_('brouillon modifié', sujet);
  return true;
}

function apiBrouillonSupprimer_(p) {
  var d = brouillon_(p.id), sujet = d.getMessage().getSubject();
  d.deleteDraft();
  logAction_('brouillon supprimé', sujet);
  return true;
}

/** Élèves à relancer (pas payés depuis X jours ou infos manquantes) */
/**
 * Qui est à relancer, et surtout POURQUOI les autres ne le sont pas.
 * Les compteurs se répartissent exactement : total = relancés + chaque motif d'exclusion.
 */
function aRelancerDetail_() {
  var c = getConfig_();
  var jours = num_(c.RELANCE_APRES_JOURS) || 10;
  var joursInter = num_(c.RELANCE_INTERVALLE_JOURS) || 7;
  var apres = jours * 86400000, inter = joursInter * 86400000;
  var now = Date.now();
  var st = { total: 0, desinscrits: 0, sansEmail: 0, payes: 0, offerts: 0, recents: 0, relanceRecente: 0, aRelancer: 0, foyers: 0, aVerifier: 0, jours: jours, joursInter: joursInter };
  var out = [];
  loadEleves_('cours').concat(loadEleves_('stages')).forEach(function (e) {
    st.total++;
    if (e.statut === 'Désinscrit') { st.desinscrits++; return; }
    if (!e.emailBest) { st.sansEmail++; return; }
    // relance de PAIEMENT : qui a payé n'y figure pas, même s'il manque des informations
    // (le complément d'infos est un autre message, à envoyer séparément)
    if (e.paye === 'Oui') { st.payes++; return; }
    if (reducEleve_(e).rien) { st.offerts++; return; }
    if (e.tsRaw && now - e.tsRaw < apres) { st.recents++; return; }
    var last = parseDate_(e.derniereRelance);
    if (last && now - last.getTime() < inter) { st.relanceRecente++; return; }
    out.push(e);
  });
  st.aRelancer = out.length;
  // regroupement par foyer : une famille avec trois enfants reçoit UN mail, pas trois
  var par = {}, ordre = [];
  out.forEach(function (e) {
    var k = norm_(e.emailBest);
    if (!par[k]) { par[k] = []; ordre.push(k); }
    par[k].push(e);
  });
  var foyers = ordre.map(function (k) {
    var l = par[k];
    var m = corpsRelance_(l);
    return { email: l[0].emailBest, eleves: l, sujet: m.sujet, corps: m.corps, montant: m.montant };
  });
  st.foyers = foyers.length;
  st.aVerifier = foyers.filter(function (f) { return f.montant === null; }).length;
  return { eleves: out, foyers: foyers, stats: st };
}

function aRelancer_() { return aRelancerDetail_().eleves; }

function apiARelancer_() {
  var d = aRelancerDetail_();
  // le logo part une fois pour toutes : recopié dans chaque mail il pèserait cent fois son poids,
  // l'ERP le remet dans l'aperçu à la place de {{LOGO}}
  var logoSrc = '';
  try { logoSrc = logo_(getConfig_().ID_LOGO) || ''; } catch (e) { logoSrc = ''; }
  return {
    liste: d.eleves.map(function (e) {
      return { id: e.id, source: e.source, nom: e.nom, prenom: e.prenom, paye: e.paye, manquants: e.manquants, ts: e.ts, email: e.emailBest, derniereRelance: e.derniereRelance };
    }),
    logo: logoSrc,
    foyers: (d.foyers || []).map(function (f) {
      return {
        email: f.email, sujet: f.sujet, montant: f.montant,
        // texte complet et rendu HTML du mail tel qu'il partira : l'aperçu n'a plus rien à demander au serveur
        corps: f.corps, html: mailHtml_(f.corps, '{{LOGO}}'),
        eleves: f.eleves.map(function (e) {
          return { id: e.id, source: e.source, nom: e.nom, prenom: e.prenom, paye: e.paye, montantRecu: e.montantRecu,
            manquants: e.manquants || [],
            reduction: String(e.reductions || '').trim(), commentaire: String(e.commentaire || '').trim(), ts: e.ts };
        })
      };
    }),
    stats: d.stats
  };
}

/**
 * Note une relance faite en dehors de l'ERP (mail écrit à la main, appel téléphonique).
 * Sans ça, la personne reste éternellement dans « À relancer » puisque la date ne s'écrit
 * que quand c'est l'ERP qui a créé le brouillon.
 */
function apiNoterRelance_(p) {
  var n = 0, err = [];
  (p.items || []).slice(0, 80).forEach(function (it) {
    try { noterEnvoi_(apiEleve_({ source: it.source, id: it.id }), 'relance'); n++; }
    catch (e) { err.push(it.id + ' : ' + e.message); }
  });
  logAction_('relances notées à la main', n + ' inscription(s)');
  return { n: n, erreurs: err };
}

/**
 * Crée les brouillons de relance pour une liste d'ids [{source,id}].
 * Les enfants d'une même famille sont regroupés : un seul brouillon par adresse mail.
 */
function apiBrouillonsRelance_(p) {
  var voulus = {};
  (p.items || []).slice(0, 60).forEach(function (it) { voulus[it.source + '|' + it.id] = 1; });
  var d = aRelancerDetail_();
  var n = 0, err = [];
  (d.foyers || []).forEach(function (f) {
    var pris = f.eleves.filter(function (e) { return voulus[e.source + '|' + e.id]; });
    if (!pris.length) return;
    try {
      var m = corpsRelance_(pris);
      creerBrouillon_(f.email, m.sujet, m.corps);
      pris.forEach(function (e) { noterEnvoi_(e, 'relance'); });
      logAction_('brouillon relance', f.email + ' (' + pris.map(function (e) { return e.prenom; }).join(', ') + ')');
      n++;
    } catch (e) { err.push(f.email + ' : ' + e.message); }
  });
  return { crees: n, erreurs: err };
}

// ---------- déclencheurs ----------

function onInscriptionCours(ev) { viaDeclencheurOuProprio_(ev); handleInscription_('cours', ev); }
function onInscriptionStage(ev) { viaDeclencheurOuProprio_(ev); handleInscription_('stages', ev); }

function handleInscription_(source, ev) {
  try {
    var c = getConfig_(true);
    clearMemo_();
    var list = loadEleves_(source); // rattache aussi le suivi d'une réponse modifiée
    var row = ev && ev.range ? ev.range.getRow() : responseSheet_(source).getLastRow();
    var e = list.filter(function (x) { return x.row === row; })[0];
    if (!e) return;
    // une réponse modifiée par le parent redéclenche le trigger : on ne réécrit pas le suivi
    if (!e.hasSuivi) saveSuivi_({ ID: e.id, Source: source, Nom: e.nom, Prenom: e.prenom, Statut: 'Actif', Paye: 'Non' });
    if (norm_(c.CONFIRMATION_AUTO) === 'oui' && e.emailBest && !e.hasSuivi && !e.confirmation) {
      var m = corpsConfirmation_(e);
      creerBrouillon_(e.emailBest, m.sujet, m.corps);
      saveSuivi_({ ID: e.id, Confirmation: new Date() });
    }
    logAction_('nouvelle inscription ' + source, e.prenom + ' ' + e.nom);
  } catch (err) {
    logAction_('ERREUR inscription ' + source, err.message);
  }
}

/** Tous les jours à 7h */
function dailyJob(ev) {
  viaDeclencheurOuProprio_(ev);
  var c = getConfig_(true);
  if (norm_(c.RELANCES_AUTO) === 'oui') {
    var items = aRelancer_().slice(0, 15);
    apiBrouillonsRelance_({ items: items.map(function (e) { return { source: e.source, id: e.id }; }) });
  }
  // fiches de présence : remet les croix en face des bons noms (ajouts / départs / changements de cours)
  try { syncToutesFiches_(); } catch (e) { logAction_('ERREUR fiches', e.message); }
  if (c.DIGEST_EMAILS) {
    try { envoyerRecap_(c.DIGEST_EMAILS); } catch (e) { logAction_('ERREUR récap', e.message); }
  }
}

function esc_(s) {
  return String(s === undefined || s === null ? '' : s).replace(/[&<>"]/g, function (ch) { return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[ch]; });
}

/** Récap quotidien pour l'équipe (pas de données santé détaillées dans le mail) */
function envoyerRecap_(to) {
  var a = apiAccueil_();
  var h = '<div style="font-family:Arial,sans-serif;max-width:640px;color:#151313">';
  h += '<h2 style="color:#C8262B;margin:0 0 4px">' + esc_(getConfig_().NOM_STRUCTURE) + ' · ' + esc_(a.date) + '</h2>';
  h += '<p style="margin:0 0 16px;color:#555">Récap automatique de l\'ERP</p>';
  if (a.coursDuJour.length) {
    h += '<h3>Cours aujourd\'hui</h3><ul>';
    a.coursDuJour.forEach(function (k) {
      h += '<li><b>' + esc_(k.horaire) + '</b> ' + esc_(k.libelle) + ' : ' + k.inscrits + ' élèves' + (k.alertes ? ' · <span style="color:#C8262B">' + k.alertes + ' alerte(s) santé, voir ERP</span>' : '') + '</li>';
    });
    h += '</ul>';
  }
  if (a.rappels.length) {
    h += '<h3>Rappels</h3><ul>';
    a.rappels.forEach(function (r) { h += '<li>' + (r.enRetard ? '<b style="color:#C8262B">En retard</b> ' : '') + esc_(r.Date) + ' ' + esc_(r.Type) + ' : ' + esc_(r.Titre) + '</li>'; });
    h += '</ul>';
  }
  if (a.agenda.length) {
    h += '<h3>Agenda (7 jours)</h3><ul>';
    a.agenda.slice(0, 12).forEach(function (ev) { h += '<li>' + esc_(ev.quand) + ' : ' + esc_(ev.titre) + '</li>'; });
    h += '</ul>';
  }
  h += '<h3>Chiffres</h3><ul><li>' + a.stats.eleves + ' élèves en cours annuels, ' + a.stats.stagiaires + ' inscrits en stage</li>' +
    '<li>' + a.stats.impayes + ' inscription(s) sans paiement noté</li>' +
    '<li>' + a.stats.aRelancer + ' à relancer · ' + a.stats.brouillons + ' brouillon(s) Gmail en attente · ' + a.stats.nonLus + ' mail(s) non lu(s)</li></ul>';
  if (a.nouveaux.length) {
    h += '<h3>Dernières inscriptions</h3><ul>';
    a.nouveaux.forEach(function (e) { h += '<li>' + esc_(e.ts) + ' : ' + esc_(e.prenom + ' ' + e.nom) + ' (' + esc_(e.source) + ')</li>'; });
    h += '</ul>';
  }
  var url = ScriptApp.getService().getUrl();
  if (url) h += '<p><a href="' + url + '" style="background:#C8262B;color:#fff;padding:10px 16px;border-radius:8px;text-decoration:none">Ouvrir l\'ERP</a></p>';
  h += '</div>';
  MailApp.sendEmail({ to: to, subject: 'Récap du jour · ' + a.date, htmlBody: h, name: 'ERP ' + getConfig_().NOM_STRUCTURE });
}

/** Derniers fils Gmail + brouillons en attente */
function apiMails_() {
  // brouillons (et fil de discussion auquel ils répondent : badge "réponse prête")
  var tous = GmailApp.getDrafts();
  var parFil = {};
  var drafts = tous.slice(0, 40).map(function (d) {
    var m = d.getMessage();
    var fil = '';
    try { fil = m.getThread().getId(); } catch (e) { fil = ''; }
    if (fil && !parFil[fil]) parFil[fil] = d.getId();
    var suj = m.getSubject() || '';
    return { id: d.getId(), sujet: sansEtiquette_(suj) || '(sans objet)', etiquette: etiquette_(suj), a: m.getTo(), date: fmtDateTime_(m.getDate()), fil: fil,
      extrait: String(m.getPlainBody() || '').replace(/\s+/g, ' ').slice(0, 120) };
  });
  // boîte de réception : les 5 derniers mails reçus (lus ou non)
  var list = GmailApp.search('in:inbox', 0, 5).map(function (t) {
    var m = t.getMessages();
    var last = m[m.length - 1];
    return { id: t.getId(), sujet: t.getFirstMessageSubject() || '(sans objet)', de: last.getFrom(), date: fmtDateTime_(t.getLastMessageDate()),
      nonLu: t.isUnread(), extrait: last.getPlainBody().replace(/\s+/g, ' ').slice(0, 160), lien: 'https://mail.google.com/mail/#inbox/' + t.getId(),
      brouillon: parFil[t.getId()] || '' };
  });
  return { threads: list, drafts: drafts, nonLus: GmailApp.getInboxUnreadCount() };
}

/** Aperçu d'un mail reçu, dans l'ERP */
function apiMailVoir_(p) {
  var t = GmailApp.getThreadById(String(p.id || ''));
  if (!t) throw new Error('Mail introuvable (déplacé ou supprimé ?).');
  var ms = t.getMessages();
  var m = ms[ms.length - 1];
  var brouillon = '';
  GmailApp.getDrafts().slice(0, 40).forEach(function (d) {
    if (brouillon) return;
    try { if (d.getMessage().getThread().getId() === t.getId()) brouillon = d.getId(); } catch (e) { }
  });
  return {
    id: t.getId(), sujet: t.getFirstMessageSubject() || '(sans objet)', de: m.getFrom(), a: m.getTo(),
    date: fmtDateTime_(m.getDate()), corps: m.getPlainBody(), html: m.getBody() || '', nonLu: t.isUnread(),
    messages: ms.length, brouillon: brouillon, lien: 'https://mail.google.com/mail/#inbox/' + t.getId()
  };
}

/** Boîte de réception : mails cochés -> corbeille ou marqués comme lus. p = { ids: [threadId], action: 'corbeille' | 'lu' } */
function apiMailsAction_(p) {
  var ids = (p.ids || []).slice(0, 50);
  if (!ids.length) throw new Error('Coche au moins un mail.');
  var n = 0;
  ids.forEach(function (id) {
    var t = null;
    try { t = GmailApp.getThreadById(String(id)); } catch (e) { t = null; }
    if (!t) return;
    if (p.action === 'corbeille') t.moveToTrash(); else t.markRead();
    n++;
  });
  logAction_('mails ' + (p.action === 'corbeille' ? 'supprimés' : 'lus'), String(n));
  return { n: n };
}
