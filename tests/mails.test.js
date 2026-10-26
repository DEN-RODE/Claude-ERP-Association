/**
 * Mails sortants. La regle centrale de ce module, et la raison pour laquelle il
 * est le plus teste : AUCUN montant incertain ne part vers une famille. Un tarif
 * fratrie, une convention avec une commune, un tarif a la carte, et le prix du
 * creneau n'est plus le montant du. Dans ce cas le mail ne chiffre rien et
 * annonce qu'un humain confirmera.
 */
const { charger } = require('./harness');
const { erpDeTest, inscrit } = require('./fixtures');
const { groupe, test, egal, profond, vrai, faux, contient, absent, leve } = global.T;

const { erp, etat } = charger(erpDeTest({ config: { CONTACT_PAIEMENTS: 'Alex', REDUC_MOTIFS: 'Villeneuve' } }));

groupe('montantDu_ : chiffrer, ou refuser de chiffrer', () => {
  test('un cours sans reduction donne le tarif du creneau', () => {
    profond(erp.montantDu_(inscrit()), { total: 300, recu: 0, reste: 300 });
  });
  test('un acompte deja recu est deduit', () => {
    profond(erp.montantDu_(inscrit({ montantRecu: 100 })), { total: 300, recu: 100, reste: 200 });
  });
  test('un trop-percu ne donne jamais un reste negatif', () => {
    profond(erp.montantDu_(inscrit({ montantRecu: 400 })), { total: 300, recu: 400, reste: 0 });
  });
  test('deux creneaux s\'additionnent', () => {
    const e = inscrit({ creneaux: '[C1], [C4]' });
    egal(erp.montantDu_(e).total, 650);
  });

  test('REDUCTION FRATRIE : aucun montant', () => {
    egal(erp.montantDu_(inscrit({ reductions: 'Fratrie' })), null);
  });
  test('TARIF A LA CARTE : aucun montant', () => {
    egal(erp.montantDu_(inscrit({ reductions: 'a la carte' })), null);
  });
  test('DEMI-ANNEE : aucun montant', () => {
    egal(erp.montantDu_(inscrit({ notes: 'demi-annee' })), null);
  });
  test('CONVENTION : aucun montant', () => {
    egal(erp.montantDu_(inscrit({ notes: 'convention avec la mairie' })), null);
  });
  test('QUOTIENT FAMILIAL : aucun montant', () => {
    egal(erp.montantDu_(inscrit({ reductions: 'quotient familial 2' })), null);
  });
  test('motif propre a la structure (cle REDUC_MOTIFS) : aucun montant', () => {
    egal(erp.montantDu_(inscrit({ notes: 'habite Villeneuve, tarif conventionne' })), null);
  });
  test('INSCRIPTION OFFERTE : aucun montant', () => {
    egal(erp.montantDu_(inscrit({ reductions: 'offert' })), null);
  });
  test('le motif est detecte sans tenir compte de la casse ni des accents', () => {
    egal(erp.montantDu_(inscrit({ reductions: 'FRATRIE' })), null);
    egal(erp.montantDu_(inscrit({ reductions: 'Demi-année' })), null);
  });
  test('le motif est cherche dans reductions, notes ET commentaire', () => {
    egal(erp.montantDu_(inscrit({ commentaire: 'on avait parle d\'un tarif fratrie' })), null);
  });
  test('un creneau non rattache ne donne pas un montant de zero euro', () => {
    egal(erp.montantDu_(inscrit({ creneaux: 'Samedi 9h - inconnu' })), null);
  });
});

groupe('reducEleve_ : distinguer "offert" de "a verifier"', () => {
  test('offert est un cas ferme, pas une verification', () => {
    const r = erp.reducEleve_(inscrit({ reductions: 'offert' }));
    vrai(r.rien, 'doit etre marque rien a reclamer');
    faux(r.aVerifier);
  });
  test('gratuit et invitation comptent comme offert', () => {
    vrai(erp.reducEleve_(inscrit({ notes: 'gratuit' })).rien);
    vrai(erp.reducEleve_(inscrit({ notes: 'invitation' })).rien);
  });
  test('fratrie demande une verification humaine', () => {
    const r = erp.reducEleve_(inscrit({ reductions: 'fratrie -20%' }));
    vrai(r.aVerifier);
    egal(r.libelle, 'fratrie -20%');
  });
  test('sans reduction, ni l\'un ni l\'autre', () => {
    profond(erp.reducEleve_(inscrit()), {});
  });
});

groupe('corpsRelance_ : ce qui part dans la boite d\'une famille', () => {
  test('un enfant, montant sur : le chiffre figure', () => {
    const m = erp.corpsRelance_(inscrit());
    contient(m.corps, '300 €');
    egal(m.montant, 300);
  });
  test('un enfant, montant incertain : AUCUN chiffre, et la phrase de prudence', () => {
    const m = erp.corpsRelance_(inscrit({ reductions: 'fratrie' }));
    egal(m.montant, null, 'le montant remonte doit etre null');
    contient(m.corps, 'nous vous confirmons le montant exact');
    faux(/\d+\s*€/.test(m.corps.split('Règlement :')[0]), 'aucun montant chiffre avant le bloc de reglement');
  });
  test('une fratrie reçoit UN seul mail, avec une ligne par enfant', () => {
    const m = erp.corpsRelance_([
      inscrit({ prenom: 'Lea' }),
      inscrit({ prenom: 'Noe', creneaux: '[C2]' })
    ]);
    contient(m.corps, 'Lea');
    contient(m.corps, 'Noe');
    contient(m.corps, 'inscriptions de vos enfants');
    egal(m.montant, 650, '300 + 350');
    contient(m.corps, 'Total : 650 €');
  });
  test('un seul enfant incertain dans la fratrie annule le total de toute la fratrie', () => {
    const m = erp.corpsRelance_([
      inscrit({ prenom: 'Lea' }),
      inscrit({ prenom: 'Noe', creneaux: '[C2]', reductions: 'fratrie' })
    ]);
    egal(m.montant, null);
    absent(m.corps, 'Total :');
  });
  test('l\'acompte deja recu est dit, pour ne pas vexer une famille qui a paye', () => {
    const m = erp.corpsRelance_(inscrit({ montantRecu: 100 }));
    contient(m.corps, '200 €');
    contient(m.corps, '100 € déjà reçus');
  });
  test('la relance rappelle qui suit les arrangements de paiement', () => {
    contient(erp.corpsRelance_(inscrit()).corps, 'Alex');
  });
  test('elle invite toujours a repondre en cas de difficulte', () => {
    contient(erp.corpsRelance_(inscrit()).corps, 'nous trouverons une solution');
  });
  test('le sujet nomme la structure et l\'enfant', () => {
    contient(erp.corpsRelance_(inscrit()).sujet, 'Lea');
    contient(erp.corpsRelance_(inscrit()).sujet, 'Association Test');
  });
  test('le sujet est au pluriel pour une fratrie', () => {
    contient(erp.corpsRelance_([inscrit(), inscrit({ prenom: 'Noe' })]).sujet, 'Règlement des inscriptions');
  });
  test('le mail ne contient jamais une information de sante', () => {
    const m = erp.corpsRelance_(inscrit({ sante: 'asthme, ventoline dans le sac' }));
    absent(m.corps, 'asthme');
    absent(m.corps, 'ventoline');
  });
  test('la signature porte les coordonnees de la structure', () => {
    const m = erp.corpsRelance_(inscrit());
    contient(m.corps, 'Association Test');
    contient(m.corps, '01 02 03 04 05');
    contient(m.corps, 'contact@example.org');
  });
});

groupe('corpsConfirmation_ : accuse de reception d\'une inscription', () => {
  test('recapitule le creneau et son tarif', () => {
    const m = erp.corpsConfirmation_(inscrit());
    contient(m.corps, 'Groupe 6-9 ans');
    contient(m.corps, '300 €');
  });
  test('nomme le responsable legal dans le bonjour', () => {
    contient(erp.corpsConfirmation_(inscrit()).corps, 'Bonjour Martin Claude,');
  });
  test('sans responsable, le bonjour reste correct', () => {
    contient(erp.corpsConfirmation_(inscrit({ responsable: '' })).corps, 'Bonjour,');
  });
  test('pour un cours, dit que l\'adhesion est comprise', () => {
    contient(erp.corpsConfirmation_(inscrit()).corps, 'comprend l\'adhésion');
  });
  test('pour un stage, dit que l\'adhesion s\'ajoute, avec son montant', () => {
    const e = inscrit({ source: 'stages', creneaux: '', stage: '26 au 30 octobre 2026 matin 200 €' });
    const m = erp.corpsConfirmation_(e);
    contient(m.corps, 's\'ajoute au prix du stage');
    contient(m.corps, '10 €');
  });
  test('liste les informations manquantes au lieu de les inventer', () => {
    const m = erp.corpsConfirmation_(inscrit({ manquants: ['date de naissance', 'personne a contacter en cas d\'urgence'] }));
    contient(m.corps, 'Il nous manque encore');
    contient(m.corps, 'date de naissance');
  });
  test('indique l\'adresse du rendez-vous', () => {
    contient(erp.corpsConfirmation_(inscrit()).corps, '1 rue des Tests');
  });
});

groupe('infosPaiement_ : coordonnees bancaires', () => {
  test('sans mode precise, les trois moyens sont proposes', () => {
    const t = erp.infosPaiement_('');
    contient(t, 'virement');
    contient(t, 'chèque');
    contient(t, 'espèces');
  });
  test('un mode precise ne montre que celui-la', () => {
    const t = erp.infosPaiement_('Virement');
    contient(t, 'IBAN');
    absent(t, 'chèque');
  });
  test('l\'IBAN et le BIC viennent de la config', () => {
    contient(erp.infosPaiement_('virement'), 'FR00 0000 0000 0000 0000 0000 000');
    contient(erp.infosPaiement_('virement'), 'TESTFRPP');
  });
  test('le cheque est a l\'ordre du titulaire configure', () => {
    contient(erp.infosPaiement_('cheque'), 'ASSOCIATION TEST');
  });
  test('le vocabulaire configure apparait dans la reference a indiquer', () => {
    contient(erp.infosPaiement_('virement'), "le nom de l'adherent en référence");
  });
});

groupe('mailHtml_ : le gabarit HTML du bandeau', () => {
  const html = erp.mailHtml_('TITRE DE SECTION\n\nBonjour,\n\nUn paragraphe.\n\n• premier point\n• second point');
  test('aucune image : pas de piece jointe, pas de "telecharger les images"', () => {
    absent(html, '<img');
    absent(html, 'cid:');
  });
  test('les puces deviennent une vraie liste HTML', () => {
    contient(html, '<ul');
    contient(html, '<li');
  });
  test('une ligne courte en majuscules devient un intertitre', () => {
    contient(html, 'TITRE DE SECTION</div>');
  });
  test('le nom de la structure est dans le bandeau', () => contient(html, 'Association Test'));
  test('la couleur d\'accent vient de la config', () => contient(html, '#9b1017'));
  test('le HTML est echappe : une saisie ne peut pas injecter de balise', () => {
    const h = erp.mailHtml_('Bonjour <script>alert(1)</script>');
    absent(h, '<script>alert');
    contient(h, '&lt;script&gt;');
  });
  test('la largeur est bornee pour rester lisible sur mobile', () => contient(html, 'max-width:640px'));
  test('un corps vide ne leve pas', () => { erp.mailHtml_(''); erp.mailHtml_(null); });
});

groupe('verifDest_ : garde-fou sur le destinataire', () => {
  test('une adresse valide passe', () => { erp.verifDest_('famille@example.org'); });
  test('une adresse vide est refusee', () => leve(() => erp.verifDest_(''), 'mail'));
  test('un texte qui n\'est pas une adresse est refuse', () => leve(() => erp.verifDest_('je ne sais pas')));
});

groupe('tout passe par un brouillon, jamais par un envoi direct', () => {
  test('creerBrouillon_ cree un brouillon et n\'envoie rien', () => {
    const avant = etat.envoyes.length;
    erp.creerBrouillon_('famille@example.org', 'Sujet', 'Corps');
    egal(etat.brouillons.length > 0, true, 'un brouillon doit exister');
    egal(etat.envoyes.length, avant, 'rien ne doit avoir ete envoye');
  });
  test('le brouillon porte bien le HTML du gabarit', () => {
    erp.creerBrouillon_('famille@example.org', 'Sujet', 'Corps');
    const d = etat.brouillons[etat.brouillons.length - 1];
    contient(d.opts.htmlBody, 'Association Test');
  });
});
