/**
 * Rattachement d'une inscription a un creneau, et lecture du statut de paiement.
 *
 * C'est le coeur fragile de l'ERP : une inscription qui ne se rattache a aucun
 * creneau disparait des listes d'appel et des effectifs. Le texte vient d'un
 * Google Form que des humains retouchent en cours de saison ("COMPLET",
 * "2 places restantes"), donc le rattachement doit survivre a ces retouches.
 */
const { charger } = require('./harness');
const { erpDeTest } = require('./fixtures');
const { groupe, test, egal, profond, vrai, faux } = global.T;

const { erp } = charger(erpDeTest());
const COURS = erp.coursList_();
const STAGES = erp.stagesList_();
const libelles = (l) => l.map((c) => c.code || c.id).sort();

groupe('coursList_ et stagesList_ : lecture de la table de reference', () => {
  test('les quatre creneaux sont lus', () => egal(COURS.length, 4));
  test('le tarif est un nombre, pas du texte', () => egal(COURS[0].tarif, 300));
  test('un Max absent retombe sur 12', () => {
    const { erp: e2 } = charger(erpDeTest({
      onglets: { Cours: [['Cle', 'Libelle', 'Jour', 'Horaire', 'Public', 'Tarif', 'Max', 'Code'], ['Lundi 17h00', 'G', 'Lundi', '17h00-18h30', 'E', 300, '', 'C1']] }
    }));
    egal(e2.coursList_()[0].max, 12);
  });
  test('un stage est identifie par sa cle ET son groupe', () => {
    egal(STAGES[0].id, '26 au 30 octobre 2026#matin');
    egal(STAGES[1].id, '26 au 30 octobre 2026#apres-midi');
  });
});

groupe('matchCours_ niveau 1 : le code entre crochets', () => {
  test('le code suffit, meme avec un libelle completement reecrit', () => {
    profond(libelles(erp.matchCours_('PLUS AUCUNE PLACE !!! [C2]', COURS)), ['C2']);
  });
  test('le code gagne sur un libelle trompeur', () => {
    profond(libelles(erp.matchCours_('Groupe 6-9 ans Lundi 17h00 [C4]', COURS)), ['C4']);
  });
  test('plusieurs codes coches rattachent plusieurs creneaux', () => {
    profond(libelles(erp.matchCours_('[C1], [C3]', COURS)), ['C1', 'C3']);
  });
  test('un code inconnu ne fait pas echouer : on retombe sur les niveaux suivants', () => {
    profond(libelles(erp.matchCours_('Lundi 17h00 [C99]', COURS)), ['C1']);
  });
});

groupe('matchCours_ niveau 2 : la cle cherchee dans le texte', () => {
  test('la cle exacte rattache le creneau', () => {
    profond(libelles(erp.matchCours_('Mercredi 11h00 - Groupe 4-6 ans', COURS)), ['C3']);
  });
  test('la casse et les accents n\'ont pas d\'importance', () => {
    profond(libelles(erp.matchCours_('MERCREDI 11H00', COURS)), ['C3']);
  });
  test('un texte decore autour de la cle ne gene pas', () => {
    profond(libelles(erp.matchCours_('>>> Mercredi 11h00 (2 places restantes) <<<', COURS)), ['C3']);
  });
});

groupe('matchCours_ niveau 3 : repli sur le jour et l\'heure', () => {
  test('jour + heure de debut retrouvent le creneau quand la cle a ete reecrite', () => {
    profond(libelles(erp.matchCours_('Le jeudi, de 19h15 a 21h15, adultes', COURS)), ['C4']);
  });
  test('accepte 19:15 comme 19h15', () => {
    profond(libelles(erp.matchCours_('jeudi 19:15', COURS)), ['C4']);
  });
  test('accepte une heure sans minutes', () => {
    profond(libelles(erp.matchCours_('mardi 17h', COURS)), ['C2']);
  });
  test('le bon jour avec la mauvaise heure ne rattache rien', () => {
    profond(erp.matchCours_('jeudi 08h00', COURS), []);
  });
  test('la bonne heure le mauvais jour ne rattache rien', () => {
    profond(erp.matchCours_('dimanche 17h00', COURS), []);
  });
});

groupe('matchCours_ : les cas ou il ne faut rien rattacher', () => {
  test('un texte vide ne rattache rien', () => profond(erp.matchCours_('', COURS), []));
  test('null ne rattache rien et ne leve pas', () => profond(erp.matchCours_(null, COURS), []));
  test('un creneau inconnu ne rattache rien', () => {
    profond(erp.matchCours_('Samedi 09h00 - Trapeze volant', COURS), []);
  });
});

groupe('matchStages_ : un parent peut cocher plusieurs sessions', () => {
  test('une session simple', () => {
    profond(libelles(erp.matchStages_('26 au 30 octobre 2026 - matin - 200 €', STAGES)), ['26 au 30 octobre 2026#matin']);
  });
  test('la cle seule ne suffit pas : il faut le groupe matin ou apres-midi', () => {
    profond(erp.matchStages_('26 au 30 octobre 2026', STAGES), []);
  });
  test('deux sessions separees par "€," donnent deux rattachements', () => {
    const r = libelles(erp.matchStages_('26 au 30 octobre 2026 matin 200 €, 15 au 19 fevrier 2027 matin 200 €', STAGES));
    profond(r, ['15 au 19 fevrier 2027#matin', '26 au 30 octobre 2026#matin']);
  });
  test('deux sessions separees par un retour a la ligne aussi', () => {
    const r = libelles(erp.matchStages_('26 au 30 octobre 2026 matin\n26 au 30 octobre 2026 apres-midi', STAGES));
    profond(r, ['26 au 30 octobre 2026#apres-midi', '26 au 30 octobre 2026#matin']);
  });
  test('matin et apres-midi de la meme semaine ne se confondent pas', () => {
    profond(libelles(erp.matchStages_('26 au 30 octobre 2026 apres-midi 250 €', STAGES)), ['26 au 30 octobre 2026#apres-midi']);
  });
});

groupe('couleur_ : le statut de paiement est lu sur le fond de la cellule', () => {
  test('le vert Google Sheets est reconnu', () => egal(erp.couleur_('#6aa84f'), 'vert'));
  test('les verts plus clairs aussi', () => {
    egal(erp.couleur_('#93c47d'), 'vert');
    egal(erp.couleur_('#b7e1cd'), 'vert');
  });
  test('le rouge est reconnu', () => egal(erp.couleur_('#e06666'), 'rouge'));
  test('le jaune est reconnu', () => egal(erp.couleur_('#ffe599'), 'jaune'));
  test('le blanc ne vaut aucun statut', () => egal(erp.couleur_('#ffffff'), ''));
  test('les bandes alternees grises ne valent aucun statut', () => egal(erp.couleur_('#f8f9fa'), ''));
  test('une valeur absente ne leve pas', () => {
    egal(erp.couleur_(''), '');
    egal(erp.couleur_(null), '');
    egal(erp.couleur_('pas une couleur'), '');
  });
  test('la casse hexadecimale n\'a pas d\'importance', () => egal(erp.couleur_('#6AA84F'), 'vert'));
  test('le diese est facultatif', () => egal(erp.couleur_('6aa84f'), 'vert'));
});

groupe('makeId_ : identifier une inscription de facon stable', () => {
  test('meme horodateur et meme nom donnent le meme identifiant', () => {
    const d = new Date(2026, 8, 7, 10, 0, 0);
    egal(erp.makeId_(d, 'Martin', 'Lea'), erp.makeId_(d, 'Martin', 'Lea'));
  });
  test('la casse et les accents ne changent pas l\'identifiant', () => {
    const d = new Date(2026, 8, 7, 10, 0, 0);
    egal(erp.makeId_(d, 'MARTIN', 'Léa'), erp.makeId_(d, 'martin', 'lea'));
  });
  test('deux enfants de la meme fratrie inscrits ensemble ont des identifiants distincts', () => {
    const d = new Date(2026, 8, 7, 10, 0, 0);
    vrai(erp.makeId_(d, 'Martin', 'Lea') !== erp.makeId_(d, 'Martin', 'Noe'));
  });
});

groupe('heureDebut_ et heuresTexte_', () => {
  test('extrait l\'heure de debut d\'une plage', () => egal(erp.heureDebut_('17h00-18h30'), '17h00'));
  test('complete une heure a un chiffre', () => egal(erp.heureDebut_('9h00-10h00'), '09h00'));
  test('complete les minutes manquantes', () => egal(erp.heureDebut_('17h-18h'), '17h00'));
  test('accepte le deux-points', () => egal(erp.heureDebut_('17:00-18:30'), '17h00'));
  test('un horaire illisible renvoie une chaine vide', () => egal(erp.heureDebut_('en soiree'), ''));
  test('liste toutes les heures d\'un texte', () => {
    profond(erp.heuresTexte_('de 17h00 a 18h30'), ['17h00', '18h30']);
  });
});

groupe('libelleCodeForm_ : le libelle a coller dans le Google Form', () => {
  test('porte le code entre crochets, qui est ce qui rend le rattachement robuste', () => {
    const l = erp.libelleCodeForm_(COURS[0]);
    vrai(l.indexOf('[C1]') >= 0, 'le code doit figurer : ' + l);
  });
  test('affiche le nombre de places restantes au singulier', () => {
    vrai(erp.libelleCodeForm_(COURS[0], 1).indexOf('1 place restante') >= 0);
  });
  test('et au pluriel', () => {
    vrai(erp.libelleCodeForm_(COURS[0], 3).indexOf('3 places restantes') >= 0);
  });
  test('affiche complet a zero place', () => {
    vrai(erp.libelleCodeForm_(COURS[0], 0).indexOf('complet') >= 0);
  });
  test('le libelle genere se rattache a son propre creneau (boucle fermee)', () => {
    COURS.forEach((c) => {
      const l = erp.libelleCodeForm_(c, 2);
      profond(erp.matchCours_(l, COURS).map((x) => x.code), [c.code], 'aller-retour pour ' + c.code);
    });
  });
});
