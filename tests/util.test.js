/** Fonctions communes : normalisation, dates, nombres, ecriture sure dans un Sheet. */
const { charger } = require('./harness');
const { erpDeTest } = require('./fixtures');
const { groupe, test, egal, profond, vrai, faux } = global.T;

const { erp } = charger(erpDeTest());

groupe('norm_ : comparer deux textes saisis par des humains', () => {
  test('supprime les accents et la casse', () => {
    egal(erp.norm_('Éveil Débutants'), 'eveil debutants');
  });
  test('unifie les trois apostrophes que produisent Word, Google Forms et un clavier', () => {
    egal(erp.norm_("L’atelier"), "l'atelier");
    egal(erp.norm_("L‘atelier"), "l'atelier");
    egal(erp.norm_('L`atelier'), "l'atelier");
  });
  test('unifie les tirets longs, que les parents copient depuis un PDF', () => {
    egal(erp.norm_('10h – 12h'), '10h - 12h');
    egal(erp.norm_('10h — 12h'), '10h - 12h');
  });
  test('compacte les espaces, insecables comprises', () => {
    egal(erp.norm_('  Lundi  17h   00  '), 'lundi 17h 00');
  });
  test('null et undefined donnent une chaine vide, pas "null"', () => {
    egal(erp.norm_(null), '');
    egal(erp.norm_(undefined), '');
  });
});

groupe('num_ : lire un montant tape a la main', () => {
  test('accepte la virgule decimale francaise', () => egal(erp.num_('12,50'), 12.5));
  test('accepte le symbole euro et les espaces', () => egal(erp.num_('1 250,00 €'), 1250));
  test('accepte l\'espace insecable des exports Sheets', () => egal(erp.num_('1 250'), 1250));
  test('lit les parentheses comptables comme un negatif', () => egal(erp.num_('(300)'), -300));
  test('une cellule vide vaut zero, pas NaN', () => egal(erp.num_(''), 0));
  test('un texte illisible vaut zero, pas NaN', () => egal(erp.num_('a payer'), 0));
  test('un nombre reste intact', () => egal(erp.num_(42.5), 42.5));
});

groupe('parseDate_ : trois formats de date dans le meme classeur', () => {
  test('format francais jj/mm/aaaa', () => {
    const d = erp.parseDate_('12/04/2017');
    profond([d.getFullYear(), d.getMonth() + 1, d.getDate()], [2017, 4, 12]);
  });
  test('format ISO aaaa-mm-jj', () => {
    const d = erp.parseDate_('2026-09-07');
    profond([d.getFullYear(), d.getMonth() + 1, d.getDate()], [2026, 9, 7]);
  });
  test('un objet Date passe tel quel', () => {
    const o = new Date(2026, 0, 15);
    egal(erp.parseDate_(o), o);
  });
  test('une saisie illisible donne null plutot qu\'une date fausse', () => {
    egal(erp.parseDate_('la semaine prochaine'), null);
    egal(erp.parseDate_(''), null);
  });
});

groupe('ageAt_ : age a une date de reference', () => {
  test('anniversaire deja passe dans l\'annee', () => {
    egal(erp.ageAt_('12/04/2017', new Date(2026, 8, 1)), 9);
  });
  test('anniversaire pas encore passe', () => {
    egal(erp.ageAt_('12/12/2017', new Date(2026, 8, 1)), 8);
  });
  test('le jour meme de l\'anniversaire, l\'age a augmente', () => {
    egal(erp.ageAt_('01/09/2017', new Date(2026, 8, 1)), 9);
  });
  test('une naissance illisible ne renvoie pas un age absurde', () => {
    egal(erp.ageAt_('', new Date(2026, 8, 1)), '');
    egal(erp.ageAt_('inconnue', new Date(2026, 8, 1)), '');
  });
});

groupe('safeCell_ : ecrire dans un Sheet sans que Sheets reinterprete', () => {
  test('un numero de telephone garde son zero initial', () => {
    egal(erp.safeCell_('0612345678'), "'0612345678");
  });
  test('un numero de facture ne devient pas une date', () => {
    egal(erp.safeCell_('2026-001'), "'2026-001");
  });
  test('un texte commencant par = n\'est pas execute comme formule', () => {
    egal(erp.safeCell_('=1+1'), "'=1+1");
  });
  test('un texte commencant par + ou @ non plus (injection CSV)', () => {
    egal(erp.safeCell_('+1'), "'+1");
    egal(erp.safeCell_('@SUM(A1)'), "'@SUM(A1)");
  });
  test('un texte normal n\'est pas modifie', () => {
    egal(erp.safeCell_('Martin'), 'Martin');
  });
  test('un nombre reste un nombre', () => {
    egal(erp.safeCell_(300), 300);
  });
});

groupe('plain_ : rien de non serialisable ne part vers le navigateur', () => {
  test('une Date devient du texte jj/mm/aaaa', () => {
    egal(erp.plain_(new Date(2026, 8, 7)), '07/09/2026');
  });
  test('null devient une chaine vide', () => egal(erp.plain_(null), ''));
  test('un nombre reste un nombre', () => egal(erp.plain_(12), 12));
});

groupe('isDate_ et escapeRegex_', () => {
  test('une Date valide est reconnue', () => vrai(erp.isDate_(new Date())));
  test('une Date invalide est rejetee', () => faux(erp.isDate_(new Date('n\'importe quoi'))));
  test('une chaine n\'est pas une Date', () => faux(erp.isDate_('12/04/2017')));
  test('les caracteres speciaux d\'une regex sont echappes', () => {
    const r = new RegExp(erp.escapeRegex_('Saint-Jean (01)'));
    vrai(r.test('Saint-Jean (01)'));
  });
});

groupe('uid_ : identifiants uniques', () => {
  test('format court et prefixe', () => {
    const id = erp.uid_();
    egal(id[0], 'R');
    egal(id.length, 10);
  });
  test('deux appels ne donnent pas le meme identifiant', () => {
    vrai(erp.uid_() !== erp.uid_());
  });
});
