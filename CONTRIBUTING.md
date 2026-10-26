# Contribuer

## La règle non négociable

**Aucune donnée d'une structure réelle dans le code.** Pas de nom, pas d'adresse mail, pas de téléphone, pas d'IBAN, pas de SIRET, pas d'identifiant de Google Sheet ou de Form, même dans un commentaire, même « pour tester ».

Ce dépôt vient d'un ERP en production qui gère des données concernant des mineurs. Les identifiants de classeurs sont des adresses vers ces données : les publier est une fuite, même si les permissions Google tiennent.

`tests/aucune-donnee.test.js` vérifie cela automatiquement. Il échoue si vous laissez passer quelque chose. Lancez-le avant de proposer une modification :

```bash
node build.js && node tests/run.js aucune-donnee
```

Toute valeur propre à une structure va dans l'onglet `Config` du classeur, jamais dans `DEFAULT_CONFIG`.

## Mise en place

```bash
git clone https://github.com/DEN-RODE/Claude-ERP-Association.git
cd Claude-ERP-Association
node build.js        # construit bundle/Code.gs
node tests/run.js    # 251 tests
```

Aucune dépendance à installer. Node 18 ou plus.

## Cycle de travail

1. On modifie **`src/`**, jamais `bundle/` — il est regénéré à chaque build.
2. `node build.js` puis `node tests/run.js`.
3. Un nouveau module `.gs` doit être déclaré dans le tableau `ORDRE` de `build.js`, sinon le build refuse de passer. L'ordre compte : Apps Script n'a pas de modules, les utilitaires viennent avant leurs appelants et `Api.gs` reste en dernier.
4. Le bundle est versionné : commitez-le avec vos modifications de `src/`, pour que quelqu'un sans Node puisse installer l'ERP.

## Conventions de code

- **ES5 dans les `.gs`.** `var`, `function`, pas de template literals, pas de fléchées. Le runtime V8 d'Apps Script accepte ES6, mais rester en ES5 garde la compatibilité avec les projets configurés sur l'ancien runtime, et le gain de confort serait mince.
- **Suffixe `_` pour le privé.** `norm_()`, `getConfig_()` : une fonction suffixée n'est pas exposée comme route et n'est pas appelable depuis le navigateur. Les routes publiques sont déclarées dans `routes_()` de `Api.gs`.
- **Commentaires qui expliquent pourquoi.** Le code dit déjà ce qu'il fait. Un commentaire utile explique une contrainte Apps Script, un usage réel, ou pourquoi une solution évidente a été écartée.
- **Pas de données en dur.** Toute valeur de gestion passe par `getConfig_()`, tout mot affiché par `L_()`.
- **Texte français sans caractères exotiques** dans le code (les accents sont normaux, les tirets cadratins et les apostrophes typographiques posent des problèmes dans les chaînes).

## Écrire un test

Le lanceur est maison, l'API tient en dix fonctions. Pas de `describe`/`it`, mais `groupe`/`test` :

```js
const { charger } = require('./harness');
const { erpDeTest, inscrit } = require('./fixtures');
const { groupe, test, egal, contient } = global.T;

const { erp } = charger(erpDeTest({ config: { VOC_MEMBRE: 'licencie' } }));

groupe('ce que fait la fonction', () => {
  test('dans le cas nominal', () => {
    egal(erp.maFonction_('entree'), 'sortie');
  });
});
```

Assertions disponibles : `egal`, `profond`, `vrai`, `faux`, `contient`, `absent`, `leve`.

`charger()` évalue `bundle/Code.gs` dans un contexte `vm` avec des doubles des services Google, et renvoie `{ erp, etat, classeur }` :

- `erp` : toutes les fonctions globales du bundle
- `etat` : ce que le code a produit (`etat.brouillons`, `etat.envoyes`, `etat.triggers`, `etat.logs`)
- `classeur` : le classeur ERP simulé, pour vérifier ce qui a été écrit

Ce qu'on teste utilement : les fonctions pures, les règles de gestion, le contrôle d'accès, ce qui part dans un mail, ce que `setup()` écrit. Ce qu'on ne teste pas ici : qu'un PDF soit réellement produit ou qu'un mail soit réellement remis — c'est l'API Google, pas notre logique. Si un test ne peut couvrir qu'une partie du comportement, dites-le dans son nom ou dans un commentaire.

**Un test doit échouer pour une bonne raison.** Un test qui passe quoi qu'on fasse ne sert à rien. Si vous corrigez un bug, écrivez d'abord le test qui le reproduit et vérifiez qu'il échoue avant votre correction.

## Ce qui est le plus utile en ce moment

- **Des tests d'interface.** `App.html` fait 2 450 lignes et n'a que la vérification syntaxique. Playwright sur l'application déployée serait un apport considérable.
- **Un retour d'une structure dont le vocabulaire ne rentre pas** dans la table de `docs/adaptation.md`. Ce sont les cas qui révèlent ce qui est encore trop spécifique.
- **La fragilité de `ModeleEco.gs`.** Il écrit des formules dans un classeur tiers en se fiant à sa mise en page. C'est la partie la moins testée et celle qui casse le plus facilement.
- **Les mentions légales hors de France.** `MENTION_TVA`, `PENALITES_INTERV` et `Convention.gs` supposent le droit français.

## Proposer une modification

Une branche, un sujet. Dans la description, dites ce que ça change pour la personne qui utilise l'ERP, pas seulement ce que ça change dans le code. Si ça modifie un mail envoyé à une famille ou un document généré, collez le avant et l'après.
