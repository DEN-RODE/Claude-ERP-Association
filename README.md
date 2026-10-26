# ERP Association

ERP de gestion pour une petite association qui fait de l'accueil de public : inscriptions, paiements, présences, facturation, prévisionnel. Il tourne entièrement sur Google Workspace, sans serveur, sans base de données et sans coût d'hébergement.

Il est né d'un besoin réel, dans une association de quatre personnes dont aucune n'a de temps pour de la saisie administrative. Il a été généralisé ici pour que d'autres structures puissent le reprendre : une école de musique, un club sportif, une MJC, une école de cirque, un atelier d'arts plastiques. Le vocabulaire affiché (adhérent, cours, stage, lieu) se change dans une feuille de calcul, sans toucher au code.

```
8 276 lignes de code  ·  14 modules  ·  251 tests  ·  0 dépendance  ·  0 € d'hébergement
```

## Ce que ça fait

**Inscriptions.** Les familles remplissent un Google Form. L'ERP lit les réponses, rattache chaque inscription à un créneau, calcule l'âge, repère les informations manquantes, et tient une liste d'attente séparée des inscrits confirmés.

**Paiements et relances.** Le suivi des règlements, les relances groupées par foyer (une famille de trois enfants reçoit un seul mail), et une règle stricte : quand un tarif particulier s'applique (fratrie, convention avec une commune, quotient familial, tarif à la carte), le mail ne chiffre **rien** et annonce qu'un humain confirmera le montant. Un montant faux envoyé à une famille coûte plus cher qu'un mail de plus.

**Rien ne part tout seul.** Tous les mails aux familles sont créés comme brouillons Gmail. L'envoi demande un clic humain. C'est un choix de conception, pas une limitation technique.

**Présences.** Feuilles d'appel par séance, avec les alertes santé remontées en tête de liste et jamais recopiées dans un mail.

**Facturation dans les deux sens.** Les devis et factures que la structure émet vers des collectivités et des écoles (Google Doc modèle vers PDF, numérotation suivie), et les factures que les intervenants indépendants adressent à la structure, avec les mentions légales obligatoires et le lissage sur l'année.

**Prévisionnel.** Un modèle économique qui part des effectifs réels et des rémunérations paramétrées, et dit où est le point d'équilibre par créneau.

**Conventions.** Génération de la convention annuelle de prestation avec les prestataires.

## Pourquoi Apps Script, et ce que ça coûte

C'est le choix structurant du projet, et il est assumé.

Une association de cette taille n'a ni budget d'hébergement, ni administrateur système, ni quiconque pour renouveler un certificat TLS un dimanche soir. Ses données sont déjà dans Google Sheets, ses formulaires dans Google Forms, sa boîte dans Gmail. Apps Script met le code **à l'intérieur** de cet écosystème : authentification Google gratuite, sauvegarde et historique de versions gratuits, aucun serveur à maintenir, et la trésorière peut continuer à ouvrir le classeur et corriger une cellule à la main. C'est ce dernier point qui fait la différence à l'usage : l'ERP ne confisque pas les données, il s'ajoute par-dessus.

Le prix à payer est réel :

- **Pas portable.** 25 appels à `SpreadsheetApp`, 23 à `DriveApp`, 31 à `Utilities`, 16 à `CacheService`. Sortir de Google Workspace serait une réécriture, pas un portage.
- **Pas de modules.** Apps Script partage un espace de noms global. D'où `build.js` (voir ci-dessous).
- **Pas de npm, pas de tests natifs.** D'où le harness de tests (voir ci-dessous).
- **Lenteur des appels.** Apps Script exécute les appels d'un même utilisateur en file, un par un. Un écran qui fait seize lectures « en parallèle » depuis le navigateur les empile en réalité. D'où le cache à deux niveaux et le regroupement des lectures.

Les trois contournements ci-dessous sont, honnêtement, la partie la plus intéressante du projet.

## Les trois contournements

### 1. `build.js` : des modules là où il n'y en a pas

Apps Script charge tous les fichiers `.gs` d'un projet dans un seul espace global, dans un ordre qu'il ne garantit pas vraiment. Déclarer `getConfig_()` dans un fichier et l'appeler depuis un autre marche ou ne marche pas selon l'ordre de chargement.

On développe donc en 14 modules lisibles dans `src/`, et `build.js` les concatène dans `bundle/Code.gs` suivant un ordre maîtrisé : les utilitaires purs d'abord, la configuration ensuite, le routeur web en dernier. C'est ce fichier unique qu'on colle dans l'éditeur Apps Script. Le script refuse de construire si un module de `src/` n'est pas déclaré dans l'ordre, ce qui évite d'oublier un fichier dans le bundle.

```bash
node build.js v2.0.0
# v2.0.0 : 14 modules, 5503 lignes -> bundle/Code.gs
```

### 2. `tests/harness.js` : tester du code qui n'existe que chez Google

Apps Script n'a pas de runner de tests, et son runtime n'existe que sur l'infrastructure Google. La plupart des projets Apps Script ne sont donc pas testés du tout.

Le harness charge `bundle/Code.gs` dans un contexte `vm` de Node et y injecte des doubles des services Google. Le double de classeur n'est pas un simple `{}` : c'est une structure en mémoire avec onglets, lignes, **couleurs de fond** et formules, parce que dans cet ERP le statut de paiement se lit sur la couleur de fond d'une cellule. `Utilities.computeDigest` fait du vrai SHA-256, sinon les tests du code PIN passeraient avec n'importe quel code.

```bash
node tests/run.js              # 251 tests
node tests/run.js mails        # un seul fichier
```

Zéro dépendance : pas de Jest, pas de `node_modules`, un lanceur de 110 lignes. Le compromis est explicite — ce harness teste la **logique métier**, pas l'API Google. Qu'un PDF soit réellement produit ou qu'un mail soit réellement remis ne se teste pas ici, et les tests le disent là où c'est le cas.

### 3. Cache à deux niveaux

Côté serveur, un cache à durée de vie par type de lecture, avec une table d'invalidation : une écriture qui touche les inscrits vide les lectures qui en dépendent, et elles seules. Un déclencheur réchauffe les lectures coûteuses toutes les dix minutes, pour que l'écran lise une réponse déjà prête. Côté navigateur, un cache *stale-while-revalidate* : l'écran s'affiche immédiatement avec la version connue, puis se corrige quand le serveur répond.

## Architecture

```
src/
  Util.gs            fonctions pures : dates, normalisation, montants, accès aux onglets
  Config.gs          paramètres, setup(), le SEUL fichier à renseigner
  Libelles.gs        couche vocabulaire (adhérent / cours / stage / lieu)
  Inscriptions.gs    lecture des réponses de formulaire, rattachement aux créneaux
  Compta.gs          dépenses, recettes, justificatifs
  Factures.gs        devis et factures émises (Google Doc -> PDF)
  Mails.gs           brouillons Gmail, relances groupées par foyer
  Divers.gs          agenda, rappels, médias, notes
  Presences.gs       appel et feuilles de présence
  ModeleEco.gs       prévisionnel et rémunération des intervenants
  FacturesInterv.gs  factures des prestataires vers la structure
  Convention.gs      convention annuelle de prestation
  Cache.gs           cache serveur et réchauffage
  Api.gs             point d'entrée web, code PIN, routeur unique
  Index.html         squelette de la page
  Styles.html        feuille de style
  App.html           interface (vanilla JS, pas de framework)
```

L'interface est une application d'une seule page en JavaScript sans framework, qui parle au serveur par un routeur unique : `google.script.run.api(action, payload, pin)`. Toutes les actions passent par le même point, donc le contrôle d'accès est à un seul endroit.

Détails dans [`docs/architecture.md`](docs/architecture.md).

## Installation

Il faut un compte Google. Compter une heure la première fois.

1. Créer un Google Sheet, l'ouvrir, puis **Extensions > Apps Script**.
2. Coller `bundle/Code.gs` dans le fichier de code, et créer trois fichiers HTML nommés `Index`, `Styles` et `App` avec le contenu de `bundle/`.
3. Exécuter la fonction `setup()` une fois. Elle crée les 13 onglets, pré-remplit la configuration et installe les déclencheurs.
4. Remplir l'onglet **Config** : nom, contacts, identifiants de vos Google Sheets et Forms, IBAN, et un code PIN de 4 caractères minimum.
5. Menu **ERP > Vérifier la configuration** pour voir ce qui manque encore.
6. **Déployer > Nouveau déploiement > Application Web**.

Pas à pas dans [`docs/installation.md`](docs/installation.md).

## Adapter à votre structure

Tout ce qui est propre à une structure tient dans **un seul fichier**, `Config.gs`, dont toutes les valeurs d'identité sont vides par défaut et se remplissent ensuite depuis la feuille de calcul. Aucun nom, aucun IBAN, aucun identifiant Google n'est codé en dur ailleurs : un test le vérifie à chaque exécution de la suite.

Le vocabulaire affiché se change dans la même feuille :

| Clé | École de musique | Club sportif | MJC |
|---|---|---|---|
| `VOC_MEMBRE` | élève | licencié | adhérent |
| `VOC_COURS` | cours | entraînement | atelier |
| `VOC_STAGE` | stage | camp | session |
| `VOC_LIEU` | au conservatoire | au gymnase | au local |

La couche `Libelles.gs` gère les formes grammaticales (`Lle_` élide devant une voyelle, `Ldu_` contracte en « du »), parce qu'un ERP qui écrit « le nom de l'licencié » dans un mail à une famille n'est pas utilisable.

Ce qui ne se configure pas et demande une décision de votre part :

- **Le statut de paiement se lit sur la couleur de fond de la cellule du nom**, dans la feuille de réponses du formulaire. C'est une convention héritée d'un usage réel : la trésorière coloriait déjà ses lignes en vert. Elle marche très bien à la main et ne demande aucune double saisie, mais c'est une convention, pas un schéma de données. Si vous préférez une colonne dédiée, c'est `couleur_()` et `toEleve_()` dans `Inscriptions.gs`.
- **La structure du modèle économique** suppose des intervenants indépendants qui facturent l'association, avec trois natures de rémunération (activité de terrain, administratif, technique). Si vos encadrants sont salariés, `ModeleEco.gs` ne correspond pas à votre réalité.

Détails dans [`docs/adaptation.md`](docs/adaptation.md).

## Données personnelles

Cet ERP manipule des données concernant des mineurs : noms, dates de naissance, coordonnées des responsables légaux, et des informations de santé saisies par les familles (allergies, traitements).

Les choix faits ici :

- Aucune donnée personnelle n'est écrite dans le stockage du navigateur. Le `localStorage` ne contient qu'un réglage d'affichage et le code PIN de session.
- Les informations de santé ne sont jamais recopiées dans un mail. Un test le vérifie.
- Les saisies de formulaire sont échappées avant tout rendu HTML, et préfixées avant toute écriture dans un Sheet (injection de formule).
- Le code PIN est comparé par empreinte SHA-256, avec blocage après cinq essais ratés.
- L'absence de code PIN **refuse** l'accès au lieu de l'ouvrir.

Cela ne fait pas de ce dépôt un outil conforme par défaut. Si vous le déployez, la responsabilité du traitement reste la vôtre : portée du déploiement Apps Script, durée de conservation, information des familles, registre de traitement. Le code PIN est une barrière raisonnable, pas un système d'authentification multi-utilisateurs avec des rôles.

## État du projet

En production dans une association depuis la rentrée 2026. La version générique publiée ici est la même base de code, avec la couche vocabulaire ajoutée et les données de la structure retirées.

Ce qui manque et que je sais :

- Pas de tests d'interface. `App.html` fait 2 450 lignes et n'est couvert que par la vérification syntaxique. Des tests Playwright sur la page déployée sont la prochaine étape.
- Pas de migration de schéma. `setup()` ajoute les colonnes manquantes, mais ne renomme rien et ne convertit rien.
- Un seul niveau de droits. C'est PIN ou rien : pas de rôle lecture seule, pas de journal par utilisateur au-delà de l'onglet `Log`.
- `ModeleEco.gs` écrit des formules dans un classeur tiers. C'est la partie la plus fragile et la moins testée, parce qu'elle dépend de la mise en page de ce classeur.

Les retours et les forks sont bienvenus, en particulier venant d'une structure dont le vocabulaire ne rentre pas dans la table ci-dessus.

## Licence

MIT. Voir [`LICENSE`](LICENSE).
