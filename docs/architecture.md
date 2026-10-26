# Architecture

## Vue d'ensemble

```
 Google Forms                    Google Sheets                    Gmail / Drive / Agenda
 (inscriptions)                  (donnees)                        (sorties)
      |                               |                                   ^
      |  onFormSubmit                 |  SpreadsheetApp                   |  brouillons
      v                               v                                   |
 +------------------------------------------------------------------------+------+
 |                        Apps Script (bundle/Code.gs)                           |
 |                                                                               |
 |   Api.gs ---- code PIN ---- routeur unique ---- Cache.gs                      |
 |      |                                            |                           |
 |      +-- Inscriptions  Presences  Compta  Factures  Mails  ModeleEco  ...     |
 |      |                                                                        |
 |      +-- Util.gs (fonctions pures)   Config.gs (parametres)  Libelles.gs      |
 +-------------------------------------------------------------------------------+
                                     ^
                                     |  google.script.run.api(action, payload, pin)
                                     |
                      Index.html + Styles.html + App.html
                      (application d'une seule page, sans framework)
```

Le classeur ERP est la base de données. Les classeurs de réponses de formulaires, le classeur du modèle économique et les dossiers Drive sont des sources externes, raccordées par leur identifiant depuis l'onglet `Config`.

## Un seul point d'entrée

Tout passe par `api(action, payload, pin)` dans `Api.gs`. Trois conséquences utiles :

1. **Le contrôle d'accès est à un seul endroit.** `checkPin_()` est appelé avant de résoudre l'action. Une route ajoutée est protégée sans que son auteur ait à y penser.
2. **Le cache est à un seul endroit.** `estCachable_()` décide, par action, si la réponse peut être servie depuis le cache. Les écritures invalident ce qu'elles touchent via la table `INVALIDE` de `Cache.gs`.
3. **Les lectures se regroupent.** `apiBloc_({ actions: [...] })` exécute jusqu'à 24 lectures en un seul aller-retour. Une lecture qui échoue renvoie son erreur sans emporter les autres.

Le point 3 est une réponse à une contrainte réelle d'Apps Script : les appels d'un même utilisateur sont **sérialisés** côté Google. Seize appels lancés « en parallèle » depuis le navigateur s'empilent en file. Les regrouper divise la latence perçue par seize, et les caches internes (inscrits, paramètres, classeurs ouverts) sont partagés par les lectures du même bloc.

## Le cache, deux niveaux

**Côté serveur** (`Cache.gs`), `CACHE_LECTURES` associe à chaque action une durée de vie. Les lectures coûteuses (la liste des relances à faire, qui croise inscriptions, paiements et historique) vivent 30 minutes. `INVALIDE` fait l'inverse : une écriture déclare quelles lectures elle périme. Un déclencheur toutes les dix minutes réchauffe les lectures les plus lourdes, pour qu'un écran ouvert à froid lise une réponse déjà prête.

**Côté navigateur** (`App.html`), un cache *stale-while-revalidate*. L'écran s'affiche instantanément avec la dernière version connue, un appel part en arrière-plan, et l'écran se corrige à la réponse. Sur une connexion mobile, c'est la différence entre « lent » et « instantané ».

Le bouton Actualiser vide les deux niveaux.

## Où est la donnée de vérité

| Donnée | Vérité | Pourquoi |
|---|---|---|
| Inscriptions | classeur de réponses du Form | là où Google les écrit, personne ne les recopie |
| Statut de paiement | **couleur de fond** de la cellule du nom, dans ce même classeur | la trésorière coloriait déjà ses lignes, zéro double saisie |
| Suivi, relances, notes | onglet `Suivi` du classeur ERP | données ajoutées par l'ERP, indexées sur l'identifiant d'inscription |
| Créneaux et tarifs | onglets `Cours` et `Stages` | modifiables à la main en cours de saison |
| Paramètres | onglet `Config` | aucun redéploiement pour changer une valeur |
| Prévisionnel | classeur du modèle économique | continue de vivre comme un tableur, avec ses formules |

Ce tableau est le choix central du projet : **l'ERP ne confisque rien**. Chaque donnée reste là où un humain peut l'ouvrir et la corriger sans passer par l'interface. C'est ce qui le rend adoptable par une association, et c'est aussi ce qui le rend fragile : une colonne déplacée à la main casse une lecture. `mapHeaders_()` dans `Inscriptions.gs` existe pour cela — il retrouve les colonnes par expression régulière sur l'en-tête normalisé, pas par position.

## Rattachement d'une inscription à un créneau

Le point le plus fragile de tout l'ERP, et celui qui a le plus de tests. Une inscription qui ne se rattache à aucun créneau disparaît des listes d'appel et des effectifs, silencieusement.

Le texte vient d'un Google Form que des humains retouchent en cours de saison : « COMPLET », « 2 places restantes », un libellé réécrit. `matchCours_()` essaie donc trois niveaux, dans l'ordre :

1. **Le code entre crochets** (`[C3]`), que `libelleCodeForm_()` génère et qu'on colle dans le Form. Insensible à tout ce qui est écrit autour.
2. **La colonne `Cle` cherchée dans le texte** (« Mercredi 11h00 »), après normalisation (accents, casse, apostrophes, espaces insécables).
3. **Repli sur le jour et l'heure de début** retrouvés dans le texte, si les deux premiers échouent.

Le niveau 1 suffit si le Form est bien configuré. Les niveaux 2 et 3 sont là pour le jour où il ne l'est pas. Un test ferme la boucle : pour chaque créneau, le libellé généré par `libelleCodeForm_()` doit se rattacher à ce créneau et à lui seul.

## La règle du montant incertain

Dans `Mails.gs`, `montantDu_()` renvoie `null` plutôt qu'un chiffre quand l'inscription porte un motif de tarif particulier : fratrie, convention avec une collectivité, quotient familial, demi-année, tarif à la carte, ou un motif propre à la structure listé dans `REDUC_MOTIFS`.

Dans ce cas la relance ne chiffre rien et dit qu'un humain confirmera le montant. Pour une fratrie, un seul enfant au montant incertain annule le total de **toute** la fratrie.

C'est une décision de produit, pas une limite technique : réclamer 350 € à une famille qui en doit 280 coûte plus cher que d'envoyer un mail de plus.

## Rien ne part sans un clic

Tous les mails aux familles passent par `GmailApp.createDraft()`. L'ERP affiche un aperçu, exactement le HTML qui partira, et l'envoi demande une action humaine. Aucun chemin de code n'appelle `sendEmail` vers une famille. Un test le vérifie.

Les mails internes (récapitulatif quotidien à l'équipe) sont la seule exception, et ils ne vont qu'aux adresses listées dans `DIGEST_EMAILS`.

## Le bundle

Apps Script charge tous les `.gs` d'un projet dans un espace global unique, dans un ordre non garanti. `build.js` concatène `src/*.gs` dans un ordre explicite — utilitaires, configuration, vocabulaire, modules métier, cache, routeur — et refuse de construire si un fichier de `src/` n'est pas déclaré dans cet ordre.

Les trois fichiers HTML sont copiés tels quels. `Index.html` est un gabarit Apps Script (`<?= nom ?>`, `<?!= include('App') ?>`), les deux autres sont inclus par lui.

## Ce que l'architecture ne fait pas

- **Pas de migration de schéma.** `setup()` ajoute les onglets et les colonnes manquants. Il ne renomme rien, ne convertit rien, ne supprime rien.
- **Un seul niveau de droits.** Le code PIN ouvre tout. Pas de rôle lecture seule, pas de permission par écran. L'onglet `Log` journalise les actions, mais ce n'est pas un contrôle d'accès.
- **Pas de transaction.** Deux personnes qui écrivent la même ligne au même instant : la dernière gagne. `LockService` protège les sections critiques (numérotation des factures), pas l'ensemble.
- **Pas de test d'interface.** `App.html` fait 2 450 lignes et n'est couvert que par une vérification syntaxique.
