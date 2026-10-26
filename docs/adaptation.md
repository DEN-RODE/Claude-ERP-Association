# Adapter l'ERP à votre structure

Trois niveaux, du moins coûteux au plus coûteux. Commencez par le premier : il couvre la plupart des cas.

## Niveau 1 : la feuille de calcul (zéro ligne de code)

Tout ce qui est propre à une structure vit dans l'onglet `Config`. Aucun nom, aucun IBAN, aucun identifiant Google n'est codé en dur dans le code publié — un test (`tests/aucune-donnee.test.js`) le vérifie à chaque exécution de la suite.

### Le vocabulaire affiché

Les mots qui apparaissent à l'écran, dans les mails et dans les documents viennent de huit clés :

| Clé | Défaut | École de musique | Club sportif | MJC | Théâtre |
|---|---|---|---|---|---|
| `VOC_MEMBRE` | adherent | élève | licencié | adhérent | comédien |
| `VOC_MEMBRES` | adherents | élèves | licenciés | adhérents | comédiens |
| `VOC_COURS` | cours | cours | entraînement | atelier | atelier |
| `VOC_COURS_PL` | cours | cours | entraînements | ateliers | ateliers |
| `VOC_STAGE` | stage | stage | camp | session | stage |
| `VOC_STAGE_PL` | stages | stages | camps | sessions | stages |
| `VOC_LIEU` | sur place | au conservatoire | au gymnase | au local | au théâtre |
| `VOC_ACTIVITE` | activite | enseignement | entraînement | animation | pratique |

`VOC_LIEU` se lit dans des phrases comme « à remettre **au gymnase** » ou « directement **au local**, un reçu vous sera remis ». Écrivez-le avec sa préposition.

`VOC_ACTIVITE` nomme la nature de rémunération de terrain, par opposition à l'administratif et au technique, dans le modèle économique.

### L'identité et l'apparence

```
NOM_STRUCTURE     le nom affiche partout
BASELINE          sous-titre du bandeau de mail, ex. "Ecole de musique associative"
MOT_ACCENTUE      un mot du nom a teinter dans la couleur d'accent (vide = nom entier en noir)
COULEUR_ACCENT    titre et intertitres des mails
COULEUR_FILET     les deux filets du bandeau
SITE_WEB          affiche en pied de mail
```

Le bandeau des mails n'utilise **aucune image**. C'est volontaire : pas de pièce jointe, pas de « télécharger les images » chez le destinataire, et le même rendu dans Gmail, Outlook et sur mobile. Si vous voulez votre logo dans l'interface, renseignez `ID_LOGO` — il s'affiche dans l'ERP, pas dans les mails.

### Les règles de gestion

```
ADHESION                   adhesion annuelle par personne (0 = aucune)
SEUIL_MIN_COURS            inscrits minimum pour maintenir un creneau
SEUIL_MIN_STAGE            inscrits minimum par groupe de session
RELANCE_APRES_JOURS        delai avant la premiere relance
RELANCE_INTERVALLE_JOURS   delai minimum entre deux relances
RELANCES_AUTO              oui/non : brouillons crees chaque matin, ou a la demande
CONFIRMATION_AUTO          oui/non : brouillon de confirmation a chaque inscription
CONTACT_PAIEMENTS          qui suit les arrangements de paiement
REDUC_MOTIFS               vos motifs de tarif particulier (voir ci-dessous)
```

`REDUC_MOTIFS` mérite une explication. L'ERP refuse de chiffrer un montant dans un mail quand l'inscription porte un motif de tarif particulier. Les motifs génériques sont déjà couverts (fratrie, convention, quotient familial, demi-année, tarif à la carte). Ajoutez les vôtres, séparés par des point-virgules : le nom d'une commune conventionnée, un dispositif local, un partenariat CAF. Une inscription portant un de ces mots dans ses notes, ses réductions ou son commentaire ne recevra jamais de montant chiffré.

## Niveau 2 : le code, pour des cas particuliers

### Un troisième type d'activité

L'ERP connaît deux sources : `cours` (régulier hebdomadaire) et `stages` (intensif sur quelques jours). Si vous avez une troisième forme — des séances à l'unité, des événements ponctuels — il n'y a pas de clé de configuration pour ça. Le chemin le plus court est un troisième Google Form rattaché à `stages`, avec des clés qui distinguent vos événements.

Ajouter une vraie troisième source demande de toucher `Inscriptions.gs` (`responseSheet_`, `loadEleves_`), `Cache.gs` et les écrans concernés. Compter une journée.

### Le statut de paiement

Il se lit sur la **couleur de fond** de la cellule du nom, dans le classeur de réponses du formulaire. Vert Google Sheets = payé.

Cette convention vient d'un usage réel : la trésorière coloriait déjà ses lignes en vert au fil des encaissements. La garder évite toute double saisie, et la fonction `couleur_()` reconnaît plusieurs nuances de vert pour tolérer la nuance choisie par chacun.

C'est néanmoins une convention, pas un schéma de données. Elle a deux défauts : rien n'empêche quelqu'un de colorier par erreur, et l'information n'est pas requêtable. Pour la remplacer par une colonne dédiée, le point d'entrée est dans `Inscriptions.gs` :

```js
// aujourd'hui
e.paye = cNom === 'vert' ? 'Oui' : '';

// avec une colonne "Paye" dans la feuille de reponses
e.paye = /oui|paye|1/i.test(o.paye) ? 'Oui' : '';
```

Les tests de `couleur_()` dans `tests/inscriptions.test.js` deviennent alors inutiles, ceux de `montantDu_()` restent valides.

### Les mentions légales

`MENTION_TVA` et `PENALITES_INTERV` portent les mentions françaises (article 293 B du CGI, articles L441-10 et D441-5 du code de commerce), et `DELAI_PAIEMENT_INTERV` est borné par le maximum légal français entre professionnels. Hors de France, ces trois clés sont à réécrire, et `Convention.gs` avec elles : son texte suppose un droit français des associations.

## Niveau 3 : ce qui ne s'adapte pas sans réécriture

Autant le dire tout de suite, pour que vous n'y perdiez pas une semaine.

**Des encadrants salariés.** `ModeleEco.gs` et `FacturesInterv.gs` supposent des intervenants indépendants qui **facturent** l'association : trois natures de rémunération, lissage sur l'année, factures entrantes avec mentions légales. Si vos encadrants sont salariés, il faut des bulletins, des charges patronales et des déclarations sociales — un autre logiciel, pas une adaptation de celui-ci. Les autres modules (inscriptions, paiements, présences, facturation sortante) restent utilisables tels quels.

**Sortir de Google Workspace.** 25 appels à `SpreadsheetApp`, 23 à `DriveApp`, 31 à `Utilities`, 16 à `CacheService`, 11 à `GmailApp`. Ce n'est pas un ERP portable, c'est un ERP Google Workspace. Une version PostgreSQL + API serait une réécriture complète, en conservant au mieux les règles de gestion.

**Des rôles et des permissions.** Le code PIN ouvre tout. Pas de rôle lecture seule, pas de restriction par écran. Pour une équipe de quatre personnes qui se font confiance, c'est suffisant ; pour vingt bénévoles, non.

**Du multi-structures.** Une installation = une structure. Un réseau de plusieurs antennes aurait besoin d'un découpage des données qui n'existe pas ici.

## Avant de déployer chez vous

- [ ] `node tests/run.js` passe sur votre fork
- [ ] `node tests/run.js aucune-donnee` passe **après** vos modifications, et aucune donnée de votre structure n'est dans le code (elle va dans l'onglet `Config`)
- [ ] un code PIN de plus de 4 caractères est posé
- [ ] le déploiement web est au périmètre le plus restreint qui fonctionne pour votre équipe
- [ ] les libellés avec leur code `[C1]` sont collés dans le Google Form
- [ ] vous avez vérifié ce que l'ERP répond à une inscription portant un tarif particulier : aucun montant chiffré
- [ ] vous savez où sont vos obligations en matière de données personnelles (information des familles, durée de conservation, registre), puisqu'elles restent les vôtres
