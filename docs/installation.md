# Installation

Compter une heure la première fois. Il faut un compte Google et savoir ouvrir un Google Sheet.

## 1. Préparer les fichiers Google

Avant de toucher au code, créer ce dont l'ERP a besoin. Tout n'est pas obligatoire : les fonctions dont les identifiants manquent se désactivent proprement.

| Fichier à créer | Obligatoire | Ce qu'il sert |
|---|---|---|
| Un Google Sheet **ERP** | oui | le classeur de l'ERP, celui qui portera le script |
| Un Google Form d'inscription aux cours, et son classeur de réponses | oui | les inscriptions |
| Un Google Form d'inscription aux stages, et son classeur de réponses | non | si vous n'organisez pas de sessions courtes, laissez vide |
| Un Google Doc modèle de devis / facture | non | la facturation vers des collectivités |
| Un classeur de modèle économique | non | le prévisionnel |
| Des dossiers Drive (devis, justificatifs, visuels) | non | créés automatiquement si vous laissez vide |

Pour chacun, l'identifiant est la partie centrale de l'URL :

```
https://docs.google.com/spreadsheets/d/1AbCdEfGhIjKlMnOpQrStUvWxYz0123456789/edit
                                      ^-------------- identifiant --------------^
```

## 2. Construire le bundle

```bash
git clone https://github.com/DEN-RODE/Claude-ERP-Association.git
cd Claude-ERP-Association
node build.js v2.0.0
node tests/run.js        # 251 tests, pour vérifier que rien n'est cassé
```

Le dossier `bundle/` contient alors `Code.gs`, `Index.html`, `Styles.html`, `App.html` et `appsscript.json`. Ces fichiers sont aussi versionnés dans le dépôt : si vous n'avez pas Node, vous pouvez les prendre tels quels.

## 3. Coller dans Apps Script

1. Ouvrir le Google Sheet **ERP**, puis **Extensions > Apps Script**.
2. Dans le fichier `Code.gs` de l'éditeur, tout remplacer par le contenu de `bundle/Code.gs`.
3. Créer trois fichiers HTML (bouton **+** à côté de « Fichiers », puis **HTML**), nommés exactement :
   - `Index` → contenu de `bundle/Index.html`
   - `Styles` → contenu de `bundle/Styles.html`
   - `App` → contenu de `bundle/App.html`

   Les noms doivent être exacts : `include('Styles')` les cherche par ce nom.
4. Sauvegarder.

### Avec clasp, si vous préférez

```bash
npm install -g @google/clasp
clasp login
clasp clone <ID_DU_SCRIPT>
# copier bundle/* dans le dossier cloné, puis
clasp push
```

`.clasp.json` est ignoré par git : il contient l'identifiant de votre projet.

## 4. Lancer setup()

Dans l'éditeur Apps Script, choisir la fonction `setup` dans la liste déroulante et cliquer **Exécuter**.

Google demande des autorisations la première fois. L'écran « Cette application n'est pas validée » est normal pour un script personnel : **Paramètres avancés > Accéder à (nom du projet)**. Les autorisations demandées correspondent aux services utilisés : Sheets, Drive, Gmail, Agenda, Documents.

`setup()` crée les 13 onglets, pré-remplit l'onglet `Config` avec toutes les clés (valeurs d'identité vides), écrit deux créneaux d'exemple, et installe les déclencheurs. Il est relançable sans risque : il n'écrase aucune valeur déjà saisie et n'ajoute que les clés et colonnes manquantes.

## 5. Remplir l'onglet Config

C'est la seule étape qui demande de la saisie. Dans le classeur ERP, onglet `Config`, colonne `Valeur`. La colonne `Description` dit à quoi sert chaque clé.

Le minimum pour que l'ERP démarre :

```
NOM_STRUCTURE        Mon Association
EMAIL_STRUCTURE      contact@mon-association.example
ID_COURS_REPONSES    <identifiant du classeur de reponses du Form cours>
PIN                  <4 caracteres minimum>
```

Fortement conseillé dans la foulée :

```
TEL_STRUCTURE        01 02 03 04 05
ADRESSE_STRUCTURE    1 rue de l'Association, 00000 Villeneuve
IBAN_TITULAIRE       MON ASSOCIATION
IBAN                 FR...
BIC                  ...
DEBUT_SAISON_COURS   2026-09-07        (le lundi de la premiere seance)
VACANCES_COURS       2026-10-17>2026-11-01; 2026-12-19>2027-01-03
```

Puis **menu ERP > Vérifier la configuration** : il liste ce qui manque, en séparant le bloquant de l'optionnel.

**Sans code PIN, l'ERP refuse l'accès.** Ce n'est pas un oubli : un déploiement Apps Script peut être atteignable par plus de monde qu'on ne croit, et ces données concernent des mineurs. Un PIN de moins de 4 caractères est traité comme absent.

## 6. Déployer l'application web

**Déployer > Nouveau déploiement > Application Web**.

| Réglage | Valeur | Pourquoi |
|---|---|---|
| Exécuter en tant que | **Moi** | le script a besoin des droits du propriétaire sur les classeurs et Gmail |
| Qui a accès | au plus restrictif qui fonctionne pour votre équipe | voir ci-dessous |

Sur « Qui a accès » :

- **Moi uniquement** : le plus sûr, mais vous seul pouvez utiliser l'ERP.
- **Les utilisateurs de mon organisation** (Workspace) : le bon réglage pour une équipe.
- **Tout le monde** : à éviter. Même avec le code PIN, cela expose l'URL publiquement. Si vous n'avez pas Workspace et que plusieurs personnes doivent accéder à l'ERP, c'est le seul choix possible — mettez alors un PIN long et non devinable, et changez-le quand quelqu'un quitte l'équipe.

L'URL donnée en fin de déploiement est celle de l'ERP. Elle s'ajoute à l'écran d'accueil d'un téléphone comme une application.

## 7. Coller les libellés dans le Google Form

Dernière étape, et celle qu'on oublie. Dans l'ERP, écran des cours, chaque créneau affiche un libellé à copier dans la question « créneau souhaité » du Google Form. Ce libellé porte un code entre crochets (`[C1]`) qui rend le rattachement des inscriptions robuste aux retouches ultérieures du Form.

Sans ces codes, le rattachement se fait sur le texte et casse le jour où quelqu'un ajoute « COMPLET » devant un créneau.

## Mise à jour vers une version suivante

1. `git pull && node build.js && node tests/run.js`
2. Recoller `bundle/Code.gs` et les fichiers HTML modifiés dans l'éditeur.
3. Relancer `setup()` — il ajoute les nouvelles clés de configuration et les nouvelles colonnes, sans toucher à vos données.
4. **Déployer > Gérer les déploiements > (crayon) > Nouvelle version**.

Si vous avez modifié le code vous-même, travaillez dans `src/` et reconstruisez, jamais directement dans le bundle : il est regénéré à chaque build.

## Problèmes courants

**« Onglet "X" introuvable dans le classeur ERP. Lance setup(). »** — `setup()` n'a pas été exécuté, ou le script n'est pas lié au bon classeur. Vérifier la propriété de script `ERP_SS_ID`.

**« Réservé au propriétaire du script. »** — `setup()` et l'installation des déclencheurs ne s'exécutent que depuis l'éditeur ou le menu du Sheet, par le propriétaire du classeur. C'est voulu.

**Les inscriptions n'apparaissent pas.** — Trois causes, dans l'ordre : `ID_COURS_REPONSES` pointe sur le mauvais classeur ; le déclencheur `onInscriptionCours` n'est pas installé (menu **ERP > Réinstaller les déclencheurs**) ; le libellé du créneau dans le Form ne correspond à rien dans l'onglet `Cours` (recoller les libellés avec leur code).

**Un écran affiche des données périmées.** — Bouton Actualiser en haut à droite, qui vide les deux niveaux de cache.

**Le statut « payé » ne remonte pas.** — Il se lit sur la **couleur de fond** de la cellule du nom dans le classeur de réponses, pas sur une colonne. Un vert Google Sheets standard suffit.
