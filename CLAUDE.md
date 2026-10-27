# ERP Association : notes pour Claude

ERP Google Workspace (Apps Script) pour petite association : inscriptions, paiements et relances, présences, facturation émise et reçue, prévisionnel, conventions. Vue d'ensemble dans `README.md`, détails dans `docs/`.

## Règles non négociables
- **Aucune donnée d'une structure réelle dans le code** (nom, mail, téléphone, IBAN, SIRET, ID de Sheet/Form, même en commentaire). Tout va dans l'onglet Config du classeur. Voir `CONTRIBUTING.md`.
- On modifie `src/`, jamais `bundle/` à la main : `node build.js` le regénère, et la CI échoue si `bundle/` n'est pas à jour.
- Zéro dépendance npm. Node 18+ uniquement pour le build et les tests.
- Les mails aux familles restent des brouillons Gmail (jamais d'envoi automatique). Les infos santé ne partent jamais dans un mail.
- Tarif particulier (fratrie, convention, QF, carte) : le mail ne chiffre rien.

## Avant chaque commit
```bash
node build.js && node tests/run.js            # suite complète
node tests/run.js aucune-donnee               # contrôle données réelles
```

## Règle de travail : tenir le repo GitHub à jour (demande de Mathieu)
Repo : https://github.com/DEN-RODE/Claude-ERP-Association (branche `main`).
À **chaque** modification du projet, sans attendre qu'on le redemande :
1. Rebuild `bundle/` et faire passer les tests.
2. Mettre à jour `README.md` (et `docs/` si besoin) quand une fonctionnalité, un module ou l'installation change ; garder le badge du nombre de tests à jour.
3. Ajouter une entrée datée dans `CHANGELOG.md` ; monter la version dans `package.json` et taguer (`vX.Y.Z`) pour une release.
4. Mettre à jour ce `CLAUDE.md` si une décision ou une contrainte change.
5. Commit en français, clair, puis push sur `main`.
