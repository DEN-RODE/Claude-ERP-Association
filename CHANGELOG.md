# Changelog

Format inspiré de [Keep a Changelog](https://keepachangelog.com/fr/1.1.0/). Versions selon [SemVer](https://semver.org/lang/fr/).

## [2.0.0] - 2026-10-26

Première version publique, générique.

- 14 modules dans `src/`, concaténés par `build.js` dans `bundle/Code.gs`.
- Couche vocabulaire (`Libelles.gs`) : adhérent / cours / stage / lieu configurables depuis l'onglet Config.
- Toutes les données propres à une structure retirées du code ; `tests/aucune-donnee.test.js` le vérifie.
- Harness de tests Node sans dépendance : 251 tests.
- CI GitHub Actions : build, bundle à jour, syntaxe, tests, absence de données réelles.
- Documentation : installation, architecture, adaptation, contribution.

### 2026-10-27
- Publication sur GitHub, ajout des badges, du CHANGELOG, de `CLAUDE.md` et des modèles d'issues.

[2.0.0]: https://github.com/DEN-RODE/Claude-ERP-Association/releases/tag/v2.0.0
