# Reprise de la session (30/09/2026)

## À fournir au démarrage

L'environnement de travail est temporaire : ce qui n'est pas dans ce dépôt sera perdu.

1. **Le classeur CR sur lequel le fichier de routage a été établi**, a priori le CRC-15 du 22/09/2026. Il doit s'agir exactement de la même version, car le routage cite des numéros de ligne de l'ancien format.
2. **Le fichier de routage** à intégrer, qui correspond à la réunion du 29/09/2026 (CRC-16).
3. **Les polices Denim INK** : Medium, Medium Italic et WD SemiBold.

## Déroulé prévu

1. Réinstaller l'environnement : `apt-get install libreoffice-calc poppler-utils`, les polices dans `~/.fonts` puis `fc-cache -f`, et `pip install lxml openpyxl fonttools`.
2. **Copie de test finale** : `tools/migration_cr.py`, puis `tools/ajuster_hauteurs.py`, sur le classeur fourni. Rapport mis à jour. Validation par José.
3. **Intégration du fichier de routage dans le nouveau format** :
   - le routage cite les lignes de l'ancien classeur. Le script de migration connaît la correspondance entre ancienne ligne et N° (`new_of_all`, `info[...]['num']`). Il faut l'exporter en table de correspondance, puis traduire le routage en N° ;
   - intégration selon les règles du registre `docs/decisions-HONGUEMARE.md` : structure du texte, règle A, « sans réponse depuis », statuts, PM, Retard, bordures des lignes créées, hauteurs ;
   - produire le CRC-16 au nouveau format.
4. Gabarit du **résumé routé** (Excel, une seule feuille) et script d'intégration réutilisable.
5. **Skill V3** : règles, scripts et prérequis, destiné à Cowork ou à un projet claude.ai.
6. Premier essai réel depuis Cowork.
