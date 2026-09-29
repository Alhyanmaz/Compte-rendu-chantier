# Outils de migration du CR

Chaîne de production d'une copie migrée (voir `docs/decisions-HONGUEMARE.md`) :

1. `python migration_cr.py SOURCE.xlsx ETAPE1.xlsx rapport.json` : structure, N°, texte, statuts, MFC, onglets masqués.
2. `python ajuster_hauteurs.py ETAPE1.xlsx FINAL.xlsx` : hauteurs de ligne pour que tout le texte soit visible dans Excel.

Prérequis :
- Python 3, `lxml`, `openpyxl` ; `cr_xml.py` (outil du skill V1, copié ici).
- LibreOffice Calc et les polices **Denim INK** (Medium, Medium Italic, WD SemiBold) installées dans `~/.fonts`
  (`fc-cache -f`). Les polices ne sont **pas** dans le dépôt : elles sont sous licence, à fournir à chaque session.
- Dans un environnement sandboxé : `export XLSX_SKILL_SCRIPTS=<chemin du dossier scripts du skill xlsx>`.

Le classeur n'est jamais enregistré par openpyxl : seules les parties XML modifiées sont réécrites.
