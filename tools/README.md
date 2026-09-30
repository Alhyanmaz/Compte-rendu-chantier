# Outils de migration du CR

Chaîne de production d'une copie migrée (voir `docs/decisions-HONGUEMARE.md`) :

1. `python migration_cr.py SOURCE.xlsx ETAPE1.xlsx rapport.json` : structure, N°, texte, statuts, MFC, onglets masqués.
2. `python ajuster_hauteurs.py ETAPE1.xlsx FINAL.xlsx` : hauteurs de ligne pour que tout le texte soit visible dans Excel.

Chaîne hebdomadaire (CR précédent annoté par José en réunion + transcription audio, voir `docs/memo-notes.md`) :

0. `python lire_notes.py CR_ANNOTE.xlsx notes.json` : liste des notes avec ce qu'elles désignent ; Claude rédige
   OPERATIONS.json (notes reformulées à l'appui de l'audio, champs `note`, `source`, `routage`, `doute`, `non_route`).
1. `python integrer_cr.py CR_PRECEDENT.xlsx OPERATIONS.json ETAPE1.xlsx rapport.json` : page de garde, masquage des
   lignes soldées au CR précédent, remise en noir du rouge et du gris, mises à jour (« Au JJ/MM/AAAA … » en rouge dans
   HISTORIQUE), nouvelles lignes (N° suivant, bordures du classeur, insertion de lignes si la réserve manque),
   relances automatiques, texte condensé recalculé, sauts de page.
2. `python ajuster_hauteurs.py ETAPE1.xlsx FINAL.xlsx`.

Relecture (José corrige le CR livré dans Excel avant diffusion, puis le redonne avec la version livrée) :

1. `python relire_cr.py CR_LIVRE.xlsx CR_CORRIGE.xlsx ETAPE1.xlsx rapport.json` : compare les deux versions
   ligne par ligne (par N°) et ne retouche que ce qui a changé (voir l'en-tête du script). Aucun masquage, aucune
   relance automatique.
2. `python ajuster_hauteurs.py ETAPE1.xlsx FINAL.xlsx`.

Outils communs : `cr_classeur.py` (lecture du modèle, écriture des cellules, couleurs, Retard, mise en page),
`cr_texte.py` (règles de texte), `cr_xml.py` (édition XML du classeur).

`OPERATIONS.json` : `{"date_cr": "JJ/MM/AAAA", "crc": "CRC-NN", "operations": [...]}` avec
`{"type": "maj", "num", "texte", "pour_le" ("+N" jours ou date), "fait_le"}` ou
`{"type": "nouvelle", "onglet", "code", "section", "texte", "pour_le", "fait_le"}`.

Prérequis :
- Python 3, `lxml`, `openpyxl` ; `cr_xml.py` (outil du skill V1, copié ici).
- LibreOffice Calc et les polices **Denim INK** (Medium, Medium Italic, WD SemiBold) installées dans `~/.fonts`
  (`fc-cache -f`). Les polices ne sont **pas** dans le dépôt : elles sont sous licence, à fournir à chaque session.
- Dans un environnement sandboxé : `export XLSX_SKILL_SCRIPTS=<chemin du dossier scripts du skill xlsx>`.

Le classeur n'est jamais enregistré par openpyxl : seules les parties XML modifiées sont réécrites.
