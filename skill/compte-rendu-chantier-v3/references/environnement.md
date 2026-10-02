# Environnement d'exécution

## Dépendances

```bash
pip install lxml openpyxl fonttools        # ajouter --break-system-packages si pip le demande
```

Facultatif, mais recommandé pour la hauteur des lignes et les contrôles :

```bash
apt-get install -y libreoffice-calc poppler-utils   # si l'environnement le permet (réseau, droits)
```

Polices **Denim INK** (Medium, Medium Italic, WD SemiBold), fournies par José à chaque session. Elles sont **sous licence** : ne jamais les recopier ailleurs, ni les mettre dans un dépôt ou un livrable.

```bash
mkdir -p ~/.fonts && cp <polices>.otf ~/.fonts/ && fc-cache -f
```

Si un skill `xlsx` est présent (dossier `.../xlsx/scripts` avec `office/soffice.py`), exporter `XLSX_SKILL_SCRIPTS=<ce dossier>` : `ajuster_hauteurs.py` l'utilise pour lancer LibreOffice dans les environnements isolés.

## Selon l'environnement

| Environnement | LibreOffice | Hauteurs de ligne | Contrôles |
|---|---|---|---|
| Session Cowork / Claude Code (conteneur) | installable en général | mesurées (fiables) | recalcul + rendu PDF |
| Projet claude.ai (exécution de code) | souvent **indisponible** | **estimées** (fontTools) : jamais moins de lignes que LibreOffice sur le test du CRC-16 (193 / 216 identiques, 23 avec 1 ou 2 lignes de plus) | ouverture et XML seulement |

Sans LibreOffice, **le dire à José** dans la réponse et lui demander un coup d'œil à l'aperçu avant impression. `CR_SANS_LIBREOFFICE=1` force l'estimation (tests).

## Contrôles avant livraison

1. L'archive s'ouvre (`zipfile.testzip()` renvoie `None`) et chaque partie XML se relit avec lxml.
2. Aucune note restante dans les onglets d'observations (`lire_notes.py` : « 0 note(s) » après intégration).
3. Si LibreOffice est disponible : recalcul (script `recalc.py` du skill `xlsx`) → seule erreur admise : `Page de garde!A6` (#VALUE!, image du logo en données enrichies, sans conséquence dans Excel) ; conversion PDF → pas de page blanche, nombre de pages cohérent avec le CR précédent.
4. Relire le rapport JSON : alertes, réouvertures, insertions.

## Limites connues

- Classeur à l'**ancien format** (sans colonne N°) : le convertir d'abord (`migration.md`). Opérations déjà
  configurées : HONGUEMARE, DUCLAIR « LE MIT » (`scripts/operations/`).
- Le code d'un tableau encore vide est retrouvé par le Référentiel (colonne G « Titre de section ») ou, pour un lot,
  par le numéro de l'onglet.
- Le rendu exact d'Excel n'est pas vérifiable dans l'environnement : seul José peut contrôler l'aperçu avant impression.
