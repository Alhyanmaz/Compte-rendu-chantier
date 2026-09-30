#!/bin/sh
# Recopie les scripts de tools/ (source de vérité) dans le skill V3, avant packaging.
set -e
cd "$(dirname "$0")/.."
DST=skill/compte-rendu-chantier-v3
for f in cr_xml.py cr_texte.py cr_classeur.py lire_notes.py integrer_cr.py relire_cr.py ajuster_hauteurs.py; do
  cp "tools/$f" "$DST/scripts/$f"
done
cp docs/memo-notes.md "$DST/references/memo-notes.md"
echo "skill synchronisé"
