# Convertir un classeur à l'ancien format (sans colonne N°)

À faire **une fois par opération**, avant le premier CR au nouveau format (fait pour HONGUEMARE le 29/09/2026 et
DUCLAIR « LE MIT » le 02/10/2026). Le résultat est une **copie de test** que José valide avant de l'adopter.

## 1. Analyser (lecture seule, openpyxl)

- Onglets : `MOE-MOA`, `Concessionnaires` s'il existe, onglets de lots (nom commençant par deux chiffres), `Modele_lot`.
- Titres des sections de MOE-MOA (ligne au-dessus de chaque tableau) : chacun reçoit un **code** de N°
  (MOA, AMO, CT, SPS, MOE, ECO, BET.xxx…).
- Page de garde : cellule « CRC-NN », date à sa droite, heure dessous, ligne « PROCHAINE RÉUNION / DATE ».
- Coordonnees : ligne de chaque intervenant (organisme, représentants).
- Valeurs de FAIT LE et de POUR LE hors liste (texte libre, dates en texte), en distinguant lignes **visibles** et
  **masquées** (archivées).
- Doublons probables (même destinataire, même objet) parmi les lignes ouvertes ; caractères isolés ; dates aberrantes.

## 2. Écrire `scripts/operations/<operation>.py`

Partir de `duclair_mit.py` (le plus récent) : `section_codes`, `referentiel` (code → ligne(s) de Coordonnees),
`page_de_garde`, `statuts`, `organismes_attente`, `parasites`, `coord_fixes`, `doublons`, valeurs propres à
l'opération (`statut_special` / `mention_speciale`, lignes visibles seulement). Ajouter le mot-clé de détection dans
`operations/__init__.py` (texte de la page de garde) ou passer le nom en 4e argument.

**Ne rien trancher seul** : tout choix (code, conversion d'un statut libre, ligne à supprimer) est une interprétation,
listée dans le rapport et soumise à José.

## 3. Convertir et contrôler

```bash
cd scripts
python migration_cr.py ANCIEN.xlsx etape1.xlsx rapport.json [operation]
python ajuster_hauteurs.py etape1.xlsx OPERATION_CR_ACAU_CRC-NN_NOUVEAU_FORMAT.xlsx
```

Contrôles de `environnement.md`, plus : nombre d'observations visibles / masquées, aucune erreur de formule,
pages (comparer au PDF d'origine), rendu de quelques pages. Puis un essai d'intégration fictive (une mise à jour,
une ligne neuve dans un tableau vide) et de remoulinage, pour vérifier que les outils hebdomadaires lisent le classeur.

## 4. Rapport à José

Ce qui a été appliqué, interprétations à valider (codes, statuts), corrections faites, anomalies signalées
(non corrigées), cellules de « Points à traiter » où coller les 4 formules (Excel 365), non vérifié (rendu Excel).
