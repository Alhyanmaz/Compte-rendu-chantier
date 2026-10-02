---
name: compte-rendu-chantier-v3
description: "CR de réunion de chantier (maîtrise d'œuvre) au nouveau format — colonne N°, HISTORIQUE, statuts dans FAIT LE — à partir du CR précédent annoté par José en réunion (notes de cellule Excel) et de la transcription audio de la réunion. Utiliser ce skill dès que José (ACAU) fournit un classeur CR avec des notes de cellule et/ou une transcription de réunion de chantier, demande le CR de la semaine (CRC-NN), ou renvoie un CR corrigé « à remouliner » avant l'édition finale, ou demande l'ordre du jour de la prochaine réunion de chantier — même s'il ne dit pas « skill » ni « V3 ». Remplace les skills compte-rendu-chantier (V1) et compte-rendu-chantier-v2 pour les classeurs au nouveau format (HONGUEMARE, DUCLAIR « LE MIT », BENOUVILLE pôle socio-culturel) ; convertit aussi un classeur à l'ancien format (sans colonne N°) quand José le demande. Ne couvre pas les courriers ni les CR de conception."
---

# CR de chantier — V3 : notes de réunion + audio, relecture dans Excel

Utilisateur : José Mazzarese, architecte (ACAU), maîtrise d'œuvre. Un **classeur vivant** par opération,
mis à jour chaque semaine ; le PDF CRC-NN remis aux entreprises en est l'impression. Le CR devient
contractuel à défaut de réserve sous 8 jours : une erreur de routage part chez la mauvaise entreprise.
C'est pourquoi les scripts font toute l'écriture, et Claude fait la rédaction et signale ses doutes.

## Le cycle, en deux passages

1. **Intégration** — entrées : le CR diffusé la semaine précédente, **annoté** par José pendant la réunion
   (notes de cellule), et la **transcription** de l'enregistrement. Sortie : le CR « à relire »
   (rouge du jour, colonne ROUTAGE, fonds roses, onglet « Non routé »).
2. **Remoulinage** — entrées : la version livrée **et** la version corrigée par José dans Excel.
   Sortie : le CR final (fonds roses retirés, ROUTAGE et « Non routé » vidés) **et l'ordre du jour de la
   réunion suivante** (Word).

Repérer le passage demandé : un classeur avec des notes + une transcription → intégration ; deux versions
du même CR (« livrée » / « corrigée ») → remoulinage. Si c'est ambigu, demander.

## Étape 0 — Préparer l'environnement

- Scripts : `scripts/` de ce skill. Les copier dans un répertoire de travail et y travailler sur des
  **copies** des fichiers reçus (jamais sur l'original).
- Dépendances, polices Denim INK, LibreOffice : voir `references/environnement.md`. Les polices sont sous
  licence : si José ne les a pas fournies, continuer (hauteurs estimées) et le lui dire.
- Ne jamais enregistrer le classeur avec openpyxl (il réécrit tout : textes, logos, boutons, impression)
  — openpyxl sert en **lecture seule** ; toute écriture passe par les scripts (édition XML chirurgicale).

## Intégration

### 1. Lire les notes

```bash
python scripts/lire_notes.py CR_ANNOTE.xlsx notes.json
```

Chaque note avec ce qu'elle désigne : `maj <N°>` (ligne existante, texte et dates actuels fournis) ou
`nouvelle(s) : <code>, <section>`. Lire aussi la Page de garde (A22 = CRC précédent, B22 = date) :
le nouveau CR est CRC + 1, à la date de la réunion (la demander si elle n'est pas donnée ; ne jamais
prendre la date du jour de traitement).

### 2. Lire la transcription et rédiger

Règles de rédaction et de routage : `references/redaction.md` (à lire à chaque fois). En bref :
- chaque note est **reformulée à l'appui de l'audio** ; la note fait foi sur la ligne, le statut,
  l'échéance ; l'audio apporte les précisions ;
- les mots-clés des notes (`fin`, `PM`, `urgent`, `retard`, `+7`, `en attente X`, `?`…) deviennent
  `fait_le` / `pour_le` / `doute`, jamais du texte (`references/memo-notes.md`) ;
- un point net de l'audio **sans note** → opération `"source": "transcription"` (fond rose) ;
- un choix à confirmer → `doute` (question fermée, fond rose soutenu) ;
- le reste de l'audio (rien à tracer, ou impossible à attribuer) → `non_route` avec la raison.

Écrire `OPERATIONS.json` (format : `references/operations-json.md`), en remplissant `routage` pour
**chaque** opération : note d'origine, extrait de l'audio utilisé, choix faits. José relit le CR avec
cette colonne sous les yeux : c'est elle qui remplace l'ancien tableau de routage.

### 3. Intégrer, ajuster, contrôler

```bash
python scripts/integrer_cr.py CR_ANNOTE.xlsx OPERATIONS.json etape1.xlsx rapport.json
python scripts/ajuster_hauteurs.py etape1.xlsx HONGUEMARE_CR_ACAU_CRC-NN.xlsx
```

Le script applique seul : page de garde, suppression des notes (il **refuse** de continuer si une note
n'est citée nulle part), masquage des lignes soldées au passage précédent, mises à jour et lignes neuves
(N° attribués, insertions de lignes si la réserve manque), relances automatiques, texte condensé,
couleurs, Retard, ROUTAGE, fonds roses, onglet « Non routé », sauts de page. Les règles qu'il applique :
`references/regles.md`.

Puis les contrôles de `references/environnement.md` (archive, notes restantes, recalcul, PDF).

### 4. Livrer

Envoyer le classeur à José, et dans la réponse :
- chiffres du rapport : mises à jour, nouvelles lignes (N°), relances automatiques, lignes masquées,
  insertions ;
- la liste des **[?]** avec leur question, et le nombre de points audio seuls (roses) ;
- le nombre de passages non routés (détail dans l'onglet) ;
- les **interprétations** faites et les angles morts (ex. hauteurs estimées sans LibreOffice) ;
- rappel : corriger dans Excel, puis renvoyer **le fichier corrigé et cette version livrée**.

## Remoulinage (après la relecture de José)

```bash
python scripts/relire_cr.py CR_LIVRE.xlsx CR_CORRIGE.xlsx etape1.xlsx rapport.json
python scripts/ajuster_hauteurs.py etape1.xlsx HONGUEMARE_CR_ACAU_CRC-NN.xlsx
```

Le script compare les deux versions ligne par ligne (par N°) et ne retouche que ce que José a changé :
ajout daté du jour dans HISTORIQUE → rouge et texte recalculé ; dates et statuts modifiés → rouge ;
ligne de réserve remplie sans N° → N° attribué ; puis retire les fonds roses, vide ROUTAGE, vide et
masque « Non routé ». Aucun masquage, aucune relance.

Lire les **alertes** du rapport et les transmettre à José, en particulier :
- « OBSERVATIONS modifié sans HISTORIQUE » : texte conservé, mais à reporter dans HISTORIQUE, sinon
  perdu à la prochaine mise à jour de la ligne ;
- [?] laissé tel quel (considéré comme validé) ; N° en double ; ligne supprimée ;
- écriture dans la dernière ligne d'un tableau (trait épais de fin de tableau).

Si le rapport liste des modifications que José n'a pas faites (Excel réécrit tout le fichier à
l'enregistrement), ne pas livrer : examiner et le lui dire.

### Ordre du jour de la réunion suivante (livré avec le CR final)

```bash
python scripts/ordre_du_jour.py HONGUEMARE_CR_ACAU_CRC-NN.xlsx ODJ_reunion_NN+1.docx [--date AAAA-MM-JJ] [--lieu "…"]
```

Toujours sur le CR **final** (après remoulinage) : l'ordre du jour reprend ses statuts et échéances.
Date et heure : Page de garde (date du CR + 7 jours, heure B23) ; si José annonce un autre jour, `--date`.
Contenu : approbation du CR ; points prioritaires (URGENT, Retard) ; points à traiter par intervenant
(échéance atteinte le jour de la réunion, ou Relance sans échéance) avec leur dernier état ; points en
attente d'un tiers (une ligne par organisme) ; visite ; questions diverses ; présence requise et convoqués.
Rien n'est rédigé à la main : si José veut ajouter un point, il le dit, et on l'ajoute dans le Word livré
(ou au CR, s'il s'agit d'une observation). Signaler dans la réponse le nombre de points par rubrique et la
date retenue. Le Word utilise la police Denim INK : conseiller à José de l'envoyer en PDF.

## Conversion d'un classeur à l'ancien format

Si le classeur n'a pas de colonne N° (onglets d'observations sans en-tête « N° »), il faut d'abord le convertir :
méthode et contrôles dans `references/migration.md` (une fois par opération, copie de test validée par José).

## Ce qui reste à José (ne pas le faire à sa place)

- Masquer les doublons (repérés en bleu), trancher les [?], valider le CR avant diffusion.
- Coller les formules de l'onglet « Points à traiter » (Excel 365) si elles manquent — elles ne sont pas
  évaluables dans l'environnement.
- Contrôler l'aperçu avant impression dans Excel.

## Interdits

- Inventer une date, une échéance, un statut, un montant ou un nom absent des notes, de l'audio ou du
  classeur ; « trancher au mieux » un doute en silence.
- Écrire dans Coordonnees, Généralités, Modele_lot, Reportage photo ; renommer un onglet.
- Poser Retard ou URGENT qui ne vient pas d'une note de José (entendu seulement dans l'audio → [?]).
- Relancer une ligne PM ; déclarer un doublon automatiquement.
- Modifier le classeur autrement que par les scripts ; livrer sans les contrôles.
