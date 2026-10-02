# Migration au nouveau format — BENOUVILLE, pôle socio-culturel, CRC-16 du 24/09/2026

Copie convertie le 02/10/2026 avec `tools/migration_cr.py` (réglages : `tools/operations/benouville_psc.py`), puis `tools/ajuster_hauteurs.py` (hauteurs et sauts de page). Le classeur d'origine n'est pas modifié. Les fichiers Excel ne sont pas versionnés (coordonnées des intervenants).

## Ce qui a été appliqué

Les **règles générales** décidées pour HONGUEMARE (`docs/decisions-HONGUEMARE.md`), comme pour DUCLAIR :

- colonne **N°** (code + numéro, jamais réutilisé), **HISTORIQUE** (texte intégral, non imprimé), texte imprimé **condensé** (392 observations) ;
- **statuts dans FAIT LE** (liste fermée) ; « PM » passé de POUR LE à FAIT LE (95 lignes) ;
- couleurs écrites dans les cellules ; mise en forme conditionnelle limitée aux fonds jaune (URGENT) et orange (Retard) ;
- lignes de réserve masquées (76 lignes vides masquées), bordures normalisées ;
- Référentiel et Points à traiter (masqués), zones d'impression, titres jamais seuls en bas de page.

Résultat : **274 observations visibles**, 118 masquées. Recalcul : aucune erreur de formule, sauf `Page de garde!A4` (#VALUE!, image du logo, sans effet dans Excel).

## Question : le nombre de pages (à trancher par José)

| Version | Pages (rendu LibreOffice) |
|---|---|
| CR d'origine | 42 |
| **Livrée** : saut de page avant chaque TRAVAUX (règle d'HONGUEMARE) | **59** |
| Variante : TRAVAUX à la suite d'ÉTUDES (titre toujours gardé avec sa 1re ligne) | 47 |

Pour HONGUEMARE, vous aviez gardé le saut en connaissance de cause (≈ 48 pages contre 43). Ici, le surcoût est de **17 pages**. Il vient des 17 lots : la plupart ont une sous-section TRAVAUX d'une à trois observations, qui occupe donc seule une page.

Les deux fichiers sont livrés. Si vous retenez la variante pour BENOUVILLE, il faudra l'inscrire dans les réglages de l'opération, pour qu'elle s'applique aussi aux CR suivants. Aujourd'hui, elle ne s'obtient que par `CR_SAUT_TRAVAUX=0`.

## Interprétations à valider

### 1. Codes des sections de MOE-MOA

| Section | Code | Section | Code |
|---|---|---|---|
| Maitre d'ouvrage_VILLE DE BENOUVILLE | MOA | BET électricité_SSI_BIELEC | BET.ELE |
| Bureau de contrôle_VERITAS | CT | BET CVC_AREHA | BET.CVC |
| CSPS_VERITAS | SPS | BET HQE_EXEO | **BET.HQE** |
| Maitrise d'oeuvre_ACAU | MOE | BET structure_KUBE | BET.STR |
| Economiste_AECO | ECO | BET VRD_AHMES | BET.VRD |
| | | BET accoustique_ORFEA | BET.ACO |
| | | Paysagiste_VERT LATITUDE | PAY |

Lots : le N° prend le numéro du lot (01-001…).

### 2. Statuts des lignes visibles (40 lignes normalisées)

| Avant | Après | Lignes |
|---|---|---|
| « En attente visa ACAU » (casse variable) | « En attente ACAU » | 20 |
| « En attente retour ACAU », « synthèse ACAU » | « En attente ACAU » | 2 (02-006, 03-014) |
| « En attente validation MOE » | « En attente ACAU » (**interprétation** : MOE = ACAU) | 1 (10-031) |
| « En attente retour MOA », « validation MOA » | « En attente MOA » | 2 (BET.CVC-005, 16-024) |
| « en attente SCF » | « En attente SCF » | 2 |
| « En attente HARETDECO » | « En attente HARET DECO » | 1 (13-004) |
| « En attente prototype » | « En attente » + « → En attente : prototype » dans le texte | 1 (13-005) |
| « En attente confirmation parquet MOA » | « En attente MOA » + « → En attente : confirmation parquet » dans le texte | 1 (12-004) |
| « DOUBLON », « doublon » | « Doublon » | 5 |
| « RELANCE », « relance » | « Relance » | 5 |

### 3. Valeurs laissées telles quelles (à corriger par vous)

| N° | Colonne | Valeur | Ligne |
|---|---|---|---|
| **BET.CVC-006** | FAIT LE | « à réaliser » | visible : ce n'est pas un statut de la liste. « Relance » ? « En cours » ? |
| **03-018** | FAIT LE | « 25/29/2026 » | visible : date impossible (mois 29). Quelle date ? |
| **02-011**, **08-002** | POUR LE | « CF tableau » | visibles : aucune échéance lisible, donc pas de Retard ni de relance automatique sur ces lignes |
| 05-001, 10-001, 13-001 | POUR LE | « CF tableau » | masquées |
| 02-029 | POUR LE | « SEPTMEBRE » | masquée |
| BET.CVC-002, BET.CVC-003 | FAIT LE | « Obsolete » | masquées |

### 4. Sujet daté d'un autre jour qu'ABORDÉ LE (date laissée dans le texte)

| N° | Constat |
|---|---|
| MOA-024, MOA-025 | sujet « Au 16/07/2026 », ABORDÉ LE 15/07/2026 |
| 03-007 | sujet daté du 26/02/2026, ABORDÉ LE 02/02/2026 |
| 03-017 | sujet daté du 30/07/2026, ABORDÉ LE 03/09/2026 |

### 5. Anomalies signalées, non corrigées

| N° | Constat |
|---|---|
| MOA-005 | ABORDÉ LE vide (POUR LE = ABORDÉ LE + 21 : sans date, pas d'échéance) |
| 04-004 | ABORDÉ LE vide |
| 15-010 | POUR LE 27/02/2026 antérieur à ABORDÉ LE 18/06/2026 |

## Corrections faites (et listées)

| Où | Correction |
|---|---|
| 10-022, 10-023, 10-024 | FAIT LE saisi en texte « 16/07/2026 » → vraie date |
| Coordonnees H42, A19, A39, A49 | « Esc » → « Exc » (Excusé) |
| 02 - GROS OEUVRE, note « 7mm » | La note de cellule posée sur l'observation de **02-033** est restée sur l'observation (B39). Elle n'a pas glissé sur le N°. Elle sera demandée au prochain CR : toute note doit être traitée. |

## Remarques

- **Page de garde** : le titre « Compte rendu de la réunion de chantier n°[16] » devient une formule sur le numéro de CR (A21) : il affichera « n°[17] » au CR suivant (vérifié par un essai). La prochaine réunion reste en **texte libre**, comme dans l'original.
- **Points à traiter** : les 4 formules (`docs/rapport-migration-test.md`) se collent en **A28, F28, K28 et P28**. L'onglet est à remasquer ensuite.
- **Doublons** (même destinataire, même objet) : aucun parmi les lignes ouvertes. Les paires de textes ressemblants examinées portent sur des objets différents. Aucun repère bleu.
- Non modifiés, à vérifier par vous :
  - le titre de l'onglet 13 indique « VRD - PEINTURES » (déjà dans l'original) ;
  - dans Coordonnees, « ISOLPLAF » et « ISOPLAF » coexistent, de même que « SEEL LEAUGEOIS » et « LAUGEOIS ».
- Les pages 5 et 6 ne contiennent que l'en-tête : c'est la partie droite de l'onglet Coordonnees, déjà imprimée ainsi dans l'original.

## Essai du cycle hebdomadaire (CRC-17 fictif, non livré)

Sur la copie convertie, `integrer_cr.py` et `ajuster_hauteurs.py` donnent :
- 2 mises à jour ;
- 2 lignes neuves : BET.HQE-003, et 14-008 dans un tableau TRAVAUX ;
- les relances automatiques ;
- page de garde CRC-17, titre « n°[17] ».

`ordre_du_jour.py` lit aussi le classeur : heure 10h30 et lieu « Visioconférence », repris de la page de garde.

## Non vérifié

- Le rendu dans **Excel** (contrôles faits avec LibreOffice) : ouverture sans message de réparation, aperçu avant impression.
- Hauteurs mesurées avec LibreOffice et les polices Denim INK. Deux passages successifs ont donné deux hauteurs différentes (75 et 61,8 pt) pour MOA-029 et 02-002. Le texte tient sur 4 lignes dans les deux cas.
