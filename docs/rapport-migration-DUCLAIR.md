# Migration au nouveau format — DUCLAIR « LE MIT », CRC-44 du 30/09/2026

Copie convertie le 02/10/2026 avec `tools/migration_cr.py` (réglages : `tools/operations/duclair_mit.py`), puis `tools/ajuster_hauteurs.py` (hauteurs et sauts de page). Le classeur d'origine n'est pas modifié. Les fichiers Excel ne sont pas versionnés (coordonnées des intervenants).

## Ce qui a été appliqué

Les **règles générales** décidées pour HONGUEMARE (`docs/decisions-HONGUEMARE.md`), sans décision propre à DUCLAIR :

- colonne **N°** à gauche (code + numéro, jamais réutilisé), **HISTORIQUE** (texte intégral, non imprimé), texte imprimé **condensé** (règle A, « [...] », « Relancé N fois… (sans réponse depuis N j) ») ;
- **statuts dans FAIT LE** (liste fermée, liste déroulante) ; « PM » passé de POUR LE à FAIT LE (112 lignes ; 76 autres lignes ont gardé leur date ou leur statut de FAIT LE) ;
- couleurs **écrites dans les cellules** (PM / soldé en gris, ajout du CR 44 en rouge, statut en rouge, PM en gris) ; mise en forme conditionnelle limitée aux **fonds** jaune (URGENT) et orange (Retard) ;
- 2 lignes de réserve par sous-section (masquées), bordures du classeur, aucune bordure sur les bords extérieurs, traits épais limités aux en-têtes ;
- page de garde en formules (titre et prochaine réunion), Référentiel et Points à traiter (masqués), zones d'impression, en-tête répété, titres jamais seuls en bas de page.

Résultat : **177 observations visibles**, 562 masquées (archivées), 31 tableaux ; **48 pages** au rendu LibreOffice, contre 43 pour le CR d'origine (sauts de page avant TRAVAUX et titres gardés avec leur tableau). Aucune erreur de formule au recalcul (1 368 formules).

## Interprétations à valider

### 1. Codes des sections de MOE-MOA

Choisis par Claude d'après les titres des sections (le N° imprimé en dépend) :

| Section | Code | Section | Code |
|---|---|---|---|
| Maitre d'ouvrage_VILLE DE DUCLAIR | MOA | BET CVC_OCEADE | BET.CVC |
| Maitre d'ouvrage_EPFN | **EPF** | BET ENERGETIQUE ENVIRONNEMENTALE_OCEADE | **BET.ENV** |
| AMO_CICLOP | AMO | BET structure_ESGCB | **BET.STR** |
| Bureau de contrôle_BUREAU VERITAS | CT | BET Signalétique_ATELIERS 59 | **BET.SIG** |
| CSPS_PRESENTS | SPS | BET VRD_AHMES | BET.VRD |
| Maitrise d'œuvre_OPC_ACAU | MOE | BET accoustique_GAMBA | BET.ACO |
| Economiste_ECLA | ECO | BET Audit-Déchet_CEDN | **BET.DEC** |
| BET CFO_CFA_OCEADE | BET.ELE | Paysagiste_ATELIER 2 PAYSAGE | **PAY** |
| | | TELECOM_KYNTUS | **CON.TEL** |

Lots : le N° prend le numéro du lot (01-001…). Le lot 01 regroupe MBTP et AMAGYS (désamiantage) dans le même onglet, donc sous le même code 01.

### 2. Statuts des lignes visibles

- « En attente retour MOA » → « En attente MOA » ; « En attente retour HARLIN » → « En attente HARLIN ».
- « En attente CT » → « **En attente VERITAS** » (03-037), comme « CT » → DEKRA à HONGUEMARE.
- Attentes d'un état d'avancement plutôt que d'un intervenant : « En attente » et objet remis dans le texte (« → En attente : hors d'eau ») — 01-044 (hors d'eau), 07-017 (hors d'air), 11-001 (synthèse trappe), 11-024 (PH R+1).
- **04-031** « Abandon MOA » → « **Annulé** », mention « → Abandon MOA » dans le texte. Conséquence : la ligne sera masquée au CR suivant.
- **01-050** « Avant enduit ext. Septembre » (POUR LE « PM ») → « **PM** », mention « → Avant enduit ext. septembre » dans le texte.

### 3. Lignes masquées : rien d'interprété

Les lignes masquées sont des points archivés. Leurs valeurs atypiques (« RETARD MBTP », « archivée », « FAIT »…) sont **laissées telles quelles** : les convertir en « Retard » ferait démarrer des compteurs de retard sur des points clos, et les listerait dans « Points à traiter ». Liste en annexe.

### 4. Lignes sans texte (lot 04)

10 lignes **masquées** n'ont pas de texte mais gardent des dates : 04-004, 04-006 à 04-009, 04-013 à 04-017. Elles ont reçu un N° et restent masquées. Faut-il les garder ou les supprimer ?

## Corrections faites (et listées)

| Où | Correction |
|---|---|
| AMO-004, 01-002 | Date du texte « 24/11/20255 », « 10/12/20255 » → **2025** (lecture la plus proche de la date du CR ; la règle d'HONGUEMARE imposait 2026 : corrigée le 02/10/2026) |
| 01-128, 11-061 | FAIT LE saisi en texte « O3/03/2026 » (lettre O), « 27/08/2026 » → vraies dates |
| 14 - COUVERTURE, B14 | Caractère « s » isolé dans ABORDÉ LE d'une ligne vide, supprimé |
| Coordonnees H46, A26, A44, A50 | « Esc » → « Exc » (Excusé) |

## Remarques

- **Page de garde** : la prochaine réunion devient une formule (date du CR + 7 jours et heure). Elle affiche « 07/10/2026   9h00 » au lieu de « 07/10/2026   09h00 », car l'heure est reprise de la cellule B22 (« 9h00 »).
- Le titre de la page de garde s'affiche « LE MIT DUCLAIRRéhabilitation… » : c'est déjà le cas dans le CR d'origine, et cette cellule n'a pas été modifiée.
- **Points à traiter** : les 4 formules du rapport d'HONGUEMARE (`docs/rapport-migration-test.md`) se collent ici en **A25, F25, K25 et P25** (et non A23…). L'onglet est à remasquer ensuite.
- **Doublons** (même destinataire, même objet) : aucun doublon probable trouvé parmi les lignes ouvertes (comparaison des sujets, seuil de ressemblance de 75 %). Aucun repère bleu.
- Le CR est rédigé par Vincent Bodelle (page de garde) : les règles et le skill V3 s'appliquent de la même façon.

## Non vérifié

- Le rendu dans **Excel** (contrôles faits avec LibreOffice) : ouverture sans message de réparation, aperçu avant impression.
- Les hauteurs de ligne sont mesurées avec les polices Denim INK installées dans l'environnement de travail ; la marge est celle validée pour HONGUEMARE.

## Annexes

### Lignes masquées : valeurs de FAIT LE conservées telles quelles (23)

| N° | Onglet | FAIT LE |
|---|---|---|
| 01-010 | 01 - GROS OEUVRE | « Coulage planning ind 0 au 06/03/2026. » |
| 01-024 | 01 - GROS OEUVRE | « A regler sur prototype » |
| 01-026 | 01 - GROS OEUVRE | « RETARD LOGI HABITAT » |
| 01-037 | 01 - GROS OEUVRE | « Retard MBTP » |
| 01-075 | 01 - GROS OEUVRE | « FAIT » |
| 01-080 | 01 - GROS OEUVRE | « Retard MBTP » |
| 01-115 | 01 - GROS OEUVRE | « Retard MBTP » |
| 01-125 | 01 - GROS OEUVRE | « Retard MBTP » |
| 02-007 | 02 - CHARPENTE BARDAGE BOIS | « URGENT. RETARD POULINGUE » |
| 02-046 | 02 - CHARPENTE BARDAGE BOIS | « RETARD POULINGUE » |
| 02-047 | 02 - CHARPENTE BARDAGE BOIS | « traitement non necessaire » |
| 02-055 | 02 - CHARPENTE BARDAGE BOIS | « RETARD 3 semaines MBTP » |
| 02-061 | 02 - CHARPENTE BARDAGE BOIS | « RETARD POULINGUE » |
| 03-005 | 03 - COUVERTURE | « archivée » |
| 03-015 | 03 - COUVERTURE | « archivée » |
| 06-014 | 06 - MEN.INT | « Archivé » |
| 06-024 | 06 - MEN.INT | « Archivée » |
| 06-035 | 06 - MEN.INT | « RETARD BTH » |
| 11-002 | 11 -  CVC | « Archivée » |
| 11-015 | 11 -  CVC | « FAIT » |
| 12-018 | 12 - ELECTRICITE | « RETARD DGS » |
| 12-042 | 12 - ELECTRICITE | « RETARD DGS » |
| 12-045 | 12 - ELECTRICITE | « RETARD DGS » |

### Sujet daté d'un autre jour qu'ABORDÉ LE (19) : date laissée dans le texte

| N° | Onglet | Constat |
|---|---|---|
| MOA-003 | MOE-MOA | sujet daté du 21/11/2025, ABORDÉ LE 26/11/2025 : date laissée dans le sujet |
| EPF-002 | MOE-MOA | sujet daté du 10/12/2025, ABORDÉ LE 03/12/2025 : date laissée dans le sujet |
| MOE-017 | MOE-MOA | sujet daté du 16/04/2026, ABORDÉ LE 15/04/2026 : date laissée dans le sujet |
| 01-105 | 01 - GROS OEUVRE | sujet daté du 22/01/2026, ABORDÉ LE 21/01/2026 : date laissée dans le sujet |
| 01-108 | 01 - GROS OEUVRE | sujet daté du 22/01/2026, ABORDÉ LE 21/01/2026 : date laissée dans le sujet |
| 01-110 | 01 - GROS OEUVRE | sujet daté du 29/07/2026, ABORDÉ LE vide : date laissée dans le sujet |
| 02-056 | 02 - CHARPENTE BARDAGE BOIS | sujet daté du 18/02/2026, ABORDÉ LE 18/02/2025 : date laissée dans le sujet |
| 02-067 | 02 - CHARPENTE BARDAGE BOIS | sujet daté du 11/05/2026, ABORDÉ LE 06/05/2026 : date laissée dans le sujet |
| 03-022 | 03 - COUVERTURE | sujet daté du 05/11/2025, ABORDÉ LE 29/10/2025 : date laissée dans le sujet |
| 03-023 | 03 - COUVERTURE | sujet daté du 05/11/2025, ABORDÉ LE 29/10/2025 : date laissée dans le sujet |
| 03-049 | 03 - COUVERTURE | sujet daté du 01/07/2025, ABORDÉ LE 01/07/2026 : date laissée dans le sujet |
| 03-050 | 03 - COUVERTURE | sujet daté du 01/07/2025, ABORDÉ LE 01/07/2026 : date laissée dans le sujet |
| 03-051 | 03 - COUVERTURE | sujet daté du 01/07/2025, ABORDÉ LE 01/07/2026 : date laissée dans le sujet |
| 03-053 | 03 - COUVERTURE | sujet daté du 15/07/2026, ABORDÉ LE vide : date laissée dans le sujet |
| 09-008 | 09 - ASCENSEUR | sujet daté du 26/03/2026, ABORDÉ LE 25/03/2026 : date laissée dans le sujet |
| 11-008 | 11 -  CVC | sujet daté du 14/01/2025, ABORDÉ LE 14/01/2026 : date laissée dans le sujet |
| 11-037 | 11 -  CVC | sujet daté du 26/08/2026, ABORDÉ LE 28/08/2026 : date laissée dans le sujet |
| 11-046 | 11 -  CVC | sujet daté du 02/12/2025, ABORDÉ LE vide : date laissée dans le sujet |
| 12-049 | 12 - ELECTRICITE | sujet daté du 25/03/2026, ABORDÉ LE 01/04/2026 : date laissée dans le sujet |

### Anomalies signalées, non corrigées (12)

| N° | Onglet | Constat |
|---|---|---|
| 01-002 | 01 - GROS OEUVRE | POUR LE 05/01/2025 antérieur à ABORDÉ LE 15/10/2025 |
| 01-110 | 01 - GROS OEUVRE | ABORDÉ LE vide (ligne masquée) |
| 01-125 | 01 - GROS OEUVRE | ABORDÉ LE vide (ligne masquée) |
| 01-151 | 01 - GROS OEUVRE | POUR LE 14/01/1900 antérieur à ABORDÉ LE 17/06/2026 |
| 02-005 | 02 - CHARPENTE BARDAGE BOIS | POUR LE en 2055 : 17/12/2055 (à confirmer) |
| 02-008 | 02 - CHARPENTE BARDAGE BOIS | POUR LE 17/11/2025 antérieur à ABORDÉ LE 19/11/2025 |
| 02-027 | 02 - CHARPENTE BARDAGE BOIS | POUR LE 03/01/1900 antérieur à ABORDÉ LE 11/03/2026 |
| 03-046 | 03 - COUVERTURE | ABORDÉ LE vide (ligne masquée) |
| 03-053 | 03 - COUVERTURE | ABORDÉ LE vide (ligne masquée) |
| 04-031 | 04 - MEN.EXT | ABORDÉ LE vide |
| 06-017 | 06 - MEN.INT | POUR LE 07/01/2025 antérieur à ABORDÉ LE 17/12/2025 |
| 14-005 | 14 - COUVERTURE | ABORDÉ LE vide |

### Statuts normalisés (39 lignes)

| Avant | Après | Lignes |
|---|---|---|
| « DOUBLON » | « Doublon » | 14 |
| « RELANCE » | « Relance » | 6 |
| « En attente retour MOA » | « En attente MOA » | 3 |
| « annulé » | « Annulé » | 3 |
| « En attente retour HARLIN » | « En attente HARLIN » | 2 |
| « doublon » | « Doublon » | 2 |
| « Sans Objet » | « Sans objet » | 2 |
| « En attente hors d'eau » | « En attente » | 1 |
| « Avant enduit ext. Septembre » | « PM » | 1 |
| « En attente CT » | « En attente VERITAS » | 1 |
| « Abandon MOA » | « Annulé » | 1 |
| « En attente Hors d'air » | « En attente » | 1 |
| « En attente Synthèse trappe » | « En attente » | 1 |
| « En attente PH R+1 » | « En attente » | 1 |