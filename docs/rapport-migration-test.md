# Rapport de migration — copie de test du CR HONGUEMARE

- **Fichier produit** : `HONGUEMARE_CR_ACAU_TEST_MIGRATION.xlsx`, construit à partir du CRC-15 du 22/09/2026.
- **Script** : `tools/migration_cr.py`. Il édite le XML de façon chirurgicale et n'enregistre jamais le classeur avec openpyxl.
- **Décisions appliquées** : `docs/decisions-HONGUEMARE.md`.

> Il s'agit d'une **copie de test**. Le classeur maître n'a pas été touché. Rien n'est validé tant que vous ne l'avez pas ouvert dans Excel (voir § 7).

---

## 1. Contrôles faits de mon côté

| Contrôle | Résultat |
|---|---|
| Archive et 100 % des parties XML relisibles | OK |
| Onglets **Généralités** et **Photos**, image de la page de garde (`richData`), dessins | **identiques à l'octet près** |
| En-têtes des 30 tableaux conformes à leurs colonnes (sinon Excel « répare » le fichier) | OK, 0 écart |
| Recalcul complet par LibreOffice | 309 formules, **1 seule erreur** : `#VALUE!` en Page de garde A6. C'est l'image dans la cellule, déjà présente dans l'original (voir proposition, § anomalies). |
| Comptage de l'onglet « Points à traiter » | 403 observations (404 − 1 CREACEPT), 26 URGENT : identique à l'original |
| Rendu PDF (LibreOffice) | La colonne HISTORIQUE n'est pas imprimée. N°, texte condensé, rouge du jour, gris et jaune sont conformes. |

**Angles morts :**
- **Largeur de page.** Dans LibreOffice, l'original déborde déjà sur une 2e page en largeur, car la police *Denim INK* n'y est pas installée. LibreOffice ne permet donc pas de juger la largeur. J'ai gardé **exactement la même largeur totale** que l'original : N° 8 + OBSERVATIONS 45,2 = 53,2 de l'ancienne colonne OBSERVATIONS. Vérifiez l'aperçu avant impression dans Excel.
- **Hauteurs de ligne** : elles sont estimées, puisque la police est absente ici. J'ai été généreux, mais une ligne peut rester trop basse ou trop haute. C'est à ajuster dans Excel si besoin.
- **Nombre de pages** : 77 → 104 dans LibreOffice. L'augmentation vient surtout des sauts de page avant chaque TRAVAUX, qui est votre décision.

---

## 2. Ce qui a changé dans la structure

**Sur chaque onglet d'observations** (MOE-MOA, Concessionnaires, 01 à 11, Modele_lot) :
- Colonne **A = N°**, en police 7. Les autres colonnes sont décalées : B = OBSERVATIONS, C = ABORDÉ LE, D = POUR LE, E = FAIT LE.
- Colonne **F = HISTORIQUE** : le texte intégral, avec ses couleurs d'origine. Elle est **hors zone d'impression**.
- Les formules POUR LE ont été décalées (`=C7+7`) et les valeurs sont inchangées.
- **Liste déroulante** sur FAIT LE (nom `ListeStatuts`). La saisie libre reste possible pour les dates.
- **Zone d'impression** définie de A à E.

**Sur les lots 01 à 11 uniquement :**
- **saut de page avant chaque TRAVAUX** ;
- *(initiative à valider)* **ligne d'en-tête répétée** en haut de chaque page. Sans elle, une page TRAVAUX commencerait sans titres de colonnes.

**Numérotation** : 403 observations numérotées dans l'ordre actuel, **lignes masquées comprises**. Exemples : `MOA-001` à `MOA-021`, `02-001` à `02-075`, `CON.ASS-003`. Le prochain numéro de chaque section est dans le Référentiel.

**Sections MOE-MOA :**
- **Supprimée** : « BET cuisine_CREACEPT », anciennes lignes 115 à 120. Elle ne contenait qu'une observation : « Réaliser mission VISA dès réception des plans entreprises. »
- **Créée** : « Maitre d'ouvrage_SIEGE 27 », sous la section MOA, avec un tableau vide de 2 lignes (code `SIE`).
- **Conservée vide** : Économiste.

**Réserve de lignes** : des lignes vides ont été ajoutées là où une sous-section en avait moins de 2. Cela représente 29 ajouts (1 ou 2 lignes chacun), dont le détail est dans `report.json`.

**Page de garde** : C22 (titre) et B26 (prochaine réunion) sont désormais **calculés depuis B22**. Chaque semaine, il ne reste qu'A22 (n° CRC) et B22 (date) à changer.
- La date est écrite sans la fonction TEXTE, parce que celle-ci dépend de la langue d'Excel.
- Petite différence : B26 affiche « 9H00 », la valeur de B23, au lieu de « 09H00 ».

---

## 3. Texte condensé (135 observations)

Format appliqué, par exemple :

```
fournir plan de coffrage .
→ Au 15/07/2026 : Relance - 25/08/26 RELANCE.
→ Au 15/09/2026 : AXL s'engage a fournir la MAJ des plans EXE pour le 29/09 et plan préau pour le 06/10.
→ Relancé 6 fois, dernière le 22/09/2026
```

**Règles :**
- On garde le constat d'origine et les **2 dernières mises à jour**, suivis de « Relancé N fois, dernière le … ».
- On ne condense que s'il y a plus de 2 mises à jour, ou au moins 2 relances seules. Une ligne courte reste telle quelle.
- Les couleurs suivent la source : rouge pour l'ajout du jour, gris pour une ligne PM ou soldée.

**Limites connues :**
- « Relance » compte les segments « Au JJ/MM/AAAA Relance / URGENT / Rappel » **seuls**. Une relance écrite sans « Au » (par exemple « - 25/08/26 RELANCE ») reste dans le texte et n'est pas comptée.
- Il arrive que le texte d'origine commence déjà par « Au JJ/MM/AAAA ». C'est alors ce premier segment qui sert de constat.

---

## 4. Statuts FAIT LE et PM

**PM** : 87 lignes sont passées de POUR LE « PM » à FAIT LE « PM ». Sur 27 autres lignes, le « PM » a été retiré de POUR LE mais FAIT LE a été conservé : une date de clôture, un « Doublon », un « En attente… » ou un « PM » déjà présent.

**Libellés normalisés : 26 lignes.** Les cas qui demandent votre décision sont en gras.

| N° | Avant | Après | Remarque |
|---|---|---|---|
| MOA-004, AMO-004 | En attente du retour du siege | En attente SIEGE 27 | |
| **MOA-012** | En attente phase 2 | En attente | **« phase 2 » perdu : à remettre dans le texte ?** |
| MOA-016, 03-014, 09-017, 11-017 | doublon / DOUBLON | Doublon | casse |
| MOA-017 | En attente devis bevelec | En attente BEVELEC | |
| **MOA-019, AMO-005** | En attente 22/09/2026 | En attente | **l'objet de l'attente n'était pas précisé** |
| MOE-008, CON.RES-002, 10-006 | En attente (RDV) concesionnaire | En attente concessionnaire | |
| CON.ELE-001 | En attente SIEGE | En attente SIEGE 27 | |
| 02-008, 11-002, 11-027 | En attente retour MOA | En attente MOA | |
| **02-011** | En attnte mise au point chaufferie | En attente | **objet perdu : à remettre dans le texte ?** |
| **02-014** | Retard AXL | Relance | **interprétation** |
| 02-024 | RELANCE | Relance | casse |
| **02-052** | En attente MAJ process | En attente | **objet perdu** |
| **08-006** | En attente BAT | En attente | **« BAT » : bon à tirer ? organisme ?** |
| 09-011 | En attente visa ACAU sur DT | En attente ACAU | |
| 10-003 | en attente enedis | En attente ENEDIS | |
| **11-004** | En attente retour HAMES | En attente AHMES | **interprétation : HAMES = AHMES (BET VRD)** |
| 11-006 | En attente CT | En attente DEKRA | |
| **04-020** (masquée) | AODEX 20 | *inchangé* | **valeur hors liste, laissée telle quelle** |

---

## 5. Anomalies

**Corrigées (35) :**
- **« 19/052026 » → « 19/05/2026 »** dans 27 observations (onglets 01, 02, 03 et 04). HISTORIQUE est corrigé lui aussi.
- **« 08/09/20256 »** (08-005) et **« 15/07/20265 »** (02-021) → 2026 *(interprétation)*.
- **02-075** : ABORDÉ LE 07/07/**2027** → 07/07/2026 *(interprétation : le texte parle du 15/07/2026)*.
- 02 : le « s » isolé en F84 est supprimé. MOE-MOA A35 : la note vide est supprimée.
- **Coordonnees** : DESAMIANRAGE → DESAMIANTAGE, AOUSTIQUE → ACOUSTIQUE, « Esc » → « Exc » (en-tête H35 et légende A41).

**Signalées, non corrigées :**
- **POUR LE antérieur à ABORDÉ LE** : MOE-008 (08/06 < 30/06), BET.ELE-001 (masquée), 02-007 et 03-003 (01/08 < 25/08).
- **ABORDÉ LE vide** : AMO-004, 03-030 (visible), 02-047 et 05-009 (masquées).
- **POUR LE en 2027** : 01-001 à 01-009 et 02-024. C'est probablement voulu (phase 2 en avril/mai 2027), à confirmer.

---

## 6. Doublons surlignés en bleu (8 lignes, 4 paires)

Le bleu est appliqué sur N° et OBSERVATIONS seulement, pour que le jaune URGENT reste visible sur les autres colonnes. **C'est vous qui masquez.**

| Paire | Lignes |
|---|---|
| RDV ENEDIS (texte identique) | AMO-004 et MOE-008 |
| Plan et méthodologie de maintien des existants (texte identique) | 01-004 et 01-005 |
| Portails et portillons en contrôle d'accès | 05-015 et 10-024 |
| Portes de circulation asservies ouvertes | MOA-001 et 06-026 |

**Correction de ma proposition initiale.** Trois des « doublons » que j'y listais ne sont **pas** des doublons de ligne. Il s'agit de deux sujets différents qui ont reçu la même phrase d'ajout :
- MOA-014 et MOA-017 (BEVELEC) ;
- MOA-019 et AMO-005 (« Voir CR réunion chaufferie ») ;
- 04-009 et 04-023 (couvertine).

Je ne les ai pas surlignés. De même, 02-056 et 03-030 (goujons) sont une **interface** entre deux lots, pas un doublon.

---

## 7. À tester dans Excel (dans cet ordre)

1. **Ouverture sans message de récupération.** Si Excel propose de réparer : arrêtez-vous et envoyez-moi le message exact.
2. **Onglet « Test MFC »** (visible, à supprimer ensuite). La colonne F dit ce que vous devez voir. Question clé : **le rouge du jour reste-t-il rouge sur une ligne grise ?** (lignes T-01 et T-02).
3. **Aperçu avant impression** d'un lot, par exemple 02. Vérifiez que les 5 colonnes tiennent sur la largeur, que HISTORIQUE est absente, que TRAVAUX commence en haut de page et qu'aucun texte n'est coupé.
4. **Liste déroulante** dans FAIT LE, et saisie d'une date dans la même colonne.
5. **Onglet « Points à traiter »** (masqué : clic droit sur un onglet > Afficher) : collez les 3 formules ci-dessous.

### Formules à coller (Excel 365 en français)

Collez chaque formule dans la cellule indiquée, sur l'onglet « Points à traiter ». Le nom `TousLesPoints` est déjà défini dans le classeur : il assemble les 30 tableaux.

**Bloc 1 URGENT (cellule A23)**
```
=LET(d;TousLesPoints;n;CHOISIRCOLS(d;1)&"";f;CHOISIRCOLS(d;5)&"";x;ASSEMB.H(CHOISIRCOLS(d;1);GAUCHE(CHOISIRCOLS(d;2);120);CHOISIRCOLS(d;4);CHOISIRCOLS(d;5));FILTRE(x;(n<>"")*(n<>"0")*(f="URGENT");"—"))
```

**Bloc 2 ÉCHÉANCE DÉPASSÉE (cellule F23)**
```
=LET(d;TousLesPoints;n;CHOISIRCOLS(d;1)&"";p;CHOISIRCOLS(d;4);f;CHOISIRCOLS(d;5)&"";x;ASSEMB.H(CHOISIRCOLS(d;1);GAUCHE(CHOISIRCOLS(d;2);120);p;CHOISIRCOLS(d;5));c;(n<>"")*(n<>"0")*ESTNUM(p)*(p>0)*(p<dateCR)*((f="")+(f="0")+(f="Relance")+(GAUCHE(f;10)="En attente"));SIERREUR(TRIER(FILTRE(x;c);3;1);"—"))
```

**Bloc 3 EN ATTENTE (cellule K23)**
```
=LET(d;TousLesPoints;n;CHOISIRCOLS(d;1)&"";f;CHOISIRCOLS(d;5)&"";o;SIERREUR(TEXTE.APRES(f;"En attente ");"");x;ASSEMB.H(CHOISIRCOLS(d;1);GAUCHE(CHOISIRCOLS(d;2);120);CHOISIRCOLS(d;4);CHOISIRCOLS(d;5));c;(n<>"")*(n<>"0")*(GAUCHE(f;10)="En attente");SIERREUR(TRIERPAR(FILTRE(x;c);FILTRE(o;c);1);"—"))
```

**Angles morts sur ces formules :**
- Je n'ai pas pu les exécuter : ces fonctions n'existent que dans Excel 365.
- Les noms français (ASSEMB.H, CHOISIRCOLS, TRIERPAR, TEXTE.APRES) sont ceux que je connais, mais je ne peux pas les vérifier ici.
- Les cellules vides deviennent « 0 » une fois assemblées, d'où les tests `<>"0"`.

**En cas d'erreur :**
- `#NOM?` sur `TousLesPoints` : recréez le nom (Formules > Gestionnaire de noms > Nouveau, nom `TousLesPoints`) avec :
  ```
  =ASSEMB.V(TableauMOEMOA_1;TableauMOEMOA_SIE;TableauMOEMOA_2;TableauMOEMOA_3;TableauMOEMOA_4;TableauMOEMOA_5;TableauMOEMOA_411;TableauMOEMOA_4118;TableauMOEMOA_411812;TableauMOEMOA_41181213;TableauMOEMOA_411812131435;TableauMOEMOA_41181213141516;TableauMOEMOA_4118121314151614;TableauMOEMOA_4118121314151617;TableauMOEMOA_411812131415161439;TableauLot3233;TableauLot323334;TableauLot323336;TableauLot32333630;TableauLot18;TableauLot19;TableauLot21;TableauLot22;TableauLot23;TableauLot24;TableauLot25;TableauLot26;TableauLot27;TableauLot28;TableauLot29)
  ```
- `#PROPAGATION!` : une cellule sous le bloc n'est pas vide. Videz-la.

**Maintenance.** Toute section créée ou supprimée doit être ajoutée ou retirée dans `TousLesPoints`. Sinon elle disparaît de la synthèse **sans message d'erreur**. Le comptage du haut de l'onglet porte sur toutes les lignes, masquées comprises.

---

## 8. Nouveaux onglets

| Onglet | État | Contenu |
|---|---|---|
| Référentiel | masqué | Intervenants (code, onglet, organisme, représentants sans e-mail, **alias vides à compléter**, prochain N°) ; liste des 39 statuts ; zones ; lexique ; sigles à confusion. Tous les termes ont été vérifiés comme présents dans le classeur. |
| Points à traiter | masqué | Comptage par onglet (formules classiques, vérifiées) et 3 blocs à compléter avec les formules ci-dessus. |
| Test MFC | visible | Test de la mise en forme conditionnelle. À supprimer après le test. |
