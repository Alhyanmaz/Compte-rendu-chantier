# Optimisation du CR de chantier pour intégrer directement les transcriptions

**Opération** : HONGUEMARE, groupe scolaire (extension et réhabilitation)
**Base analysée** : `HONGUEMARE_CR_ACAU - test opti.xlsx`, état CRC-15, réunion du 22/09/2026
**Objectif** : supprimer le tableau de routage. Il reste deux relectures : la transcription corrigée, puis le CR final.

---

## 0. Sources et angles morts

**Ce que j'ai lu :**

- le classeur fourni, en entier et par script : les 18 onglets, les formules, les couleurs du texte enrichi, les lignes masquées, les tableaux Excel et les notes ;
- les skills `compte-rendu-chantier` (V1) et `compte-rendu-chantier-v2`, dans l'état où ils sont synchronisés dans cette session. La V2 contient les décisions du 23/09/2026 : la chaîne en 4 temps, la « relecture allégée » et les préférences de routage apprises.

**Ce que je n'ai pas pu lire (angle mort principal) :**

- **le projet claude.ai « HONGUEMARE » et la discussion où vous avez déjà répondu à des questions.** Ce projet n'est pas accessible depuis cette session. Je ne connais vos réponses que si elles ont été reportées dans les skills. Si une proposition ci-dessous contredit une réponse déjà donnée, c'est votre réponse qui compte.
- l'outil de transcription audio que vous utilisez. Je ne sais pas s'il identifie les locuteurs ni s'il accepte un vocabulaire personnalisé. Cela change beaucoup la fiabilité de l'attribution « qui doit faire » (voir § 3.2).
- le CCTP et les limites de prestation entre lots. Elles ne figurent pas dans le classeur.

**Convention** : ce qui est une **interprétation** de ma part est marqué *(interprétation)*. Tout le reste a été mesuré sur le fichier.

---

## 1. Constat mesuré sur le CRC-15

| Mesure | Valeur | Méthode |
|---|---|---|
| Observations visibles (hors en-têtes) | 211 | onglets MOE-MOA, Concessionnaires, 01 à 11 |
| Cellules avec 5 ajouts « Au JJ/MM » ou plus | 53 sur 211 | décompte par regex |
| Longueur d'une observation | médiane 179 caractères, max 982 | idem |
| Lignes créées le 22/09/2026 | 26 | ABORDÉ LE = 22/09/2026 |
| Ajouts de fond sur des lignes existantes | 36 | texte rouge autre que « Relance » |
| Relances automatiques seules | 38 | texte rouge = « Au 22/09/2026 Relance » |
| Libellés différents dans FAIT LE | 32 (lignes masquées comprises) | ex. `En attente`, `En attnte mise au point chaufferie`, `en attente enedis`, `RELANCE`, `DOUBLON`, `doublon`, `AODEX 20`, `Retard AXL`… |
| Tableaux Excel dans MOE-MOA | 15, aux noms générés (`TableauMOEMOA_411812131415161439`) | `ws.tables` |

Une réunion produit donc **environ 60 points de fond** (26 lignes créées et 36 ajouts). Ce sont ces points que vous relisez. Les 38 relances sont mécaniques.

### Ce que ces chiffres montrent pour l'automatisation

1. **Une observation n'a pas de référence stable.** Pour la retrouver, il faut soit son numéro de ligne Excel, qui bouge dès qu'on insère une ligne, soit une comparaison de texte. **C'est la raison d'être du tableau de routage** : il sert à valider une correspondance que le classeur ne permet pas d'établir seul.
2. **La même information est écrite à plusieurs endroits.** Dans le seul CRC-15 :
   - « portes asservies ouvertes » : MOE-MOA L7 et 06 L32 ;
   - « SRPN / citerneau DN 32 / siège (27) » : texte identique en MOE-MOA L10 et L35, puis Concessionnaires L38 ;
   - « visite BEVELEC le 01/10/2026 » : MOE-MOA L20 et L23 ;
   - « Voir CR de la réunion chaufferie » : MOE-MOA L25 et L36 ;
   - « goujons en quinconce » : 02 L62 et 03 L36 ;
   - « GOUJON perfore son support » : 04 L15 et L29.

   Chaque doublon oblige à choisir entre deux lignes, ce qui fait du routage « Douteux ». Et une mise à jour oubliée sur l'une des deux lignes donne un CR contradictoire.
3. **FAIT LE mélange trois choses** : une date de clôture, un statut (Relance, URGENT) et un commentaire libre (« Retard AXL », « En attente 22/09/2026 »). Une machine ne peut pas l'interpréter sans deviner.
4. **Des marques d'incertitude ont fini dans le CR diffusé** : « Le RDV concessionnaire ne semble plus d'actualité ? » (MOE-MOA L67), « prévoir une visite dans un showroom pour choix ? » (11 L17), le mot-clé « URGENT » collé au texte (03 L9 « …officiellement.URGENT »), et un ajout daté « Au 16/09/2026 » dans le CR du 22/09 (03 L8).

---

## 2. Le principe de la nouvelle chaîne, et où je ne suis pas d'accord

### Nouvelle chaîne proposée (2 relectures au lieu de 3)

```
Transcription ──► [Claude] RÉSUMÉ ROUTÉ ──► vous corrigez ──► [Claude] intégration directe ──► vous relisez le CR
                  (déjà rédigé au format CR,                   (aucune reformulation,
                   chaque point porte sa référence)             copie du texte validé)
```

La règle qui rend cela possible : **le texte que vous relisez dans le résumé est, mot pour mot, celui qui sera écrit dans le CR.** À l'intégration, il n'y a plus de reformulation, donc plus rien à revalider entre les deux.

### Mon désaccord : supprimer le tableau ne supprime pas la décision de routage

Le routage ne disparaît pas, il **déménage dans le résumé**. Sinon je range les points sans contrôle. Vous ne découvririez alors les erreurs qu'à la relecture du CR final, c'est-à-dire au pire moment. Sur 211 lignes, une erreur de placement est plus difficile à voir que dans une liste de 60 points. Et le CR devient opposable faute de réserve sous 8 jours.

À titre d'ordre de grandeur, la V2 enregistre, le 23/09/2026 sur DUCLAIR CRC-43, **13 items « Certain » corrigés sur 35**. Avec un tel taux, le placement ne peut pas se passer de validation.

La suppression du tableau n'est donc sûre qu'à **deux conditions** :

1. chaque point du résumé affiche **où il va** : la référence de la ligne existante, ou « NOUVEAU » avec l'onglet et la sous-section ;
2. **aucun point marqué `[?]` n'est intégré**. Il reste dans le résumé tant que vous ne l'avez pas tranché. Cela règle aussi le cas des « ? » passés dans le CRC-15.

### Ce que je retire de la V2 (la « relecture allégée » du 23/09)

- **Le journal des modifications devient inutile** si vous relisez le CR final de toute façon. Le texte rouge est déjà le journal du jour. Une seule chose n'est pas visible dans le CR : **les lignes masquées ce jour**. Je les listerai dans ma réponse finale, avec leur référence.
- **Le mode « Certains directs »** n'a plus d'objet : tout passe par le résumé.

---

## 3. Modifications proposées du classeur

### Priorité 1 : indispensables pour supprimer le routage

#### 3.1 Une référence unique par observation

- **Nouvelle colonne à gauche, `N°`** (largeur ≈ 8), avant OBSERVATIONS.
- **Format** : `[code de la section]-[numéro sur 3 chiffres]`, attribué à la création et **jamais réutilisé ni renuméroté**. Il ne dépend pas de ÉTUDES/TRAVAUX, pour qu'une référence reste stable si un point change de sous-section.
  - Lots : `02-014`, `05-027`…
  - MOE-MOA, un code par intervenant : `MOA-`, `AMO-`, `CT-`, `SPS-`, `MOE-`, `ECO-`, `BET.ELE-`, `BET.CVC-`, `BET.THE-`, `BET.VRD-`, `BET.CUI-`, `BET.SBE-`, `BET.SBO-`, `BET.ACO-`, `BET.AMI-`.
  - Concessionnaires : `CON.RES-`, `CON.EAU-`, `CON.ELE-`, `CON.ASS-`.
- **Numérotation initiale** : dans l'ordre actuel des lignes, **y compris les lignes masquées**, pour qu'une ligne réaffichée garde sa place.
- **Bénéfice hors automatisation** : les entreprises peuvent répondre « point 02-031 » par mail, et vous pouvez dicter « sur le 02-031 : relance » en réunion. Si ce réflexe s'installe, le routage d'une bonne partie des points devient trivial *(interprétation : suppose que l'habitude prenne)*.

> **Alternative écartée** : mettre la référence en colonne E, à droite. Aucune formule ne bougerait et les skills garderaient A–D. Mais la référence serait loin de l'œil, et hors de la zone lue en premier dans le PDF. Comme les skills doivent être réécrits de toute façon, je recommande la colonne de gauche. À vous de trancher (question Q1).

#### 3.2 Un onglet `Référentiel` (non imprimé)

Le tableau Coordonnees est conçu pour être imprimé, pas pour la machine. Je propose de le laisser tel quel et d'ajouter un onglet de travail :

| Code | Onglet | Section | Organisme | Représentants | **Noms entendus (alias)** |
|---|---|---|---|---|---|
| 02 | 02 - GROS ŒUVRE | — | AXL | Chloé RECHER, Sylvain JIROT | *à compléter : ex. « Axel », « AXL Construction »* |
| 09 | 09 - CVP | — | ELAIRGIE CAEN | Eric GARDINIER | *ex. « Élargie », « Elairgie »* |
| MOA | MOE-MOA | Maître d'ouvrage | Mairie de HONGUEMARE | David TAURIN, Arnaud LEBAS, Loïc DRIEU | *« la mairie », « la commune »* |
| AMO | MOE-MOA | AMO | CICLOP | Nicolas GAUDIN | *« Ciclop », « Cyclope »* |
| … | | | | | |

La colonne **« Noms entendus »** est le point clé pour les transcriptions. Elle permet de relier une déformation phonétique à la bonne entreprise sans rapprochement « par ressemblance », ce que la V2 interdit à juste titre. Dans le même onglet, deux listes courtes :

- **zones et ouvrages du chantier** : bâtiment A, préau, chaufferie, restaurant scolaire, école existante, extension, mairie, SHED, cours anglaises, vide sanitaire… ;
- **lexique propre à l'opération** : RSD, longrines, hourdis, courettes, acodrains, MOB, tebopin…

Les alias réels des représentants, c'est vous qui les connaissez : **je ne les invente pas**. Les exemples en italique du tableau sont des hypothèses à confirmer.

Si votre outil de transcription accepte un vocabulaire personnalisé, ces deux listes peuvent aussi lui être fournies *(angle mort : je ne connais pas votre outil)*.

#### 3.3 FAIT LE en liste fermée

La colonne FAIT LE ne contient plus que **deux types de valeur** :

- une **date** : le point est soldé ;
- un **statut**, choisi dans une liste déroulante (validation de données) :
  `Relance` · `URGENT` · `En cours` · `En attente MOA` · `En attente MOE` · `En attente AMO` · `En attente CT` · `En attente SPS` · `En attente BET` · `En attente concess.` · `En attente [n° lot]` · `Annulé` · `Doublon` · `Sans objet` · `Refusé`.

Le **« de quoi »** de l'attente (« mise au point chaufferie », « phase 2 », « devis BEVELEC ») va dans le texte, sous la forme « Au JJ/MM/AAAA en attente de … ».

`PM` n'est plus admis dans FAIT LE : il est aujourd'hui présent 9 fois, alors qu'il relève de POUR LE.

La **migration** des 32 libellés existants vers cette liste se fait par une table de correspondance, que je vous soumets **une seule fois** avant application. Par exemple `En attnte mise au point chaufferie` → `En attente MOE` + ajout texte ? Je ne le sais pas : c'est une question pour vous.

#### 3.4 Une information, une ligne, un responsable

Règle de rédaction : un sujet est suivi **sur une seule ligne**, celle de **celui qui doit agir**. Ailleurs, on écrit un **renvoi** court (`cf. CON.ELE-003`), et on ne recopie pas le texte.

Pour l'existant, je propose de traiter au cas par cas les six doublons listés au § 1. Exemple : le suivi ENEDIS est porté par CON.ELE, et un renvoi est posé en AMO et en MOE. Cela se décide une fois, avec vous.

### Priorité 2 : fiabilisation

| # | Modification | Pourquoi |
|---|---|---|
| 3.5 | **Gris et jaune par mise en forme conditionnelle** sur chaque tableau : fond jaune si FAIT LE = URGENT ; police grise si POUR LE = PM (sauf FAIT LE « En attente… » ou « En cours ») ou si FAIT LE est une date ou un statut terminal. | Supprime deux étapes manuelles fragiles, le grisage PM et le jaune URGENT (y compris son retrait), qui ont donné lieu à plusieurs corrections d'incident dans la V1. **Angle mort** : je n'ai pas vérifié si une couleur conditionnelle l'emporte sur les couleurs du texte enrichi (le dernier ajout rouge). C'est à tester dans votre Excel avant de généraliser. Le rouge « ajout du jour » reste manuel dans tous les cas. |
| 3.6 | **Page de garde en formules** : C22 = titre calculé depuis B22 ; B26 = « RDV chantier » + B22+7 + heure ; A22 reste saisi. | Il ne reste plus qu'une ou deux cellules à changer chaque semaine. Aujourd'hui B22 et B26 sont saisies en dur. |
| 3.7 | **Noms harmonisés** : onglets au format `NN - LIBELLÉ` (`06 - MENSUIERIE INT.` → `06 - MENUISERIES INT.`, `02 -GROS OEUVRE` → `02 - GROS ŒUVRE`, `09 -CVP` → `09 - CVP`…) ; titres de section MOE-MOA alignés sur Coordonnees ; tableaux Excel renommés par code (`T_MOE`, `T_BET_AMI`, `T_02`…). | Les skills interdisent aujourd'hui de renommer, parce que le nom sert de clé. Avec des références stables, cette contrainte tombe. **Aucune formule inter-onglets** (vérifié dans le XML des 18 feuilles : 216 formules, toutes internes à leur feuille). Seuls 4 noms définis pointent vers la page de garde. |
| 3.8 | **Réserve de lignes vides** : au moins 5 lignes vides en fin de chaque sous-section. | Évite l'insertion de ligne dans le XML (`insert_row`), l'opération la plus risquée des skills. |

### Anomalies à corriger lors de la migration (vérifiées)

- **Page de garde A6 : ce n'est pas une anomalie, contrairement à ce qu'indique le skill V1.** Le XML montre `<c r="A6" t="e" vm="1">` et des parties `xl/richData/`. C'est une **image placée dans la cellule** (fonction Excel 365). `#VALUE!` n'est que la valeur de repli que lisent openpyxl ou LibreOffice. *Angle mort : je ne l'ai pas vu dans Excel. À confirmer : voyez-vous bien le visuel ?* Conséquence pour les skills : le patch doit conserver l'attribut `vm` et les parties `richData` / `metadata.xml`.
- **MOE-MOA** : section « BET cuisine_CREACEPT » **absente de Coordonnees** ; SIEGE 27 (2e maître d'ouvrage dans Coordonnees) n'a **pas de section** ; section Économiste vide ; **note de cellule résiduelle sur A35** (« José Mazzarese: » sans texte) ; L35 sans ABORDÉ LE.
- **POUR LE antérieur à ABORDÉ LE** : MOE-MOA L67 (08/06 < 30/06), MOE-MOA L94 (masquée), 02 L13 et 03 L9 (01/08 < 25/08).
- **Dates suspectes** : 02 L83 ABORDÉ LE = 07/07/**2027** ; 01 L18 et 02 L30 POUR LE = 01/05/2027 *(à confirmer : peut être voulu)*.
- **« 19/052026 »** (barre manquante) dans 27 cellules (onglets 01, 02, 03, 04). Sans conséquence à l'écran, mais invisible pour toute recherche « Au JJ/MM/AAAA ».
- **02 F84** : un « s » isolé hors tableau.
- **Coordonnees** : « DESAMIANRAGE », « BET AOUSTIQUE », légende « Esc » en H35 contre « Exc » en H22.
- Le skill V1 mentionne un bloc **SYNTHESE** en 03 et 04 : **il n'existe plus** dans ce fichier. Toutes les feuilles de lot ont seulement ÉTUDES / TRAVAUX.

### Priorité 3 : ce que je déconseille (pour l'instant)

- **Compacter l'historique des relances** (« Relancé les 23/06, 30/06, 07/07… ») : **déconseillé.** Les Généralités prévoient une pénalité de 300 €/jour pour retard de remise de documents. La trace datée de chaque relance, notifiée dans un CR non contesté, est votre pièce justificative. La lisibilité en souffre, mais c'est le prix de la preuve. Et cela violerait la règle « on complète, on ne réécrit jamais ».
- **Remplir les présences (C/P/Abs/Exc) à partir de la transcription** : **déconseillé.** Un participant silencieux n'apparaît pas dans l'audio. Or le compteur d'absences sert aux pénalités (Généralités : « Absence en réunion : 200 €/ jours calendaires »). Il reste en saisie manuelle.
- **Registre unique et génération complète du classeur** : une seule table de données (une ligne par ajout daté), à partir de laquelle Claude régénère chaque semaine les onglets imprimables, couleurs comprises. C'est l'architecture la plus robuste pour l'automatisation, puisque les couleurs deviennent calculées et qu'il n'y a plus de patch XML. Mais c'est une refonte complète : **à ne pas faire en cours de chantier** (HONGUEMARE en est à 4,5 mois sur 20). Une piste pour une prochaine opération, dès la première réunion *(interprétation : votre calendrier d'opérations ne m'est pas connu)*.

---

## 4. Le format du « résumé routé »

C'est le document que vous corrigez. Il remplace à la fois le résumé actuel et le tableau de routage. Il est organisé **comme le CR** : onglet, puis sous-section, puis point.

### Syntaxe d'un point (une ligne = un point)

```
[réf. existante | NOUVEAU]  Texte exact qui sera écrit  ‖ POUR LE : +N / JJ/MM / PM / —  ‖ FAIT LE : statut / date / —
```

- **Déplacer un point** vers un autre titre = le re-router (couper/coller, rien d'autre).
- **Supprimer un point** = il n'est pas intégré.
- **`[?] question`** en fin de ligne = bloquant, pas d'intégration tant que ce n'est pas tranché.
- En fin de document : trois blocs **pour information** :
  - « Relances automatiques » : une liste de références, rien à relire ;
  - « Clôtures / masquages prévus » : liste de références ;
  - « Propositions connexes » : non appliquées par défaut ; écrire « OK » pour les appliquer.

### Exemple reconstitué à partir du CRC-15 (réunion du 22/09/2026)

*Le texte est le texte rouge réellement présent dans le fichier. Les références sont remplacées par le **numéro de ligne actuel** (ex. `05·L8`), puisque la numérotation n'existe pas encore.*

> **05 – MENUISERIES EXT. – AVA**
> *ÉTUDES*
> - `[05·L8]` Pas de coffre, uniquement le rouleau. ‖ POUR LE : — ‖ FAIT LE : —
> - `[NOUVEAU]` SHED : prévoir une remontée à 90° du précadre pour permettre la fixation de la couvertine de GOUJON VALLEE. ‖ POUR LE : +14
> - `[NOUVEAU]` Fournir le nuancier de toile des stores. ‖ POUR LE : +14
> - `[NOUVEAU]` Préciser au CSPS la méthodologie d'intervention pour la pose des SHED. ACAU évoque un platelage et un garde-corps pris sur la charpente. ‖ POUR LE : +7
> - `[05·L21]` ACAU transfère le retour de la MOA ce jour.
>
> *TRAVAUX*
> - `[05·L33]` AVA évoque un changement de sous-traitant. `[?] faut-il demander la notification au CSPS et à la MOA, comme 02·L68 ?`
>
> **11 – VRD – CFB TP**
> - `[11·L17]` Prévoir une visite dans un showroom pour choix. `[?] qui organise, et pour quelle date ?`
>
> **Propositions connexes** (non appliquées sans « OK »)
> - `[03·L36]` Ajout « Au 22/09/2026 cf. 02·L62 : AXL demande une fixation des goujons en quinconce. », au lieu de recopier le texte. Type 4, interface GO/charpente.
>
> **Relances automatiques** (information) : 02·L31, L33, L34, L35, L39, L41, L47, L60, L61, L68, L74 ; 06·L8, L15 à L19, L21 à L24, L26, L27 ; …

Remarque : dans le CRC-15, les points 05·L33 et 11·L17 sont passés dans le CR avec un « ? » ou sans suite claire. Dans ce format, ils seraient restés bloqués dans le résumé *(interprétation : je ne sais pas si ces « ? » étaient volontaires)*.

### Support du résumé

Je propose de garder un **.docx**, qui est déjà votre support de correction selon la V2. Un titre par onglet et par sous-section (styles Titre 1 / Titre 2), une puce par point. La relecture se fait à l'écran ou imprimée, et les corrections directement dans le texte. La correction dans la conversation reste possible pour les petites réunions.

---

## 5. Conséquences sur les skills (à faire dans un second temps)

À réécrire en une **V3 unique**. V1 et V2 seraient archivées : leur coexistence explique déjà deux arbitrages contradictoires (§ « Arbitrages du 18/09/2026 » de la V2).

| Règle actuelle | Devient |
|---|---|
| Colonnes A–D (OBSERVATIONS, ABORDÉ LE, POUR LE, FAIT LE) | A = N°, B–E décalées (si Q1 = gauche) |
| Recherche de la ligne par comparaison de texte (V2 § 3.3) | Recherche par référence ; comparaison de texte seulement pour les points `NOUVEAU` (détection de doublon) |
| Tableau de routage Excel (V2 étape 4) | **Supprimé** : remplacé par le résumé routé (§ 4) |
| Journal des modifications | **Supprimé** : remplacé par la liste des lignes masquées dans la réponse finale |
| Statuts FAIT LE libres, comparés « sans casse ni espaces » | Liste fermée (§ 3.3) ; toute valeur hors liste est signalée |
| Grisage PM et jaune URGENT écrits à la main | Mise en forme conditionnelle, si le test § 3.5 est concluant |
| « Ne jamais renommer un onglet » | Onglets et tableaux identifiés par leur code (Référentiel) |
| Correspondance des noms via Coordonnees « sans ressemblance » | Idem, **plus** la colonne alias du Référentiel |
| Création d'une ligne neuve | Attribution de la référence suivante de la section (compteur dans le Référentiel) |

Ce qui **ne change pas** : on complète sans jamais réécrire, le gabarit « Au JJ/MM/AAAA », le rouge pour l'ajout du jour, le patch XML chirurgical (jamais `openpyxl.save()` sur le classeur réel), la date de réunion comme date d'écriture, et les propositions connexes jamais appliquées sans « OK ».

---

## 6. Migration

1. **Entre deux réunions**, sur une **copie** du classeur. Pas la veille d'une réunion.
2. Je prépare la version migrée : N°, Référentiel, liste FAIT LE, mise en forme conditionnelle, formules de la page de garde, renommages, réserve de lignes. Avec elle, je fournis **la table de correspondance des 32 statuts** et **la liste des anomalies**, pour que vous les tranchiez.
3. Vous l'ouvrez dans Excel. Vérifications : **aucun message de récupération**, impression PDF identique au CRC-15 à la colonne N° près, test de la mise en forme conditionnelle sur une ligne rouge.
4. Premier CR en nouveau format. **Pendant les deux ou trois premières réunions, je garde le verbatim de la transcription en regard de chaque point du résumé.** Cette précaution de la V2 me semble toujours justifiée.

---

## 7. Questions à trancher (réponses fermées)

| # | Question | Options |
|---|---|---|
| Q1 | Où placer la colonne N° ? | **Gauche (recommandé)** / droite (E) |
| Q2 | Format de référence : `02-014` pour les lots et `MOE-003` / `BET.CVC-002` pour les intervenants, convient-il ? | Oui / autre format |
| Q3 | La référence doit-elle apparaître dans le PDF diffusé aux entreprises ? | **Oui (recommandé)** / non (colonne masquée à l'impression) |
| Q4 | Liste fermée des statuts FAIT LE (§ 3.3) : à valider, compléter ou réduire ? | Validée / modifications |
| Q5 | Doublons du § 1 : qui porte chaque sujet (ex. ENEDIS : CON.ELE, AMO ou MOE) ? | À trancher sujet par sujet |
| Q6 | Mise en forme conditionnelle (§ 3.5) : acceptez-vous de la tester sur la copie ? | Oui / non |
| Q7 | Renommage des onglets (§ 3.7) : l'accepter, sachant que les anciens PDF garderont les anciens noms ? | Oui / non |
| Q8 | Support du résumé routé : .docx, ou correction dans la conversation ? | .docx / conversation / les deux selon la taille |
| Q9 | Faut-il ajouter les sections manquantes (SIEGE 27 en MOA) et corriger CREACEPT (absent de Coordonnees) ? | Oui / non / à voir |
