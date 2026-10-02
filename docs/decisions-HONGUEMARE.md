# Décisions — optimisation du CR HONGUEMARE

Registre des arbitrages de José Mazzarese. Il fait foi sur `proposition-optimisation-CR.md` en cas d'écart.

## Décisions du 29/09/2026 (session précédente, d'après son README)

- Colonne « NB REL. » : **refusée**.
- Regroupement des lignes « mission VISA » dans Généralités : **refusé**.
- Doublons : **traitement manuel uniquement**, jamais de déclaration « Doublon » automatique (ex. ENEDIS MOE-MOA l.35).
- Échéances POUR LE antérieures à ABORDÉ LE : **signaler, ne pas corriger**.
- Modifications en bleu : pour la relecture uniquement, à retirer dans la version finale.

## Décisions du 29/09/2026 (cette session)

| Sujet | Décision |
|---|---|
| Statuts | **Liste fermée dans FAIT LE** (date ou statut). Pas de colonnes STATUT / ATTENTE DE. |
| « En attente » | **En attente + organisme** (ENEDIS, SIEGE 27, BEVELEC…), liste générée depuis le Référentiel. |
| Historique | **Texte condensé** : constat d'origine + **2 dernières mises à jour** + « Relancé N fois, dernière le … ». L'intégral va dans une colonne HISTORIQUE, **exclue de la zone d'impression**. Exception assumée à la règle « on complète, on ne réécrit jamais ». **Pas de validation finale avant de l'avoir vu en réel.** |
| Colonne N° | **À gauche**, format `02-014` / `MOA-014`, jamais réutilisé. **Imprimée** dans le PDF. |
| Alias | Onglet `Référentiel` **non imprimé**. Alias laissés vides au départ et enrichis au fil des réunions, à partir des réponses aux `[?]`. |
| Mise en forme conditionnelle | Accord de principe réservé : **à tester sur une copie réelle** avant toute adoption. |
| Renommage des onglets | **Non.** |
| « 19/052026 » | **Correction automatique** en « 19/05/2026 ». |
| Autres anomalies (02 L83 en 2027, 02 F84 « s », note vide MOE-MOA A35, fautes de Coordonnees) | **Corriger et lister.** |
| Économiste | **Garder** la section, même vide. |
| Doublons existants | **Surligner en bleu** (remplissage) ; masquage manuel par José. |
| Nouveaux ajouts | **Pas de bleu** : ils sont déjà validés dans l'Excel de résumé routé. |
| PM | **Devient un statut de FAIT LE** : ligne conservée et grisée, jamais masquée. POUR LE peut garder une vraie date, ou rester vide. Un statut actif (Relance, URGENT, En attente) remplace PM. |
| Résumé routé | **Excel, une seule feuille**, ordonnée comme le CR ; les points `[?]` sont bloquants. |
| Sous-sections | Saut de page avant chaque ÉTUDES / TRAVAUX, recalculé à chaque CR (validé en connaissance du surcroît de pages). |
| CREACEPT (BET cuisine) | **Supprimer** les lignes de la section MOE-MOA ; suppression signalée dans le rapport. |
| SIEGE 27 | **Créer** une section dans MOE-MOA, code `SIE-`, placée après celle du maître d'ouvrage. |
| Page de garde | **Formules** (titre et prochaine réunion calculés depuis la date du CR). |
| Lignes vides | **2 lignes vides** en fin de sous-section. Ajout de lignes autorisé si besoin. |
| Mise en page | Chaque sous-section (ÉTUDES / TRAVAUX) **commence en haut d'une page**. |
| Relance des lignes PM | **Jamais**, même si POUR LE porte une date dépassée. |
| Structure de toutes les observations | **Toutes** les observations suivent la même structure, qu'elles aient un historique ou non : **sujet initial en gras** ; puis « → Au JJ/MM/AAAA : … » pour les 2 dernières remarques ; puis « → Relancé N fois, dernière le … ». Les lignes suivantes ne sont pas en gras et gardent les couleurs de la source. |
| Sujet qui commence par « Au JJ/MM/AAAA » | Date **retirée** du sujet si elle est égale à ABORDÉ LE ; sinon **laissée et soumise à José** (02-066, 03-035, 05-009 au 29/09/2026). |
| Relances sans « Au JJ/MM/AAAA » (ex. « - 25/08/26 RELANCE ») | **Supprimées** de l'affichage (restent dans HISTORIQUE). Ne se produira plus quand Claude rédigera 100 % des CR. |
| ABORDÉ LE corrigés | **02-066** : 26/05/2026 → 02/06/2026 ; **05-009** : vide → 20/05/2026 ; **03-035** : 15/09/2026 → 16/06/2026 (demande à AGC issue de 02-065) ; la date est retirée du sujet (décision du 29/09/2026). |
| Relances dans le texte condensé | **Règle A stricte** : seules les relances **postérieures à la dernière remarque** sont affichées (« Relancé N fois, dernière le … ») ; le compteur repart de 0 à chaque remarque. Les relances antérieures ne restent que dans HISTORIQUE. Les mises à jour masquées sont signalées par **« [...] »**. Choix confirmé par José en connaissance de cause (MOA-004, MOA-012, MOA-014 n'affichent plus leurs relances anciennes). |
| Onglet « Points à traiter » | Repris de la discussion claude.ai (§ 8), **adapté** à la liste fermée dans FAIT LE : voir ci-dessous. |

## Onglet « Points à traiter » (Excel 365)

Récapitulatif fourni par José (discussion claude.ai), adapté par Claude à la décision « statuts dans FAIT LE » : il n'y a pas de colonnes STATUT ni ATTENTE DE.

- Formules uniquement (VSTACK + FILTRE + TRIER + LET) sur les tableaux de tous les onglets d'observations, avec dateCR = 'Page de garde'!B22.
- **Bloc 1 URGENT** : FAIT LE = « URGENT ».
- **Bloc 2 ÉCHÉANCE DÉPASSÉE** : POUR LE < dateCR et FAIT LE vide, « Relance » ou commençant par « En attente », trié par POUR LE croissant. PM est exclu.
- **Bloc 3 EN ATTENTE** : FAIT LE commençant par « En attente », trié par organisme (texte après « En attente »).
- Colonnes : N° | GAUCHE(OBSERVATIONS;120) | POUR LE | FAIT LE.
- « Aucun résultat » : 3e argument de FILTRE = "—". Blocs côte à côte ou très espacés (#PROPAGATION!).
- En haut : tableau de comptage par onglet (NB.SI / SOMMEPROD).
- Onglet interne, exclu de l'impression et du PDF.
- À tester dans Excel 365 : ces fonctions ne sont pas évaluables dans l'environnement de Claude.

## Tests Excel (§ 7 du rapport) — validés par José le 29/09/2026

Ouverture sans réparation, onglet « Test MFC », aperçu avant impression, liste déroulante FAIT LE, 3 formules de « Points à traiter » collées par José : **OK**.

## Décisions du 29/09/2026 (suite)

| Sujet | Décision |
|---|---|
| Sujet initial en gras | **Annulé** : mise en forme d'origine conservée. |
| Mise en forme conditionnelle | **Généralisée** à tous les tableaux (jaune si URGENT, gris si PM / date / statut terminal). Jaune manuel retiré. |
| Objet des attentes hors liste | **Remis dans le texte** : « → En attente : phase 2 », etc. (sauf les dates « 22/09/2026 »). |
| AODEX | Avis du contrôleur technique : FAIT LE « En attente DEKRA » + « → En attente : avis CT AODEX 20 » (04-020). |
| HAMES | = AHMES (BET VRD). |
| En-tête répétée en haut de page | **Gardée**. |
| Bordures | Toutes les lignes comme la 1re ligne d'ÉTUDES ; **aucune bordure sur les bords extérieurs gauche et droit**. |
| Style de bordure des lignes créées | **Respecter le style du classeur** : ligne courante = traits fins haut et bas + traits fins entre colonnes ; en-tête = trait moyen en bas ; ÉTUDES / TRAVAUX = traits moyens haut et bas ; **la dernière ligne de chaque tableau reste visible** (elle porte le trait épais de bas de tableau du style de tableau). À appliquer aussi par le futur skill à chaque ligne ajoutée. |
| Zone d'impression | Jusqu'à la **fin du dernier tableau** (correction d'un bug de Claude qui imprimait des pages vides). |
| Nombre de pages | À optimiser : lignes vides des tableaux **masquées** (réserve conservée), pas de saut de page avant un TRAVAUX sans ligne visible. |
| Hauteur des lignes | Tout le texte doit être visible. **Méthode (29/09/2026)** : polices Denim INK fournies par José, installées dans l'environnement de travail (jamais versionnées : licence) ; `tools/ajuster_hauteurs.py` fait calculer par LibreOffice le nombre de lignes de texte de chaque observation, puis applique la hauteur de ligne d'Excel (12,75 pt par ligne de texte en 9 pt + 4,5 pt de marge). |

## Décisions du 29/09/2026 (fin)

| Sujet | Décision |
|---|---|
| « Retard AXL » (02-014) | → « Relance » : **validé**. |
| Saut de page avant TRAVAUX | **Gardé** (≈ 48 pages contre 43 pour l'original ; ≈ 40 sans les sauts). |
| Sections MOE-MOA vides | **Non masquées**. |
| En-tête « p » vu sur une capture | Modification manuelle de José : sans objet. |

## Relances, urgences et retards (29/09/2026)

| Statut | Posé par | Sens | Visuel |
|---|---|---|---|
| **Relance** | Claude, automatiquement | échéance atteinte, pas de réponse | — |
| **URGENT** | José (résumé routé) | point bloquant, à traiter en priorité | fond jaune (MFC) |
| **Retard** | José seulement (résumé routé) | le non-respect **impacte le planning** : constat de maîtrise d'œuvre, ouvre la voie aux pénalités | fond orange `FFF8CBAD` (MFC) |

- **Couleur des statuts (30/09/2026)** : tous les statuts de FAIT LE en **rouge non gras**, sauf **PM en gris**. Écrits en texte enrichi pour résister au gris de la mise en forme conditionnelle. Les dates de clôture gardent la règle V1 (rouge la semaine de la clôture).
- Priorité (FAIT LE ne porte qu'une valeur) : **Retard > URGENT > Relance** ; seule la Relance est automatique.
- **Compteur factuel imprimé** : « → Relancé N fois, dernière le … **(sans réponse depuis N j)** » — jours calendaires depuis la 1re relance qui suit la dernière remarque ; absent sur une ligne soldée ou PM.
- **Compteur de retard** : quand José passe une ligne en « Retard », la date du CR est inscrite en colonne masquée **DATE RETARD** (G) ; la colonne masquée **JOURS RETARD** (I) compte les **jours calendaires** (correction de José du 29/09/2026) depuis cette date jusqu'à la date du CR, ou jusqu'à la date de clôture. Formule vérifiée sur 3 cas.
- **Texte imprimé** (à écrire par le skill à chaque CR) : « → En retard depuis le JJ/MM/AAAA : N jours ».
- **Récapitulatif interne** : colonnes « Retard » et « Jours de retard (calendaires) » dans le comptage par onglet ; **bloc 4 RETARDS** (N°, observation, section, depuis le, jours calendaires) dans « Points à traiter », pour lister les entreprises en retard à l'AMO. Colonne masquée **SECTION** (H) : ÉTUDES / TRAVAUX ou code de l'intervenant.
- Périmètre : **ÉTUDES et TRAVAUX** ; **jours seulement**, pas de montant.

## Intégration du CRC-16 (30/09/2026)

| Sujet | Décision |
|---|---|
| « sans réponse depuis N j » | **Absent le jour même** de la 1re relance (0 j) ; affiché à partir d'1 jour. |
| Outil | `tools/integrer_cr.py` : intègre le routage (mises à jour et nouvelles lignes) dans le CR au nouveau format. |
| Hauteur des lignes | **Plus de marge** : mesure sur colonnes rétrécies de 14 %, 13,2 pt par ligne de texte + 9 pt. |
| Couleurs (retour de José sur le CRC-16) | Écrites **dans les cellules** par le script : ajout du jour en **rouge** (texte, dates C/D, date ou statut de FAIT LE, PM compris) ; ligne **PM / soldée / sans objet en gris** (N°, texte, dates) hors ajout du jour ; sinon **noir** ; statut de FAIT LE **rouge**, PM **gris** à partir du CR suivant. |
| MFC de police grise | **Retirée** : elle écrasait la typo manuelle de José (et le rouge d'une date de clôture du jour). Seuls les fonds jaune (URGENT) et orange (Retard) restent en MFC. Conséquence : un PM saisi à la main dans la semaine ne grise la ligne qu'à l'intégration suivante. |
| Texte des lignes non traitées | **Conservé tel quel** (corrections manuelles de José) ; seuls les couleurs et le compteur « sans réponse depuis » sont mis à jour. Le texte n'est recalculé depuis HISTORIQUE que sur les lignes mises à jour ou relancées. |
| Bug corrigé | Après insertion de lignes, les formules de MFC n'étaient pas décalées (AMO-003 grisée à tort) ; les lignes insérées n'étaient pas ajustées en hauteur. |

## Mode relecture (30/09/2026)

Validé par José : il corrige le CR livré dans Excel, puis redonne **le fichier corrigé et la version livrée** ; Claude « remouline » avec `tools/relire_cr.py`.

| José modifie | Traitement |
|---|---|
| HISTORIQUE (F), avec « Au JJ/MM/AAAA » à la date du CR | Ce segment passe en rouge ; OBSERVATIONS recalculé depuis F. |
| OBSERVATIONS (B) seul | Conservé ; **signalé** (à reporter dans HISTORIQUE, sinon perdu à la prochaine mise à jour de la ligne). |
| B et F | F l'emporte ; signalé. |
| ABORDÉ LE / POUR LE / FAIT LE | En rouge. Statut remis à la casse de la liste (« pm » → « PM »). « Relance » posée à la main : inscrite dans HISTORIQUE comme une relance automatique. Ligne passée en PM / soldée : grise sauf l'ajout du jour ; compteur « sans réponse » mis à jour. |
| Ligne de réserve remplie **sans N°** | N° attribué, style du tableau, SECTION, JOURS RETARD, texte rouge, ABORDÉ LE = date du CR s'il est vide. |
| Ligne supprimée, N° en double, dernière ligne de tableau utilisée | Signalé. |

Limite : une modification de mise en forme seule (sans changement de contenu) n'est pas détectée.

## Statut Retard dans les scripts (30/09/2026)

Appliqué par l'intégration et la relecture : DATE RETARD (G) = date du CR à la pose du statut ; ligne imprimée en rouge « → En retard depuis le JJ/MM/AAAA : N jours » (sans compteur le premier jour, comme « sans réponse depuis ») ; ligne retirée quand le statut est levé.

## Nouvelle méthode d'intégration (30/09/2026)

Remplace le tableau de routage (deux relectures jugées trop longues). Détail et conventions : `docs/memo-notes.md`.

| Sujet | Décision |
|---|---|
| Entrées | **Notes de cellule** de José sur le CR précédent (prises en réunion, moins précises qu'avant) + **transcription audio**. |
| Rôle des sources | Les notes **font foi** (ligne, statut, échéance) ; l'audio sert à **reformuler** les notes (toujours) et à repérer les points non notés. Désaccord : la note l'emporte, écart signalé. |
| Notes | Toutes traitées puis supprimées ; une note non citée **bloque** l'intégration. Note sur ÉTUDES / TRAVAUX, sur le titre ou l'en-tête d'un tableau = nouvelle(s) ligne(s) (un paragraphe par ligne). |
| Colonne ROUTAGE | Colonne J, visible à l'écran, **non imprimée**, remplie pour **toutes** les observations touchées ; vidée au remoulinage. |
| Couleur | **Rose** (et non lavande) : `FFFADADD` pour un point venu de l'audio seul, `FFF4A6C0` pour un [?]. Retiré au remoulinage. |
| Onglet « Non routé » | Créé à l'intégration, **non imprimé**, **vidé et masqué** au remoulinage (demande de José, 30/09/2026). |
| Cycle | Notes + audio → CR à relire → corrections de José dans Excel → remoulinage (`relire_cr.py`) → édition finale. |

## Numérotation et skill V3 (30/09/2026)

| Sujet | Décision |
|---|---|
| Colonne N° | **Gardée et imprimée** (option A) malgré la méthode par notes : clé du remoulinage, de « Points à traiter » et du suivi des retards ; référence commune avec les entreprises et l'AMO. |
| Skill V3 | `skill/compte-rendu-chantier-v3/` (scripts recopiés depuis `tools/` par `tools/sync_skill.sh`), paquet `compte-rendu-chantier-v3.skill`. Remplace V1 et V2 pour les classeurs au nouveau format ; V1 / V2 restent pour les classeurs à l'ancien format (ex. DUCLAIR_MIT). |
| Hauteurs sans LibreOffice | Estimation par les métriques de la police (projet claude.ai) : jamais moins de lignes que LibreOffice sur le CRC-16 (193 / 216 identiques, 23 avec 1 ou 2 lignes de plus). |

## Doublons (02/10/2026)

| Sujet | Décision |
|---|---|
| Définition | **Doublon = même destinataire + même objet.** Deux lignes sur le même objet chez deux destinataires différents (deux lots, ou deux intervenants de MOE-MOA) sont une **interface**, pas un doublon : chacune porte l'action de son destinataire. |
| Paires de la migration | Seul **01-004 / 01-005** est un doublon (01-005 masqué par José). 05-015 / 10-024, MOA-001 / 06-026 et AMO-004 / MOE-008 (destinataires différents) sont des interfaces, conservées. |
| Repère bleu | Retiré au remoulinage du CRC-16 (il n'existait pas dans le classeur d'origine). |
| Skill | Avant de créer une ligne : même objet chez le même destinataire → mise à jour de la ligne existante, ou [?] en cas d'hésitation. Interface entre onglets : renvoi dans ROUTAGE seulement. Jamais de « Doublon » automatique (traitement manuel par José). |

## Premier remoulinage réel (CRC-16, 02/10/2026)

Fichier corrigé par José dans Excel : 69 lignes modifiées, 2 lignes ajoutées, 3 supprimées, aucun faux positif dû à l'enregistrement par Excel. Ajustements de `relire_cr.py` :

| Cas rencontré | Traitement |
|---|---|
| Ajouts datés d'un autre jour que la réunion (25/09, 30/09, 02/10) ou sans « Au » | **Tout texte ajouté par José est rouge** (comparaison avec la version livrée), quelle que soit la date ; les dates postérieures à la réunion sont listées dans le rapport, conservées telles quelles. |
| Correction faite dans OBSERVATIONS | **Reportée dans HISTORIQUE** (recherche par le contexte, espaces multiples tolérés), puis texte recalculé ; si le report est impossible, texte corrigé conservé et signalé. |
| Ajout rouge au milieu d'une remarque | Couleurs conservées run par run dans le texte condensé (`cr_texte.condense`). |
| PM saisi dans POUR LE | Déplacé dans FAIT LE (décision du 29/09/2026), signalé. |
| Ligne neuve commençant par « Au JJ/MM/AAAA » | Date retirée (elle est dans ABORDÉ LE). |
| Valeur recalculée par Excel sur une formule inchangée | Pas une modification de José. |
| Lignes supprimées (01-005, 02-055, 03-042) | Signalées ; N° non réutilisés. Préférer masquer (l'historique est perdu). |
| 03-035 | POUR LE (formule « ABORDÉ LE + 14 ») avait glissé au 30/06/2026 après la correction d'ABORDÉ LE ; **figée au 29/09/2026** (échéance du CRC-15). `migration_cr.py` corrigé : POUR LE figée quand ABORDÉ LE est corrigé. |

## Mise en page (retour de José du 02/10/2026)

| Sujet | Décision / correction |
|---|---|
| Titre d'intervenant seul en bas de page (MOE-MOA) | `tools/paginer.py` (lancé par `ajuster_hauteurs.py`) simule la pagination et garde ensemble titre + en-tête + 1re ligne, et ÉTUDES / TRAVAUX + 1re observation. **Tous les sauts sont manuels** (marge 3 %) : Excel et LibreOffice n'arrondissent pas les hauteurs pareil, une coupure automatique tomberait ailleurs. |
| Ligne en gras entre deux sous-sections | Traits épais de haut de page de l'ancienne mise en page, restés en milieu de page (MOE-MOA lignes 73, 115, 154, 174) : supprimés. |
| Traits épais entre observations (charpente) | Bug de migration (modèle de bordure pris sur 03-002, qui avait un trait moyen) : `normaliser_bordures` ramène au trait fin tout trait épais hors en-tête, ligne 1 et ÉTUDES / TRAVAUX, à chaque intégration et remoulinage. |
| Pages blanches en fin de CR | Onglet Photos sans zone d'impression (Excel imprimait A1:N88) : zone limitée au contenu (texte et images), recalculée à chaque passage. |
| Page vide en fin d'onglet (BENOUVILLE lot 06, 02/10/2026) | Lignes à hauteur non fixée (en-têtes, ÉTUDES / TRAVAUX) recalculées un peu plus hautes par le tableur : `paginer.py` leur compte 10 % de plus. Sur HONGUEMARE CRC-16 : 3 sauts déplacés, même nombre de pages (51). |
| Notes de cellule à la migration | Elles suivent leur cellule quand la colonne N° est insérée (BENOUVILLE : note « 7mm » de 02-033, restée sur le N°, ramenée sur l'observation). |

## Ordre du jour de la réunion suivante (02/10/2026, proposition de Claude à valider)

`tools/ordre_du_jour.py`, intégré au skill V3 (étape du remoulinage) plutôt qu'en skill séparé : mêmes entrées, mêmes outils, une seule installation. Word écrit sans bibliothèque. Contenu : approbation du CR ; prioritaires (URGENT, Retard) ; à traiter par intervenant (échéance atteinte le jour de la réunion, ou Relance sans échéance) avec dernier état ; en attente d'un tiers (une ligne par organisme) ; visite ; questions diverses ; présence requise (intervenants ayant un point) et convoqués (colonne C). Premier essai : réunion n° 17 du 06/10/2026, 5 pages (19 prioritaires, 59 à traiter, 45 en attente).

## En attente

- Validation du CRC-16 par José.
- Installation du skill V3 par José, puis premier essai réel depuis Cowork (le gabarit de résumé routé est abandonné au profit des notes de cellule).
- Premier essai réel : un CR annoté dans Excel (les notes de test ont été posées par script, pas par Excel).
