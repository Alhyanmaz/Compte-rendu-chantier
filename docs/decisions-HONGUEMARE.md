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

## En attente

- Pose des formules de « Points à traiter » : par Claude (XML), ou collées par José dans Excel.
- Exclusion de « Points à traiter » du PDF : onglet masqué, ou visible et export par sélection d'onglets.
