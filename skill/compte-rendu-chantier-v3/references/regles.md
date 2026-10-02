# Règles du CR au nouveau format (registre de José, état du 30/09/2026)

Les scripts appliquent ces règles d'eux-mêmes. Ce fichier sert à **comprendre** ce qu'ils font, à rédiger
OPERATIONS.json en cohérence, et à répondre à José. Source complète : `docs/decisions-HONGUEMARE.md` du dépôt
`compte-rendu-chantier`.

## Sommaire
1. Structure d'une ligne d'observation
2. Texte affiché (OBSERVATIONS) et HISTORIQUE
3. Statuts de FAIT LE
4. Relances automatiques
5. Retard
6. Couleurs
7. Masquage, lignes vides, mise en page
8. Ce que le script ne fait pas

## 1. Structure d'une ligne d'observation

| Col. | Contenu | Imprimée |
|---|---|---|
| A | **N°** `02-014`, `MOA-014`, `BET.SBE-005`… : code de l'intervenant ou du lot + numéro, jamais réutilisé, attribué par le script. Gardé et imprimé (décision du 30/09/2026) : clé du remoulinage, de « Points à traiter » et du suivi des retards. | oui |
| B | **OBSERVATIONS** : texte condensé, recalculé depuis HISTORIQUE | oui |
| C | ABORDÉ LE (date de création, ne change jamais) | oui |
| D | POUR LE (échéance : date, formule `=C+N` pour une ligne neuve, ou vide) | oui |
| E | FAIT LE : **date de clôture ou statut** (liste fermée) | oui |
| F | HISTORIQUE : texte intégral, **fait foi** | non |
| G | DATE RETARD (masquée) | non |
| H | SECTION (masquée) : ÉTUDES / TRAVAUX ou code de l'intervenant | non |
| I | JOURS RETARD (masquée, formule) | non |
| J | ROUTAGE (relecture seulement, vidée au remoulinage) | non |

Onglets masqués : `Référentiel` (codes, prochain N°, statuts, alias), `Points à traiter` (formules Excel 365 collées par José), `Non routé` (relecture seulement).
Onglets intouchables : Coordonnees, Généralités, Modele_lot, Reportage photo. Page de garde : seuls A22 (CRC-NN) et B22 (date) changent ; le reste est en formules.

## 2. Texte affiché et HISTORIQUE

- HISTORIQUE (F) reçoit chaque mise à jour à la suite : `[texte existant]. Au JJ/MM/AAAA [ajout]`, en rouge. On complète, on ne réécrit jamais.
- « Au » ou « AU » avec majuscule seulement : un « au » minuscule est du texte ordinaire.
- OBSERVATIONS (B) est **recalculé** depuis F pour les lignes mises à jour ou relancées :
  sujet initial, puis `→ Au JJ/MM/AAAA : …` pour les **2 dernières remarques** (« [...] » si des remarques plus anciennes sont masquées), puis `→ Relancé N fois, dernière le JJ/MM/AAAA (sans réponse depuis N j)`.
- **Règle A stricte** : seules les relances **postérieures à la dernière remarque** sont comptées ; le compteur repart de 0 à chaque remarque.
- « sans réponse depuis N j » : jours calendaires depuis la 1re relance qui suit la dernière remarque ; absent le jour même (0 j), sur une ligne soldée ou PM.
- Objet d'une attente conservé à part : `→ En attente : phase 2`.
- Lignes **non traitées** ce jour : B conservé tel quel (corrections manuelles de José) ; seuls les couleurs et le compteur changent.

## 3. Statuts de FAIT LE (liste fermée, onglet Référentiel)

`Relance`, `URGENT`, `Retard`, `PM`, `En cours`, `En attente`, `En attente <ORGANISME>` (ENEDIS, SIEGE 27, AXL…), `Annulé`, `Doublon`, `Sans objet`, `Refusé` — ou une **date** (clôture).

- **PM** : pour mémoire ; ligne conservée, grise, jamais relancée, jamais masquée. POUR LE peut garder une date ou rester vide. Un statut actif (Relance, URGENT, En attente) remplace PM.
- Priorité quand plusieurs s'appliquent : **Retard > URGENT > Relance**.
- **Relance** : posée par le script (automatique). **URGENT** et **Retard** : posés par José seulement (via ses notes).
- Date en FAIT LE = clôture : rouge la semaine de clôture, texte gris, ligne masquée au CR suivant.

## 4. Relances automatiques (script)

- Ligne visible, non soldée, sans mise à jour ce jour :
  - FAIT LE = URGENT → relancée à chaque CR ;
  - FAIT LE vide ou Relance, **et** POUR LE atteint (≤ date du CR) → relancée ; FAIT LE vide devient « Relance ».
- Jamais sur PM, En attente…, En cours, Retard, ni sur une ligne masquée ou soldée.
- Effet : ` Au JJ/MM/AAAA Relance` en rouge dans HISTORIQUE, ligne « → Relancé… » recalculée.

## 5. Retard (José seulement : le non-respect impacte le planning)

- DATE RETARD (G) = date du CR à la pose du statut ; JOURS RETARD (I) compte les jours **calendaires** jusqu'à la date du CR ou de clôture.
- Ligne imprimée, en rouge : `→ En retard depuis le JJ/MM/AAAA : N jours` (sans compteur le 1er jour). Retirée si le statut est levé.
- Fond orange `FFF8CBAD` (mise en forme conditionnelle). Bloc 4 RETARDS de « Points à traiter » pour la liste à l'AMO.

## 6. Couleurs (écrites dans les cellules par le script)

- **Ajout du jour en rouge** : texte, dates C/D écrites ce jour, date ou statut de FAIT LE (PM compris).
- Ligne **PM / soldée / statut terminal** : N°, texte et dates en **gris**, sauf l'ajout du jour.
- Autres lignes : **noir** ; statut de FAIT LE **rouge** ; PM gris à partir du CR suivant.
- Mise en forme conditionnelle : seulement les **fonds** jaune (URGENT) et orange (Retard). Pas de MFC de police (elle bloquait la typo manuelle de José).
- Relecture : fond **rose** `FFFADADD` = point venu de l'audio seul ; **rose soutenu** `FFF4A6C0` = [?]. Retirés au remoulinage.
- Couleurs autres que noir / rouge / gris (ex. bleu des doublons) : jamais touchées.

## 7. Masquage, lignes vides, mise en page

- Au CR suivant, une ligne soldée (date ou statut terminal) lors d'un passage précédent et non touchée ce jour est **masquée**.
- 2 lignes de réserve vides (masquées) en fin de sous-section ; insertion de lignes si besoin, bordures du classeur respectées ; la dernière ligne de chaque tableau reste visible (trait épais de fin de tableau).
- Saut de page avant chaque TRAVAUX qui a au moins une ligne visible. Zone d'impression A:E jusqu'à la fin du dernier tableau ; en-tête répété.
- Hauteur des lignes : `ajuster_hauteurs.py` (LibreOffice + polices Denim INK ; sinon estimation, à contrôler).

## 8. Ce que le script ne fait pas (à faire par Claude, ou à signaler)

- Rédiger : c'est le rôle de Claude (voir `redaction.md`).
- Décider d'une clôture, d'un statut ou d'une échéance absents des notes et de l'audio.
- Détecter un doublon entre deux lignes. **Doublon = même destinataire + même objet** (décision du 02/10/2026) ; même objet chez deux destinataires = interface, pas un doublon. Traitement manuel par José, jamais de « Doublon » automatique.
- Corriger une échéance antérieure à ABORDÉ LE : la **signaler**, ne pas la corriger.
