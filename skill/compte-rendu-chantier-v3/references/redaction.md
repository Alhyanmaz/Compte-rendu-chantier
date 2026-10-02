# Rédaction et routage (notes de José + transcription audio)

Repris des skills V1 et V2 (corrections de José des 18, 21, 22 et 23/09/2026), adapté à la méthode par notes du 30/09/2026.

## Sommaire
1. Rôle des deux sources
2. Lire une note
3. Utiliser la transcription
4. Points sans note (audio seul)
5. Rédiger
6. Choisir la ligne ou la sous-section
7. Doute : [?] ou non routé

## 1. Rôle des deux sources

- La **note** de José fait foi : la cellule où elle est posée désigne la ligne ou la sous-section ; ses mots-clés fixent statut et échéance.
- La **transcription** sert à **reformuler** chaque note (les notes prises en réunion sont brèves et imprécises), à retrouver les chiffres, dates, noms et formulations, et à repérer ce qui n'a pas été noté.
- Désaccord entre les deux : **la note l'emporte**. L'écart se signale dans ROUTAGE (« Audio : … ≠ note »).

## 2. Lire une note

- `lire_notes.py` donne, pour chaque note, ce qu'elle désigne : `maj <N°>` ou `nouvelle(s) : <code>, <section>`.
- Une note sur ÉTUDES / TRAVAUX, sur le titre ou l'en-tête d'un tableau, ou sur une ligne vide : **nouvelle(s) ligne(s)**, un paragraphe = une ligne.
- Mots-clés (voir `memo-notes.md`) : ils pilotent POUR LE / FAIT LE et ne sont **jamais recopiés** dans le texte.
  - `fin`, `fait`, `fait le JJ/MM` → `fait_le` = date ; `PM`, `urgent`, `retard`, `relance`, `en attente X`, `doublon`, `annulé`, `sans objet`, `refusé` → statut ; `+7`, `+14`, `pour le JJ/MM` → `pour_le`.
  - Nouvelle ligne sans mot-clé d'échéance → `pour_le` = `+7`.
  - `?` en fin de note → point [?] (champ `doute`).
- Texte entre [crochets] : consigne pour Claude, jamais recopiée.
- Une note sur une ligne soldée (date en FAIT LE) qui n'est pas un mot-clé de clôture **rouvre** la ligne : le script vide FAIT LE si l'opération ne donne pas de `fait_le`.
- Note illisible, ambiguë, ou mot-clé contradictoire (deux statuts) : ne pas choisir en silence → `doute`.

## 3. Utiliser la transcription

Matière bruitée : filtrer ce qui appelle une **action ou un constat à tracer**.

- **Sigles mal transcrits** : DT/DICT, FT/PT, CVP/CVC, CFO-CFA, VISA/VIC, EXE/DESC, PPSPS/PGC (voir `lexique-btp.md`). Un sigle incertain ne se corrige pas au jugé → `doute`.
- **Noms** d'entreprises et de personnes : à valider contre l'onglet `Coordonnees` et les alias du `Référentiel`. Jamais de rapprochement « par ressemblance ». Alias connu : HAMES = AHMES (BET VRD).
- **Chiffres, dates, délais** : un délai entendu ne devient une échéance que s'il est clairement attribué à un point. Vérifier le jour de la semaine d'une date citée.
- **Qui doit faire** : critère de rangement principal ; souvent perdu dans une transcription.
- Un fichier audio brut n'est pas exploitable dans l'environnement : il faut la transcription texte.

## 4. Points sans note (audio seul)

- Une action ou une décision nette, non notée par José → opération avec `"source": "transcription"` (fond rose).
- Tout le reste (bavardage, rappel déjà acté, information commerciale non actée, point non attribuable) → `non_route` avec la raison.
- Ne jamais inventer d'échéance ni de statut pour un point audio seul : `pour_le` par défaut (`+7`) sur une ligne neuve, rien sur une mise à jour.

## 5. Rédiger

- Verbe à l'infinitif en tête : fournir, confirmer, prévoir, communiquer, faire un retour sur, étudier, réaliser, valider, transmettre, notifier. Pas de « il conviendra de ».
- Constat : sujet + verbe au présent (« AXL indique… », « La MOA confirme… »).
- Registre bref et factuel ; sigles et noms d'entreprises en capitales, tels quels.
- Mise à jour : écrire **seulement l'ajout** (le script ajoute « Au JJ/MM/AAAA »). Une ligne neuve ne commence jamais par « Au [date] ».
- Formulation ambiguë : rester au plus près de l'audio et le signaler dans ROUTAGE ; ne jamais enrichir d'un détail non dit.
- Ne jamais inventer une date, un statut, une échéance, un montant ou un nom absent des notes, de l'audio ou du classeur.

## 6. Choisir la ligne ou la sous-section (points sans note, ou note « ? »)

1. **Qui doit faire** décide de l'onglet ; la MOA ou la MOE qui doit agir → `MOE-MOA`, sous-section de celui qui agit. Exception : « ACAU se rapproche de X pour connaître… » → demande à X, dans l'onglet de X.
2. Lot cité, entreprise citée (`Coordonnees`), sigle d'intervenant → onglet correspondant. Concessionnaire ou réseau → `Concessionnaires`.
3. **ÉTUDES** = se solde par un document (plan, FT, DT, note de calcul, VISA, PPSPS, agrément). **TRAVAUX** = se solde par un ouvrage constaté sur site. Les deux → deux opérations.
4. **Même objet** (même document, même ouvrage, même zone) → mise à jour. Sujet seulement **voisin** → ligne neuve.
   Avant toute création, chercher le même objet **chez le même destinataire** (doublon, définition du 02/10/2026) : s'il existe, mise à jour de cette ligne, ou `doute` en cas d'hésitation. Le même objet chez un **autre** destinataire est une interface : créer la ligne et écrire le renvoi dans `routage` (« interface avec 10-024 »), jamais dans le texte imprimé.
5. Décision ou demande de la MOA → ligne neuve dans MOE-MOA, sous l'intervenant concerné.
6. Point répété sans nouveauté → mise à jour « Relance » (`fait_le` = `Relance`).
7. ACAU ou la MOA doit répondre → `En attente ACAU` / `En attente MOA` (à proposer dans ROUTAGE, pas à imposer).

## 7. Doute : [?] ou non routé

- **[?]** (`doute`) : on sait **à peu près** quoi écrire et où, mais un choix reste à confirmer (ligne, statut, sigle, nom, date). Question **fermée** : « 02-017 ou nouvelle ligne TRAVAUX ? ».
- **Non routé** : on ne sait pas où le mettre, ou il n'y a rien à tracer. Jamais « par défaut dans MOE-MOA » : un point mal rangé est notifié à la mauvaise entreprise.
