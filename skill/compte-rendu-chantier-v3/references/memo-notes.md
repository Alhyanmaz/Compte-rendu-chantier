# Mémo — prise de notes en réunion (CR HONGUEMARE)

Méthode validée le 30/09/2026 : José prend des **notes de cellule** sur le CR diffusé la semaine précédente, pendant la réunion, et enregistre l'audio. Claude reformule chaque note à l'appui de la transcription, puis intègre. José relit et corrige le CR produit dans Excel. Claude remoule le fichier (mode relecture) avant l'édition finale.

## Comment noter

Toujours par **clic droit > Nouvelle note** (pas « Nouveau commentaire »). Ne rien écrire directement dans les cellules pendant la réunion.

| Cas | Où poser la note |
|---|---|
| Point existant | Sur la ligne concernée, de préférence dans OBSERVATIONS |
| Nouveau point dans un lot | Sur la ligne **ÉTUDES** ou **TRAVAUX** du lot |
| Nouveau point pour un intervenant (MOE-MOA, Concessionnaires) | Sur le titre de l'intervenant (ex. « Maitre d'ouvrage_SIEGE 27 ») ou sur l'en-tête de son tableau |
| Plusieurs nouveaux points au même endroit | Une seule note, **un paragraphe par point** |
| Emplacement incertain | Au plus près, terminé par **?** |

**Contenu** : qui + quoi, en abrégé ; orthographe et accents sans importance. Les chiffres, dates et dimensions sont précieux (l'audio les restitue mal). Le texte entre **[crochets]** est une consigne pour Claude, jamais recopiée.

## Mots-clés (en fin de note)

| Mot-clé | Effet |
|---|---|
| *(aucun)* | Mise à jour simple ; nouvelle ligne : échéance à +7 j |
| **+7**, **+14**, **pour le 06/10** | Échéance (POUR LE) |
| **fin** / **fait** / **fait le 02/10** | Clôture (date dans FAIT LE) |
| **PM** | Pour mémoire |
| **urgent** | URGENT (fond jaune, relancé à chaque CR) |
| **retard** | Retard : impact planning, compteur de jours (fond orange) |
| **en attente ENEDIS** | En attente + organisme |
| **relance** | Relance avant l'échéance (après l'échéance, elle est automatique) |
| **doublon**, **annulé**, **sans objet**, **refusé** | Statut terminal |
| **?** | Claude complète avec l'audio et marque le point [?] |

## Ce que produit Claude (CR « à relire »)

- Chaque note est **reformulée à l'appui de l'audio** (les notes prises en réunion sont moins précises que celles écrites au bureau avec le V1). Aucune note ne reste dans le classeur : une note non traitée bloque l'intégration.
- **Colonne ROUTAGE** (J, non imprimée) sur **toutes** les observations touchées : note d'origine, extrait de l'audio, choix de Claude ; « Relance automatique » pour les relances.
- **Fond rose** `FFFADADD` : point venu de la transcription seule, sans note. **Rose soutenu** `FFF4A6C0` : [?] à trancher (question écrite dans ROUTAGE).
- **Onglet « Non routé »** (non imprimé) : passages de l'audio écartés, avec la raison.

## Relecture par José, puis remoulinage

Corriger directement dans Excel (voir `decisions-HONGUEMARE.md`, « Mode relecture ») : ajouts dans HISTORIQUE avec « Au JJ/MM/AAAA », dates et statuts librement, nouvelles lignes dans une ligne de réserve sans N°. Redonner **le fichier corrigé et la version livrée**. Le remoulinage retire les fonds roses, vide ROUTAGE, vide et masque « Non routé » ; un [?] laissé tel quel est signalé et considéré comme validé.
