# OPERATIONS.json — format d'entrée de `integrer_cr.py`

```json
{
  "date_cr": "2026-10-06",
  "crc": "CRC-17",
  "operations": [
    {"type": "maj", "num": "02-066",
     "texte": "AXL indique la mise à jour des plans de fondation pour la fin de semaine (idéalement jeudi).",
     "pour_le": "+7",
     "note": "02 -GROS OEUVRE!B75",
     "routage": "Note : « plans fondation MAJ fin de semaine jeudi +7 » | Audio : « on vous envoie ça jeudi »"},

    {"type": "maj", "num": "04-012", "texte": "Plan des panelettes reçu et visé par ACAU.",
     "fait_le": "2026-10-06", "note": "04 -COUVERTURE!B20", "routage": "Note : « plan panelettes recu, vise ACAU fin »"},

    {"type": "nouvelle", "onglet": "02 -GROS OEUVRE", "code": "02", "section": "TRAVAUX",
     "texte": "Coulage de la dalle de compression prévu le 06 ou 07/10/2026 (première partie jusqu'au JD).",
     "pour_le": "+7", "note": "02 -GROS OEUVRE!B65", "routage": "Note, 1er paragraphe"},

    {"type": "nouvelle", "onglet": "MOE-MOA", "code": "SIE", "section": null,
     "texte": "Communiquer le plan des réseaux existants.", "pour_le": "+14",
     "source": "transcription", "routage": "Audio seul : « le Siège 27 voudrait le plan des réseaux »",
     "doute": "SIEGE 27 ou MOA ?"}
  ],
  "non_route": [
    {"extrait": "« AXL parle du prix des aciers »", "raison": "Information commerciale, non actée en séance"},
    {"extrait": "Note MOA-010 : « rien de neuf ? »", "raison": "Rien de nouveau dans l'audio ; ligne laissée en l'état",
     "note": "MOE-MOA!B16"}
  ]
}
```

## Champs

| Champ | Type | Sens |
|---|---|---|
| `date_cr` | `AAAA-MM-JJ` | Date de la **réunion** (jamais la date de traitement) |
| `crc` | `CRC-NN` | Numéro du CR : celui de la Page de garde A22 du CR précédent + 1 |
| `type` | `maj` / `nouvelle` | Mise à jour d'une ligne existante / création |
| `num` | N° | (maj) Ligne visée, lue dans `lire_notes.py` |
| `onglet`, `code`, `section` | | (nouvelle) Nom **exact** de l'onglet ; code de l'intervenant ou du lot (`02`, `MOA`, `SIE`, `BET.SBE`, `CON.ELE`…) ; `ÉTUDES` / `TRAVAUX`, ou `null` pour un tableau d'intervenant |
| `texte` | | Texte rédigé (maj : **seulement l'ajout**, sans « Au JJ/MM/AAAA ») |
| `pour_le` | `"+N"` / `"AAAA-MM-JJ"` / absent | Échéance. Ligne neuve : `+N` donne la formule `=C+N` |
| `fait_le` | statut / `"AAAA-MM-JJ"` / absent | Statut de la liste fermée, ou date de clôture |
| `note` | `"Onglet!B12"` | Note de José traitée (supprimée du classeur). Une note peut être citée par plusieurs opérations (un paragraphe = une ligne) |
| `source` | `note` (défaut) / `transcription` | `transcription` : point sans note → fond rose |
| `routage` | texte | Colonne ROUTAGE : note d'origine, extrait de l'audio, choix de Claude, écarts |
| `doute` | texte | Question [?] fermée → fond rose soutenu, écrite dans ROUTAGE |
| `non_route` | liste | Passages écartés (`extrait`, `raison`, `note` si une note est concernée) → onglet « Non routé » |

## Contrôles faits par le script

- **Toute** note de cellule d'un onglet d'observations doit être citée (`note` d'une opération ou de `non_route`), sinon arrêt avec la liste des notes oubliées.
- Section introuvable (code / section) : arrêt.
- Rapport JSON : masquées, mises à jour, nouvelles (N° attribués), relances automatiques, réouvertures, insertions, alertes, notes supprimées.
