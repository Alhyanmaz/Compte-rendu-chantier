# -*- coding: utf-8 -*-
"""Réglages de migration DUCLAIR — LE MIT (réhabilitation et extension de l'ancienne école), CRC-44 du 30/09/2026.

Établis par Claude le 02/10/2026 d'après le classeur (aucune décision propre à DUCLAIR n'a encore été prise par José) :
les règles générales du registre docs/decisions-HONGUEMARE.md s'appliquent ; les choix ci-dessous sont des
INTERPRÉTATIONS listées dans docs/rapport-migration-DUCLAIR.md, à valider.
"""
import datetime

ORGANISMES = ['MOA', 'EPFN', 'METROPOLE', 'CICLOP', 'VERITAS', 'PRESENTS', 'ACAU', 'ECLA', 'OCEADE', 'AHMES', 'ESGCB',
              'ATELIERS 59', 'GAMBA', 'CEDN', 'ATELIER 2P', 'MBTP', 'AMAGYS', 'POULINGUE', 'ISOTOIT', 'LA FRATERNELLE',
              'DEGROISILLE', 'BTH', 'BONAUD', 'SOGEP', 'KONE', 'TOTOCOLLANT PUB', 'HARLIN', 'DGS', 'GOUJON VALLEE',
              'ENEDIS', 'GRDF', 'KYNTUS']

CONFIG = {
    'nom': 'DUCLAIR_MIT',
    'date_cr': datetime.datetime(2026, 9, 30),
    # titres des sections de MOE-MOA (minuscules, « œ » lu « oe ») -> code du N°
    'section_codes': [
        ("maitre d'ouvrage_ville", 'MOA'), ("maitre d'ouvrage_epfn", 'EPF'), ('amo_', 'AMO'),
        ('bureau de contr', 'CT'), ('csps_', 'SPS'), ("maitrise d'oeuvre", 'MOE'), ('economiste', 'ECO'),
        ('bet cfo', 'BET.ELE'), ('bet cvc', 'BET.CVC'), ('bet energetique', 'BET.ENV'), ('bet structure', 'BET.STR'),
        ('bet signal', 'BET.SIG'), ('bet vrd', 'BET.VRD'), ('accoustique', 'BET.ACO'), ('bet audit', 'BET.DEC'),
        ('paysagiste', 'PAY'), ('telecom_', 'CON.TEL')],
    'doublons': [],                          # aucun doublon probable parmi les lignes ouvertes (contrôle du 02/10/2026)
    'attente_map': {'': 'En attente', 'CT': 'En attente VERITAS'},
    'attente_notes': {'CT': '« CT » lu comme VERITAS (bureau de contrôle) (interprétation)'},
    'attente_objets': set(),
    'organismes_attente': {o: o for o in ORGANISMES} | {'HARLIN ENERGIE': 'HARLIN', 'BUREAU VERITAS': 'VERITAS'},
    'attente_inconnue': 'objet',            # « En attente hors d'eau » : « En attente » + « → En attente : hors d'eau »
    'speciaux_lignes_masquees': False,      # lignes masquées (archivées) : valeurs atypiques conservées telles quelles
    'abord_fix': {},
    'abord_annee': [],
    'notes_a_supprimer': [],
    'parasites': [('14 - COUVERTURE ', 'B14', 'Caractère « s » isolé dans ABORDÉ LE (ligne vide) supprimé')],
    'sections_supprimees': [],
    'sections_ajoutees': [],
    'coord_fixes': [('H46', 'Esc', 'Exc'), ('A26', 'Esc : Excusé', 'Exc : Excusé'),
                    ('A44', 'Esc : Excusé', 'Exc : Excusé'), ('A50', 'Esc : Excusé', 'Exc : Excusé')],
    'page_de_garde': {'titre': 'C21', 'date': 'B21', 'heure': 'B22', 'prochaine': 'B26',
                      'prochaine_formule': '{d}&"   "&{h}',
                      'valeurs': ('Compte rendu de la réunion de chantier du 30/09/2026', '07/10/2026   9h00')},
    # (code, ligne(s) de Coordonnees) : organisme = 1re ligne, représentants = toutes les lignes
    'referentiel': [('MOA', 7), ('EPF', 8), ('AMO', 10), ('CT', 11), ('SPS', 12), ('MOE', 15), ('ECO', 16),
                    ('BET.ELE', 17), ('BET.CVC', 18), ('BET.ENV', 19), ('BET.VRD', 20), ('BET.STR', 21),
                    ('BET.SIG', 22), ('BET.ACO', 23), ('BET.DEC', 24), ('PAY', 25), ('CON.TEL', 49),
                    ('01', [29, 30]), ('02', 31), ('03', 32), ('04', 33), ('05', 34), ('06', 35), ('07', 36),
                    ('08', 37), ('09', 38), ('10', 39), ('11', 40), ('12', 41), ('13', 42), ('14', 43)],
    'statuts': ['Relance', 'URGENT', 'Retard', 'PM', 'En cours', 'En attente'] +
               ['En attente %s' % o for o in ORGANISMES] + ['En attente concessionnaire',
                                                            'Annulé', 'Doublon', 'Sans objet', 'Refusé'],
    'zones': ['ancienne école', 'extension', 'pensionnat', 'base vie', 'préau'],
    'lexique': ['RSD', 'TEAMS', 'prorata', 'DROC', 'VDI', 'BAL', 'SSI'],
}

# Valeurs de FAIT LE propres à DUCLAIR, sur les lignes VISIBLES (interprétations à valider par José)
_SPECIAUX = {
    'ABANDON MOA': ('Annulé', 'Abandon MOA',
                    'FAIT LE « Abandon MOA » → « Annulé », mention remise dans le texte (interprétation)'),
    'AVANT ENDUIT EXT. SEPTEMBRE': ('PM', 'Avant enduit ext. septembre',
                                    'FAIT LE « Avant enduit ext. Septembre » (POUR LE « PM ») → « PM », mention remise '
                                    'dans le texte (interprétation)'),
}


def statut_special(u, raw):
    sp = _SPECIAUX.get(u)
    return (sp[0], sp[2]) if sp else None


def mention_speciale(raw):
    sp = _SPECIAUX.get(' '.join(raw.upper().split()))
    return sp[1] if sp else None
