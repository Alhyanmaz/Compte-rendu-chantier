# -*- coding: utf-8 -*-
"""Réglages de migration BENOUVILLE — pôle socio-culturel (construction d'un centre socio-culturel), CRC-16 du
24/09/2026.

Établis par Claude le 02/10/2026 d'après le classeur : règles générales de docs/decisions-HONGUEMARE.md ; les choix
ci-dessous sont des INTERPRÉTATIONS listées dans docs/rapport-migration-BENOUVILLE.md, à valider par José.
"""
import datetime

ORGANISMES = ['MOA', 'VERITAS', 'ACAU', 'AECO', 'BIELEC', 'AREHA', 'AHMES', 'KUBE', 'ORFEA', 'EXEO', 'VERT LATITUDE',
              'HNTP', 'SEEL', 'CHANU HD', 'ENC CGB', 'EDB', 'TAILLE PIERRES', 'AVA', 'LE COGUIC', 'ISOPLAF',
              'HARET DECO', 'BONAUD', 'GILSON', 'ORONA', 'VOLTEC', 'SCF', 'OUEST TERRASSEMENT', 'CAEN LA MER',
              'SDEC', 'ENEDIS', 'GRDF', 'COVAGE']

CONFIG = {
    'nom': 'BENOUVILLE_PSC',
    'date_cr': datetime.datetime(2026, 9, 24),
    'section_codes': [
        ("maitre d'ouvrage_", 'MOA'), ('bureau de contr', 'CT'), ('csps_', 'SPS'), ("maitrise d'oeuvre", 'MOE'),
        ('economiste', 'ECO'), ('bet électricité', 'BET.ELE'), ('bet cvc', 'BET.CVC'), ('bet hqe', 'BET.HQE'),
        ('bet structure', 'BET.STR'), ('bet vrd', 'BET.VRD'), ('accoustique', 'BET.ACO'), ('paysagiste', 'PAY')],
    'doublons': [],                          # ressemblances examinées le 02/10/2026 : objets différents
    'attente_map': {'': 'En attente', 'VISA ACAU': 'En attente ACAU', 'RETOUR ACAU': 'En attente ACAU',
                    'SYNTHÈSE ACAU': 'En attente ACAU', 'VALIDATION MOE': 'En attente ACAU',
                    'VALIDATION MOA': 'En attente MOA', 'HARETDECO': 'En attente HARET DECO'},
    'attente_notes': {'VALIDATION MOE': '« validation MOE » lu comme ACAU (maître d\'œuvre) (interprétation)'},
    'attente_objets': set(),
    'organismes_attente': {o: o for o in ORGANISMES},
    'attente_inconnue': 'objet',
    'speciaux_lignes_masquees': False,
    'abord_fix': {},
    'abord_annee': [],
    'notes_a_supprimer': [],
    'parasites': [],
    'sections_supprimees': [],
    'sections_ajoutees': [],
    'coord_fixes': [('H42', 'Esc', 'Exc'), ('A19', 'Esc : Excusé', 'Exc : Excusé'),
                    ('A39', 'Esc : Excusé', 'Exc : Excusé'), ('A49', 'Esc : Excusé', 'Exc : Excusé')],
    # titre « n°[16] » conservé (en formule sur le numéro de CR) ; prochaine réunion laissée en texte libre
    'page_de_garde': {'titre': 'C21', 'date': 'B21', 'heure': 'B22', 'crc': 'A21', 'prochaine': None,
                      'titre_formule': '"Compte rendu de la réunion de chantier n°["&MID({crc},5,4)&"] "',
                      'valeurs': ('Compte rendu de la réunion de chantier n°[16] ', None)},
    'referentiel': [('MOA', 7), ('CT', 8), ('SPS', 9), ('MOE', 10), ('ECO', 11), ('BET.ELE', 12), ('BET.CVC', 13),
                    ('BET.VRD', 14), ('BET.STR', 15), ('BET.ACO', 16), ('BET.HQE', 17), ('PAY', 18)] +
                   [('%02d' % k, 21 + k) for k in range(1, 18)],
    'statuts': ['Relance', 'URGENT', 'Retard', 'PM', 'En cours', 'En attente'] +
               ['En attente %s' % o for o in ORGANISMES] + ['Annulé', 'Doublon', 'Sans objet', 'Refusé'],
    'zones': [],
    'lexique': ['RE2020', 'VS', 'DT', 'EXE', 'SSI'],
}


# Valeurs de FAIT LE propres à l'opération, lignes VISIBLES (interprétations à valider)
_SPECIAUX = {
    'EN ATTENTE CONFIRMATION PARQUET MOA': ('En attente MOA', 'En attente : confirmation parquet',
                                           '« En attente confirmation parquet MOA » → « En attente MOA », objet remis '
                                           'dans le texte (interprétation)'),
}


def statut_special(u, raw):
    sp = _SPECIAUX.get(u)
    return (sp[0], sp[2]) if sp else None


def mention_speciale(raw):
    sp = _SPECIAUX.get(' '.join(raw.upper().split()))
    return sp[1] if sp else None
