# -*- coding: utf-8 -*-
"""Réglages de migration HONGUEMARE (groupe scolaire) — décisions de José du 29/09/2026
(docs/decisions-HONGUEMARE.md). Extraits tels quels de migration_cr.py le 02/10/2026."""
import datetime

CONFIG = {
    'nom': 'HONGUEMARE',
    'date_cr': datetime.datetime(2026, 9, 22),
    # (fragment du titre de section, code) -- MOE-MOA et Concessionnaires
    'section_codes': [
        ("maitre d'ouvrage_ville", 'MOA'), ('amo_', 'AMO'), ('bureau de contr', 'CT'), ('csps_', 'SPS'),
        ("maitrise d'oeuvre", 'MOE'), ('economiste', 'ECO'), ('bet cvc', 'BET.CVC'),
        ('bet électricité', 'BET.ELE'), ('bet thermique', 'BET.THE'), ('bet vrd', 'BET.VRD'),
        ('bet cuisine', 'BET.CUI'), ('structure béton', 'BET.SBE'), ('structure bois', 'BET.SBO'),
        ('accoustique', 'BET.ACO'), ('amiante', 'BET.AMI'),
        ('concessionaires_réseaux', 'CON.RES'), ('concessionaires_eaux', 'CON.EAU'),
        ('concessionaires_electricite', 'CON.ELE'), ('concessionaires_assainissement', 'CON.ASS')],
    'doublons': [('MOE-MOA', 35), ('MOE-MOA', 67), ('01 - DESAMIANTAGE', 10), ('01 - DESAMIANTAGE', 11),
                 ('05 - MENUISERIES EXT.', 21), ('10 -CFO-CFA', 30), ('MOE-MOA', 7),
                 ('06 - MENSUIERIE INT.', 32)],
    'attente_map': {'': 'En attente', 'MOA': 'En attente MOA', 'RETOUR MOA': 'En attente MOA',
                    'DU RETOUR DU SIEGE': 'En attente SIEGE 27', 'SIEGE': 'En attente SIEGE 27',
                    'RDV CONCESIONNAIRE': 'En attente concessionnaire',
                    'CONCESIONNAIRE': 'En attente concessionnaire', 'ENEDIS': 'En attente ENEDIS',
                    'ORANGE': 'En attente ORANGE', 'DEVIS BEVELEC': 'En attente BEVELEC',
                    'ACAU': 'En attente ACAU', 'VISA ACAU SUR DT': 'En attente ACAU',
                    'CT': 'En attente DEKRA', 'RETOUR HAMES': 'En attente AHMES'},
    'attente_notes': {'RETOUR HAMES': '« HAMES » lu comme AHMES (BET VRD) (interprétation)'},
    'attente_objets': {'PHASE 2', 'MISE AU POINT CHAUFFERIE', 'MAJ PROCESS', 'BAT', '22/09/2026'},
    'attente_inconnue': 'note',       # précision non reconnue : « En attente » + note au rapport
    'abord_fix': {'02-066': datetime.datetime(2026, 6, 2), '05-009': datetime.datetime(2026, 5, 20),
                  '03-035': datetime.datetime(2026, 6, 16)},
    'abord_annee': [('02 -GROS OEUVRE', 83, 2027, datetime.datetime(2026, 7, 7),
                     'ABORDÉ LE 07/07/2027 corrigé en 07/07/2026 (interprétation : faute de frappe sur l\'année, '
                     'le texte parle du 15/07/2026)')],
    'notes_a_supprimer': [('MOE-MOA', 'A35', 'Note de cellule vide « José Mazzarese: » supprimée')],
    'parasites': [('02 -GROS OEUVRE', 'F84', 'Caractère « s » isolé hors tableau supprimé')],
    'sections_supprimees': ['BET.CUI'],
    'sections_ajoutees': [{'apres': 'MOA', 'titre': "Maitre d'ouvrage_SIEGE 27", 'code': 'SIE',
                           'ligne_titre': 4, 'ligne_vide': 5}],
    'coord_fixes': [('A23', 'DESAMIANRAGE', 'DESAMIANTAGE'), ('A20', 'AOUSTIQUE', 'ACOUSTIQUE'),
                    ('H35', 'Esc', 'Exc'), ('A41', 'Esc : Excusé', 'Exc : Excusé')],
    'page_de_garde': {'titre': 'C22', 'date': 'B22', 'heure': 'B23', 'prochaine': 'B26',
                      'prochaine_formule': '" RDV chantier "&{d}&" à "&{h}',
                      'valeurs': ('Compte rendu de la réunion de chantier du 22/09/2026', ' RDV chantier 29/09/2026 à 9H00')},
    'referentiel': [('MOA', 6), ('SIE', 7), ('AMO', 8), ('CT', 9), ('SPS', 10), ('MOE', 11), ('ECO', 12),
                    ('BET.ELE', 13), ('BET.CVC', 14), ('BET.THE', 15), ('BET.VRD', 16), ('BET.SBO', 17),
                    ('BET.SBE', 18), ('BET.AMI', 19), ('BET.ACO', 20)] +
                   [('%02d' % k, 22 + k) for k in range(1, 12)] +
                   [('CON.EAU', 36), ('CON.ELE', 37), ('CON.ASS', 38), ('CON.RES', 39)],
    'statuts': ['Relance', 'URGENT', 'Retard', 'PM', 'En cours', 'En attente', 'En attente MOA', 'En attente SIEGE 27',
                'En attente CICLOP', 'En attente DEKRA', 'En attente VERITAS', 'En attente ACAU', 'En attente ECLA',
                'En attente CONCEPT NF', 'En attente ECHOS', 'En attente AHMES', 'En attente BESB', 'En attente ESGCB',
                'En attente ACCEO', 'En attente GAMBA', 'En attente DEMOLAF', 'En attente AXL', 'En attente AGC',
                'En attente GOUJON VALLEE', 'En attente AVA', 'En attente MCO', 'En attente REVNOR', 'En attente NORDEC',
                'En attente ELAIRGIE', 'En attente BEVELEC', 'En attente CFB TP', 'En attente concessionnaire',
                'En attente ENEDIS', 'En attente ORANGE', 'En attente SRPN', 'En attente SPANC',
                'Annulé', 'Doublon', 'Sans objet', 'Refusé'],
    'zones': ['préau', 'SHED', 'bâtiment A', 'chaufferie', 'restaurant scolaire', 'école existante', 'extension',
              'mairie', 'cours anglaises', 'vide sanitaire', 'base vie', 'salle polyvalente', 'maternelle',
              'circulation', 'cuisine', 'self', 'pignon'],
    'lexique': ['RSD', 'longrines', 'hourdis', 'courettes', 'acodrains', 'MOB', 'tebopins', 'pré-isolé', 'soubassement',
                'rejingot', 'couvertine', 'précadre', 'claire-voie', 'citerneau', 'ANC', 'AEP', 'arbalétrier', 'bac',
                'noue', 'tranchée commune', 'dallage', 'sous-œuvre', 'enduit', 'cloisonnettes', 'CTA', 'DAS', 'BPE',
                'TEAMS', 'prorata'],
}


def statut_special(u, raw):
    if u.startswith('AODEX'):
        return 'En attente DEKRA', ('avis du CT « %s » : statut « En attente DEKRA », référence reportée dans le texte '
                                    '(interprétation)' % raw.strip())
    if u == 'RETARD AXL':
        return 'Relance', '« Retard AXL » converti en « Relance » (interprétation)'
    return None


def objet_special(raw):
    if raw.upper().startswith('AODEX'):
        return 'avis CT %s' % raw
    return None
