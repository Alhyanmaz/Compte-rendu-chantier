# -*- coding: utf-8 -*-
"""Réglages propres à chaque opération pour migration_cr.py (un module par opération).

Chaque module définit CONFIG (dict) et, au besoin, deux fonctions facultatives :
  statut_special(u, raw)  -> (statut, note) ou None   : valeurs de FAIT LE propres à l'opération
  objet_special(raw)      -> texte ou None              : objet à remettre dans le texte (« → En attente : … »)
"""
import importlib


def charger(nom):
    return importlib.import_module('operations.%s' % nom)


def detecter(texte_page_de_garde):
    t = (texte_page_de_garde or '').upper()
    if 'HONGUEMARE' in t:
        return 'honguemare'
    if 'DUCLAIR' in t:
        return 'duclair_mit'
    if 'BENOUVILLE' in t:
        return 'benouville_psc'
    return None
