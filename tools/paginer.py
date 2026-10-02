# -*- coding: utf-8 -*-
"""Sauts de page des onglets d'observations : un titre n'est jamais laissé seul en bas de page.

Usage : python paginer.py ENTREE.xlsx SORTIE.xlsx   (appelé par ajuster_hauteurs.py, après le calcul des hauteurs)

Excel n'a pas d'option « solidaire du suivant ». On simule donc sa pagination (hauteur imprimable de la page,
hauteur de chaque ligne visible, lignes de titre répétées) et on pose un saut manuel avant chaque bloc qui ne
tiendrait pas sur la page en cours :
- titre d'intervenant (MOE-MOA, Concessionnaires) + en-tête du tableau + première ligne visible du tableau ;
- ligne ÉTUDES / TRAVAUX + première observation visible.
Les sauts avant chaque TRAVAUX ayant une ligne visible (décision du 29/09/2026) sont conservés.
Tous les sauts sont manuels (y compris les coupures ordinaires) : la pagination ne dépend plus des arrondis
d'Excel ou de LibreOffice. Marge de sécurité : 3 % de la hauteur de page. Demande de José du 02/10/2026.
"""
import datetime
import os
import re
import sys

from lxml import etree

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from cr_xml import Book  # noqa: E402
import cr_classeur as K  # noqa: E402

N = K.N
PAPIER_PT = {9: 841.89, 1: 792.0, 8: 1190.55}     # hauteur portrait : A4, Letter, A3
MARGE = 0.97


def hauteur_imprimable(ws):
    ps = ws.find(N + 'pageSetup')
    pm = ws.find(N + 'pageMargins')
    paper = int(ps.get('paperSize', 9)) if ps is not None else 9
    h = PAPIER_PT.get(paper, 841.89)
    if ps is not None and ps.get('orientation') == 'landscape':
        h = 595.28
    top = float(pm.get('top', 0.75)) if pm is not None else 0.75
    bot = float(pm.get('bottom', 0.75)) if pm is not None else 0.75
    scale = int(ps.get('scale', 100)) if ps is not None else 100
    return (h - (top + bot) * 72) * 100.0 / scale * MARGE


def titres_repetes(sh):
    wbx = K.b.xml('xl/workbook.xml')
    names = [s.get('name') for s in wbx.find(N + 'sheets')]
    for dn in wbx.iter(N + 'definedName'):
        if dn.get('name') == '_xlnm.Print_Titles' and dn.get('localSheetId') == str(names.index(sh)):
            m = re.search(r'\$(\d+):\$(\d+)', dn.text or '')
            if m:
                return int(m.group(1)), int(m.group(2))
    return None


def fin_zone(sh):
    wbx = K.b.xml('xl/workbook.xml')
    names = [s.get('name') for s in wbx.find(N + 'sheets')]
    for dn in wbx.iter(N + 'definedName'):
        if dn.get('name') == '_xlnm.Print_Area' and dn.get('localSheetId') == str(names.index(sh)):
            m = re.search(r'\$[A-Z]+\$(\d+)\s*$', dn.text or '')
            if m:
                return int(m.group(1))
    return None


def paginer(sh):
    ws = K.b.ws(sh)
    rows, tabs, info = K.model(sh)
    sfp = ws.find(N + 'sheetFormatPr')
    defaut = float(sfp.get('defaultRowHeight', 15)) if sfp is not None else 15.0
    last = fin_zone(sh) or max(rows)

    def h(r):
        row = rows.get(r)
        if row is None:
            return defaut
        if row.get('hidden') in ('1', 'true'):
            return 0.0
        return float(row.get('ht', defaut))

    visibles = [r for r in range(1, last + 1) if h(r) > 0]

    def suivante_visible(r, stop):
        return next((k for k in range(r + 1, stop + 1) if h(k) > 0), None)

    # blocs à garder ensemble : {première ligne: dernière ligne}
    blocs = {}
    for t in tabs:
        r0, r1 = t['r0'], t['r1']
        titre = next((k for k in range(r0 - 1, max(0, r0 - 4), -1)
                      if k in rows and any(K.text_of(c).strip() for c in rows[k]) and k not in info), None)
        debut = titre or r0
        fin = suivante_visible(r0, r1) or r0
        if info.get(fin, {}).get('kind') == 'section':        # en-tête puis ÉTUDES : garder aussi la 1re observation
            fin = suivante_visible(fin, r1) or fin
        blocs[debut] = fin
        for r, it in info.items():
            if it.get('table') == t['name'] and it['kind'] == 'section' and h(r) > 0 and r != fin:
                nxt = suivante_visible(r, r1)
                if nxt:
                    blocs.setdefault(r, nxt)
    # sauts imposés : avant chaque TRAVAUX qui a une observation visible
    imposes = set()
    for r, it in info.items():
        if it['kind'] == 'section' and it['section'] == 'TRAVAUX':
            t = [t for t in tabs if t['name'] == it['table']][0]
            if any(info[k]['kind'] == 'obs' and not info[k]['hidden'] for k in range(r + 1, t['r1'] + 1)):
                imposes.add(r)
    H = hauteur_imprimable(ws)
    rep = titres_repetes(sh)
    h_rep = sum(h(k) for k in range(rep[0], rep[1] + 1)) if rep else 0.0
    sauts, used, page_start = [], 0.0, True
    for r in visibles:
        if r in imposes and not page_start:
            sauts.append(r - 1)
            used, page_start = 0.0, True
        if page_start and used == 0.0 and sauts and rep and r > rep[1]:
            used = h_rep
        need = sum(h(k) for k in range(r, blocs[r] + 1)) if r in blocs else h(r)
        if used + need > H and not page_start:
            # tous les sauts sont posés en manuel : Excel et LibreOffice n'arrondissent pas les hauteurs de la même
            # façon, une coupure automatique pourrait tomber ailleurs et laisser un titre seul en bas de page
            sauts.append(r - 1)
            used = h_rep if (rep and r > rep[1]) else 0.0
        used += h(r)
        page_start = False
    rb = ws.find(N + 'rowBreaks')
    if rb is None:
        rb = etree.Element(N + 'rowBreaks')
        anchor = ws.find(N + 'headerFooter')
        if anchor is None:
            anchor = ws.find(N + 'pageSetup')
        if anchor is None:
            anchor = ws.find(N + 'pageMargins')
        anchor.addnext(rb)
    for brk in list(rb):
        rb.remove(brk)
    for k in sorted(set(sauts)):
        etree.SubElement(rb, N + 'brk', id=str(k), max='16383', man='1')
    rb.set('count', str(len(set(sauts))))
    rb.set('manualBreakCount', str(len(set(sauts))))
    if not len(rb):
        ws.remove(rb)
    K.b.touch(K.b.sheets[sh])
    return sorted(set(sauts))


def main(src, dst):
    K.init(Book(src), datetime.datetime.now())
    for sh in K.obs_sheets():
        if sh == 'Modele_lot':
            continue
        s = paginer(sh)
        print('%-28s sauts après les lignes %s' % (sh, s))
    K.b.save(dst)


if __name__ == '__main__':
    main(sys.argv[1], sys.argv[2])
