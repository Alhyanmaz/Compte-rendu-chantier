# -*- coding: utf-8 -*-
"""Migration du classeur CR HONGUEMARE vers le format optimisé (copie de test).

Usage : python migration_cr.py SOURCE.xlsx SORTIE.xlsx RAPPORT.json

Édition XML chirurgicale (cr_xml.py du skill V1) : seules les parties modifiées sont
réécrites. openpyxl n'est utilisé qu'en LECTURE. Décisions appliquées : voir
docs/decisions-HONGUEMARE.md.
"""
import copy
import datetime
import json
import os
import re
import sys
import zipfile

from lxml import etree

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from cr_xml import (Book, N, RNS, RED, BLACK, GREY, EPOCH, _set_color, _rpr_from_font,  # noqa: E402
                    color_of, col_index, col_letter)
import openpyxl  # noqa: E402

SRC, DST, REPORT = sys.argv[1:4]
DATE_CR = datetime.datetime(2026, 9, 22)
PNS = 'http://schemas.openxmlformats.org/package/2006/relationships'
CTNS = 'http://schemas.openxmlformats.org/package/2006/content-types'
BLUE_FILL = 'FFBDD7EE'
HIDE_EMPTY = True     # masquer les lignes vides des tableaux (optimisation des pages)
W_NUM = 8.0       # largeur de la colonne N° ; retirée à OBSERVATIONS pour garder la largeur de page
CPL_B = 50        # caractères par ligne estimés dans OBSERVATIONS (largeur ~45, police 9 pt)

LOT_SHEETS = ['01 - DESAMIANTAGE', '02 -GROS OEUVRE', '03 - CHARPENTE', '04 -COUVERTURE',
              '05 - MENUISERIES EXT.', '06 - MENSUIERIE INT.', '07 - CARRELAGE SOL SOUPLES',
              '08 - PEINTURES', '09 -CVP', '10 -CFO-CFA', '11 - VRD']
OBS_SHEETS = ['MOE-MOA', 'Concessionnaires'] + LOT_SHEETS
MIGR_SHEETS = OBS_SHEETS + ['Modele_lot']
SECTION_CODES = [  # (fragment du titre de section, code) -- MOE-MOA et Concessionnaires
    ("maitre d'ouvrage_ville", 'MOA'), ('amo_', 'AMO'), ('bureau de contr', 'CT'), ('csps_', 'SPS'),
    ("maitrise d'oeuvre", 'MOE'), ('economiste', 'ECO'), ('bet cvc', 'BET.CVC'),
    ('bet électricité', 'BET.ELE'), ('bet thermique', 'BET.THE'), ('bet vrd', 'BET.VRD'),
    ('bet cuisine', 'BET.CUI'), ('structure béton', 'BET.SBE'), ('structure bois', 'BET.SBO'),
    ('accoustique', 'BET.ACO'), ('amiante', 'BET.AMI'),
    ('concessionaires_réseaux', 'CON.RES'), ('concessionaires_eaux', 'CON.EAU'),
    ('concessionaires_electricite', 'CON.ELE'), ('concessionaires_assainissement', 'CON.ASS')]
DOUBLONS = [('MOE-MOA', 35), ('MOE-MOA', 67), ('01 - DESAMIANTAGE', 10), ('01 - DESAMIANTAGE', 11),
            ('05 - MENUISERIES EXT.', 21), ('10 -CFO-CFA', 30), ('MOE-MOA', 7),
            ('06 - MENSUIERIE INT.', 32)]
ATTENTE_MAP = {'': 'En attente', 'MOA': 'En attente MOA', 'RETOUR MOA': 'En attente MOA',
               'DU RETOUR DU SIEGE': 'En attente SIEGE 27', 'SIEGE': 'En attente SIEGE 27',
               'RDV CONCESIONNAIRE': 'En attente concessionnaire',
               'CONCESIONNAIRE': 'En attente concessionnaire', 'ENEDIS': 'En attente ENEDIS',
               'ORANGE': 'En attente ORANGE', 'DEVIS BEVELEC': 'En attente BEVELEC',
               'ACAU': 'En attente ACAU', 'VISA ACAU SUR DT': 'En attente ACAU',
               'CT': 'En attente DEKRA', 'RETOUR HAMES': 'En attente AHMES'}
ATTENTE_PERTE = {'PHASE 2', 'MISE AU POINT CHAUFFERIE', 'MAJ PROCESS', 'BAT', '22/09/2026'}
ABORD_FIX = {  # ABORDÉ LE corrigés sur décision de José (29/09/2026) : N° -> date
    '02-066': datetime.datetime(2026, 6, 2), '05-009': datetime.datetime(2026, 5, 20),
    '03-035': datetime.datetime(2026, 6, 16)}
def attente_objet(d):
    """Objet d'une attente qui ne rentre pas dans la liste fermée (remis dans le texte, décision du 29/09/2026)."""
    if not isinstance(d, str):
        return None
    raw = ' '.join(d.split())
    if raw.upper().startswith('AODEX'):
        return 'avis CT %s' % raw
    m = re.match(r'^en att?e?n?te\s*(.*)$', raw.replace('attnte', 'attente').replace('Attnte', 'Attente'), re.I)
    if m and m.group(1).strip().upper() in ATTENTE_PERTE and not re.match(r'^\d{2}/\d{2}/\d{4}$', m.group(1).strip()):
        return m.group(1).strip()
    return None


TERMINAUX = {'ANNULÉ': 'Annulé', 'ANNULE': 'Annulé', 'DOUBLON': 'Doublon', 'SANS OBJET': 'Sans objet',
             'REFUSÉ': 'Refusé', 'REFUSE': 'Refusé'}

report = {'statuts': [], 'pm': [], 'anomalies_corrigees': [], 'anomalies_signalees': [],
          'structure': [], 'doublons': [], 'condenses': 0, 'numerotation': {}, 'hors_liste': []}

# ------------------------------------------------------------------ lecture seule
wb_f = openpyxl.load_workbook(SRC)                    # formules
wb_v = openpyxl.load_workbook(SRC, data_only=True)    # valeurs en cache

b = Book(SRC)
b.formulas_added = True    # formules décalées/ajoutées : calcChain retiré, recalcul forcé
_xf_cache = {}


def xf_clone(s, color=None, fill_rgb=None, halign=None, valign=None, wrap=None, numfmt=None, bold=None, size=None):
    """Clone un style de cellule (cellXfs) avec des modifications ; ne touche jamais l'existant."""
    key = (int(s or 0), repr(color), fill_rgb, halign, valign, wrap, numfmt, bold, size)
    if key in _xf_cache:
        return _xf_cache[key]
    st = b.styles
    fonts, fills, xfs = st.find(N + 'fonts'), st.find(N + 'fills'), st.find(N + 'cellXfs')
    xf = copy.deepcopy(xfs[int(s or 0)])
    if color is not None or bold is not None or size is not None:
        f = copy.deepcopy(fonts[int(xf.get('fontId', 0))])
        if size is not None:
            for e in f.findall(N + 'sz'):
                e.set('val', str(size))
        if color is not None:
            _set_color(f, color)
        if bold is not None:
            for e in f.findall(N + 'b'):
                f.remove(e)
            if bold:
                f.insert(0, etree.Element(N + 'b'))
        fonts.append(f)
        fonts.set('count', str(len(fonts)))
        xf.set('fontId', str(len(fonts) - 1))
        xf.set('applyFont', '1')
    if fill_rgb is not None:
        fl = etree.SubElement(fills, N + 'fill')
        pf = etree.SubElement(fl, N + 'patternFill', patternType='solid')
        etree.SubElement(pf, N + 'fgColor', rgb=fill_rgb)
        etree.SubElement(pf, N + 'bgColor', indexed='64')
        fills.set('count', str(len(fills)))
        xf.set('fillId', str(len(fills) - 1))
        xf.set('applyFill', '1')
    if numfmt is not None:
        xf.set('numFmtId', str(numfmt))
        xf.set('applyNumberFormat', '1')
    if halign or valign or wrap is not None:
        al = xf.find(N + 'alignment')
        if al is None:
            al = etree.SubElement(xf, N + 'alignment')
            xf.remove(al)
            xf.insert(0, al)
        if halign:
            al.set('horizontal', halign)
        if valign:
            al.set('vertical', valign)
        if wrap is not None:
            al.set('wrapText', '1' if wrap else '0')
        xf.set('applyAlignment', '1')
    xfs.append(xf)
    xfs.set('count', str(len(xfs)))
    b.touch('xl/styles.xml')
    _xf_cache[key] = str(len(xfs) - 1)
    return _xf_cache[key]


def custom_numfmt(code):
    nfs = b.styles.find(N + 'numFmts')
    if nfs is None:
        nfs = etree.Element(N + 'numFmts', count='0')
        b.styles.insert(0, nfs)
    for f in nfs:
        if f.get('formatCode') == code:
            return int(f.get('numFmtId'))
    nid = max([int(f.get('numFmtId')) for f in nfs] + [163]) + 1
    etree.SubElement(nfs, N + 'numFmt', numFmtId=str(nid), formatCode=code)
    nfs.set('count', str(len(nfs)))
    b.touch('xl/styles.xml')
    return nid


def si_new(runs):
    """Nouvelle entrée sharedStrings ; runs = [(rPr|None, texte)]. Renvoie l'index."""
    si = etree.SubElement(b.sst, N + 'si')
    if len(runs) == 1 and runs[0][0] is None:
        t = etree.SubElement(si, N + 't')
        t.text = runs[0][1]
        if runs[0][1] != runs[0][1].strip() or '\n' in runs[0][1]:
            t.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
    else:
        for rpr, txt in runs:
            r = etree.SubElement(si, N + 'r')
            if rpr is not None:
                r.append(copy.deepcopy(rpr))
            t = etree.SubElement(r, N + 't')
            t.text = txt
            if txt != txt.strip() or '\n' in txt:
                t.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
    b.si.append(si)
    b.sst.set('uniqueCount', str(len(b.si)))
    b.sst.set('count', str(int(b.sst.get('count', len(b.si))) + 1))
    b.touch('xl/sharedStrings.xml')
    return len(b.si) - 1


def runs_of_si(idx, s):
    base = b._font_of(s)
    src = b.si[idx]
    rs = src.findall(N + 'r')
    if not rs:
        t = src.find(N + 't')
        return [(_rpr_from_font(base), t.text or '')] if t is not None else []
    out = []
    for r in rs:
        rpr = r.find(N + 'rPr')
        out.append((copy.deepcopy(rpr) if rpr is not None else _rpr_from_font(base), r.find(N + 't').text or ''))
    return out


def set_cell_si(c, idx, s=None):
    for ch in list(c):
        c.remove(ch)
    c.set('t', 's')
    if s is not None:
        c.set('s', s)
    etree.SubElement(c, N + 'v').text = str(idx)


def set_cell_empty(c):
    for ch in list(c):
        c.remove(ch)
    c.attrib.pop('t', None)


def cell_in_row(row, col, s=None):
    """Cellule `col` de l'élément <row>, créée à sa place si absente."""
    ref = '%s%s' % (col, row.get('r'))
    ci = col_index(ref)
    for c in row.iter(N + 'c'):
        if c.get('r') == ref:
            return c
        if col_index(c.get('r')) > ci:
            new = etree.Element(N + 'c', r=ref)
            c.addprevious(new)
            if s is not None:
                new.set('s', s)
            return new
    new = etree.SubElement(row, N + 'c', r=ref)
    if s is not None:
        new.set('s', s)
    return new


# ------------------------------------------------------------------ texte : typos et condensation
TYPO = re.compile(r'\b(\d{2})/(\d{2})(20\d{2})\b')
TYPO5 = re.compile(r'\b(\d{2})/(\d{2})/(202)(\d)(\d)\b')   # « 08/09/20256 » -> 2026 (interprétation)
AU = re.compile(r'(?:(?<=\s)|(?<=\.)|^)(?:[Aa]u|AU)\s+(\d{1,2}/\d{1,2}/(?:\d{4}|\d{2}))(?![\d])')
BARE = re.compile(r'^[\s.,;:!-]*(relance|urgent|rappel)[\s.,;:!-]*$', re.I)


def fix_runs(runs):
    out, fixes = [], []
    for rpr, t in runs:
        t2 = TYPO.sub(r'\1/\2/\3', t)
        t3 = TYPO5.sub(lambda m: '%s/%s/2026' % (m.group(1), m.group(2)), t2)   # « 20256 », « 20265 » : année du chantier
        if t3 != t:
            fixes += ['%s/%s%s' % x for x in TYPO.findall(t)]
            fixes += ['%s/%s/%s%s%s' % x for x in TYPO5.findall(t2)]
        out.append((rpr, t3))
    return out, fixes


def rpr_at(runs, pos):
    k = 0
    for rpr, t in runs:
        if k + len(t) > pos:
            return rpr
        k += len(t)
    return runs[-1][0]


def bold_rpr(rpr):
    r = copy.deepcopy(rpr)
    for e in r.findall(N + 'b'):
        r.remove(e)
    r.insert(0, etree.Element(N + 'b'))
    return r


def plain_rpr(rpr):
    r = copy.deepcopy(rpr)
    for e in r.findall(N + 'b'):
        r.remove(e)
    return r


def slice_runs(runs, a, b):
    out, k = [], 0
    for rpr, t in runs:
        lo, hi = max(a, k), min(b, k + len(t))
        if lo < hi:
            out.append((rpr, t[lo - k:hi - k]))
        k += len(t)
    return out


def head_runs(runs, a, b):
    """Sujet initial : mise en forme d'origine conservée (gras annulé le 29/09/2026), espaces de bord retirés."""
    rs = slice_runs(runs, a, b)
    while rs and not rs[0][1].strip():
        rs.pop(0)
    while rs and not rs[-1][1].strip():
        rs.pop()
    if rs:
        rs[0] = (rs[0][0], rs[0][1].lstrip())
        rs[-1] = (rs[-1][0], rs[-1][1].rstrip())
    return rs


def full_date(dt):
    d, m, y = dt.split('/')
    return '%02d/%02d/%s' % (int(d), int(m), y if len(y) == 4 else '20' + y)


UNDATED = re.compile(r'\s*[-–]\s*\d{1,2}/\d{1,2}/\d{2,4}\s*(?:relance|urgent|rappel)\b\.?', re.I)


def drop_spans(runs, spans):
    out, pos = [], 0
    text_len = sum(len(t) for _, t in runs)
    for a, b in spans + [(text_len, text_len)]:
        out += slice_runs(runs, pos, a)
        pos = b
    return out


def condense(runs, abord=None):
    """Structure toute observation : sujet initial en gras, puis « → Au JJ/MM/AAAA : … » et
    « → Relancé N fois, dernière le … ». Renvoie (runs, nb relances, date dernière relance)."""
    text = ''.join(t for _, t in runs)
    und = [(m.start(), m.end()) for m in UNDATED.finditer(text)]
    if und:                                        # relances sans « Au JJ/MM/AAAA » : supprimées (29/09/2026)
        runs = drop_spans(runs, und)
        text = ''.join(t for _, t in runs)
    segs = [(m.start(), m.end(), m.group(1)) for m in AU.finditer(text)]
    if not text.strip():
        return None
    if not segs:                                   # observation sans historique : sujet seul, en gras
        return head_runs(runs, 0, len(text)), 0, None, None
    pieces = []
    for k, (st, en, dt) in enumerate(segs):
        stop = segs[k + 1][0] if k + 1 < len(segs) else len(text)
        body = text[en:stop].strip().lstrip(':').strip()
        pieces.append((st, dt, body))
    head_end = segs[0][0]
    head = text[:head_end].strip()
    head_pos = 0
    head_start, mismatch = 0, None
    if not head:                                   # texte qui commence par « Au JJ/MM/AAAA »
        head_end = segs[1][0] if len(segs) > 1 else len(text)
        pieces.pop(0)
        d0 = full_date(segs[0][2])
        if isinstance(abord, datetime.datetime) and abord.strftime('%d/%m/%Y') == d0:
            head_start = segs[0][1]                # date = ABORDÉ LE : « Au JJ/MM/AAAA » retiré du sujet
            while head_start < head_end and text[head_start] in ' :,;.-\u00a0':
                head_start += 1
        else:
            mismatch = (d0, abord.strftime('%d/%m/%Y') if isinstance(abord, datetime.datetime) else 'vide')
    bare = [p for p in pieces if p[2] and BARE.match(p[2])]
    # Seules les relances postérieures à la dernière remarque sont affichées
    # (« Relancé N fois, dernière le … ») : le compteur repart de 0 à chaque remarque.
    events = []                                   # ('S', piece) ou ('R', [pieces])
    for p in pieces:
        if not p[2]:
            continue
        if BARE.match(p[2]):
            if events and events[-1][0] == 'R':
                events[-1][1].append(p)
            else:
                events.append(('R', [p]))
        else:
            events.append(('S', p))
    s_idx = [i for i, e in enumerate(events) if e[0] == 'S']
    start = s_idx[-2] if len(s_idx) >= 2 else (s_idx[0] if s_idx else 0)
    # Règle A (29/09/2026) : seules les relances postérieures à la dernière remarque sont affichées.
    last_s = s_idx[-1] if s_idx else -1
    shown = [e for i, e in enumerate(events[start:], start) if e[0] == 'S' or i > last_s]
    omitted = [e for e in events[:start] if e[0] == 'S']
    out = head_runs(runs, head_start, head_end)
    if omitted:
        n = len(omitted)
        out.append((plain_rpr(rpr_at(runs, omitted[0][1][0])), '\n[...]'))
    for kind, e in shown:
        if kind == 'S':
            st, dt, body = e
            out.append((plain_rpr(rpr_at(runs, st)), '\n→ Au %s : %s' % (full_date(dt), body)))
        else:
            last = e[-1]
            out.append((plain_rpr(rpr_at(runs, last[0])), '\n→ Relancé %d fois, dernière le %s' % (len(e), full_date(last[1]))))
    # Sujet initial en gras ; lignes suivantes sans gras, couleurs de la source
    return out, len(bare), bare[-1][1] if bare else None, mismatch


# ------------------------------------------------------------------ statuts FAIT LE
def norm_status(d):
    """Valeur FAIT LE (lecture openpyxl) -> (nouvelle valeur, note)."""
    if d is None or isinstance(d, datetime.datetime):
        return d, None
    raw = str(d)
    u = ' '.join(raw.upper().split())
    if u == 'RELANCE':
        return 'Relance', None
    if u == 'URGENT':
        return 'URGENT', None
    if u == 'PM':
        return 'PM', None
    if u == 'EN COURS':
        return 'En cours', None
    if u in TERMINAUX:
        return TERMINAUX[u], None
    if u.startswith('AODEX'):
        return 'En attente DEKRA', 'avis du CT « %s » : statut « En attente DEKRA », référence reportée dans le texte (interprétation)' % raw.strip()
    if u == 'RETARD AXL':
        return 'Relance', '« Retard AXL » converti en « Relance » (interprétation)'
    m = re.match(r'^EN ATT?E?N?TE\s*(.*)$', u.replace('ATTNTE', 'ATTENTE'))
    if m:
        rest = m.group(1).strip()
        if rest in ATTENTE_MAP:
            note = '« HAMES » lu comme AHMES (BET VRD) (interprétation)' if rest == 'RETOUR HAMES' else None
            return ATTENTE_MAP[rest], note
        if rest in ATTENTE_PERTE:
            return 'En attente', 'objet de l\'attente « %s » non repris dans FAIT LE : à reporter dans le texte ?' % rest.lower()
        return 'En attente', 'précision « %s » non reconnue' % rest
    return None, 'valeur hors liste laissée telle quelle'


# ------------------------------------------------------------------ formules
A1 = re.compile(r"(?<![A-Za-z_\[\]!'\$\.])(\$?)([A-Z]{1,3})(\$?)(\d+)(?![\d(A-Za-z_])")


def remap_formula(txt, rowmap, colshift=1):
    def rep(m):
        r = int(m.group(4))
        if r not in rowmap:
            raise ValueError('référence vers une ligne supprimée : %s' % txt)
        return '%s%s%s%d' % (m.group(1), col_letter(col_index(m.group(2)) + colshift), m.group(3), rowmap[r])
    return A1.sub(rep, txt)


def offset_formula(txt, dr):
    def rep(m):
        r = int(m.group(4)) + (0 if m.group(3) else dr)
        return '%s%s%s%d' % (m.group(1), m.group(2), m.group(3), r)
    return A1.sub(rep, txt)


def expand_shared(ws):
    masters = {}
    for c in ws.iter(N + 'c'):
        f = c.find(N + 'f')
        if f is not None and f.get('t') == 'shared' and f.get('ref'):
            masters[f.get('si')] = (int(re.sub(r'[A-Z]', '', c.get('r'))), f.text)
    for c in ws.iter(N + 'c'):
        f = c.find(N + 'f')
        if f is not None and f.get('t') == 'shared':
            r0, txt = masters[f.get('si')]
            r = int(re.sub(r'[A-Z]', '', c.get('r')))
            f.text = offset_formula(txt, r - r0)
            for k in ('t', 'ref', 'si'):
                f.attrib.pop(k, None)


# ------------------------------------------------------------------ tableaux
def resolve(base_part, target):
    d = base_part.rsplit('/', 1)[0]
    out = []
    for x in (d + '/' + target).split('/'):
        out.pop() if x == '..' else out.append(x)
    return '/'.join(out)


def sheet_tables(sheet):
    part = b.sheets[sheet]
    d, f = part.rsplit('/', 1)
    rp = '%s/_rels/%s.rels' % (d, f)
    res = []
    for r in b.xml(rp):
        if r.get('Type').endswith('/table'):
            tp = resolve(part, r.get('Target'))
            t = b.xml(tp)
            a, z = t.get('ref').split(':')
            res.append(dict(part=tp, rid=r.get('Id'), name=t.get('name'),
                            r0=int(re.sub(r'\D', '', a)), r1=int(re.sub(r'\D', '', z))))
    return sorted(res, key=lambda x: x['r0'])


def section_code(title):
    low = (title or '').lower()
    for k, code in SECTION_CODES:
        if k in low:
            return code
    raise KeyError(title)


# ------------------------------------------------------------------ nouvelle partie de paquet
def add_part(name, data_bytes, content_type):
    b.data[name] = data_bytes
    zi = zipfile.ZipInfo(name, date_time=(2026, 9, 29, 12, 0, 0))
    zi.compress_type = zipfile.ZIP_DEFLATED
    b.infos.append(zi)
    ct = b.xml('[Content_Types].xml')
    etree.SubElement(ct, '{%s}Override' % CTNS, PartName='/' + name, ContentType=content_type)
    b.touch('[Content_Types].xml')


def add_rel(part_rels, rtype, target):
    if part_rels not in b.data:
        b.data[part_rels] = b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="%s"/>' % PNS.encode()
        zi = zipfile.ZipInfo(part_rels, date_time=(2026, 9, 29, 12, 0, 0))
        zi.compress_type = zipfile.ZIP_DEFLATED
        b.infos.append(zi)
    root = b.xml(part_rels)
    ids = {r.get('Id') for r in root}
    k = 1
    while 'rId%d' % k in ids:
        k += 1
    etree.SubElement(root, '{%s}Relationship' % PNS, Id='rId%d' % k, Type=rtype, Target=target)
    b.touch(part_rels)
    return 'rId%d' % k


def next_table_id():
    mx = 0
    for name in b.data:
        if name.startswith('xl/tables/table'):
            mx = max(mx, int(b.xml(name).get('id')))
    return mx + 1


# ================================================================== 0. note vide MOE-MOA A35
b.delete_note('MOE-MOA', 'A35')
report['anomalies_corrigees'].append(['MOE-MOA', 'A35', 'Note de cellule vide « José Mazzarese: » supprimée'])

# ================================================================== 1. analyse (coordonnées d'origine)
info = {}   # (sheet, r) -> dict
numbering_order = []
for sh in MIGR_SHEETS:
    wsf, wsv = wb_f[sh], wb_v[sh]
    tabs = sheet_tables(sh)
    lot_code = sh[:2] if sh in LOT_SHEETS else None
    counters = {}
    for t in tabs:
        if sh == 'Modele_lot':
            code = None
        elif lot_code:
            code = lot_code
        else:
            title = None
            for r in range(t['r0'] - 1, 0, -1):
                v = wsf.cell(r, 1).value
                if isinstance(v, str) and v.strip():
                    title = v
                    break
            code = section_code(title)
        t['code'] = code
        for r in range(t['r0'] + 1, t['r1'] + 1):
            a = wsf.cell(r, 1).value
            at = a.strip() if isinstance(a, str) else ''
            others = [wsf.cell(r, k).value for k in (2, 3, 4)]
            kind = 'empty'
            if at.upper() in ('ÉTUDES', 'TRAVAUX', 'SYNTHESE'):
                kind = 'section'
            elif at or any(v is not None for v in others):
                kind = 'obs'
            num = None
            if kind == 'obs' and code:
                counters[code] = counters.get(code, 0) + 1
                num = '%s-%03d' % (code, counters[code])
            info[(sh, r)] = dict(kind=kind, table=t['name'], num=num, hidden=bool(wsf.row_dimensions[r].hidden),
                                 A=a, B=wsf.cell(r, 2).value, C=wsf.cell(r, 3).value, D=wsf.cell(r, 4).value,
                                 Cv=wsv.cell(r, 3).value, Dv=wsv.cell(r, 4).value)
        info[(sh, t['r0'])] = dict(kind='header', table=t['name'])
    report['numerotation'].update({c: n for c, n in counters.items()})
    info[(sh, '_tables')] = tabs

# ================================================================== 2. structure (lignes, colonne N°, HISTORIQUE)
new_of_all = {}
new_data_rows = {}
for sh in MIGR_SHEETS:
    part = b.sheets[sh]
    ws = b.ws(sh)
    sd = ws.find(N + 'sheetData')
    expand_shared(ws)
    if sh == '02 -GROS OEUVRE':   # « s » isolé hors tableau
        for c in list(ws.iter(N + 'c')):
            if c.get('r') == 'F84':
                c.getparent().remove(c)
        report['anomalies_corrigees'].append([sh, 'F84', 'Caractère « s » isolé hors tableau supprimé'])
    rows = {int(r.get('r')): r for r in sd.findall(N + 'row')}
    max_row = max(rows)
    tabs = info[(sh, '_tables')]
    deleted, inserts = set(), {}   # inserts[after_row] = [(kind, template_row, table_name, payload)]

    if sh == 'MOE-MOA':
        cre = [t for t in tabs if t['code'] == 'BET.CUI'][0]
        deleted |= set(range(cre['r0'] - 1, cre['r1'] + 2))       # titre, tableau, ligne vide suivante
        report['structure'].append([sh, 'lignes %d à %d' % (cre['r0'] - 1, cre['r1'] + 1),
                                    'Section « BET cuisine_CREACEPT » supprimée (1 observation : « %s »)'
                                    % str(info[(sh, cre['r0'] + 1)]['A']).strip()])
        moa = [t for t in tabs if t['code'] == 'MOA'][0]
        inserts.setdefault(moa['r1'] + 1, []).extend([
            ('title', 4, None, "Maitre d'ouvrage_SIEGE 27"), ('blank', 5, None, None),
            ('header', moa['r0'], 'NEW_SIE', None), ('data', moa['r1'], 'NEW_SIE', None),
            ('data', moa['r1'], 'NEW_SIE', None), ('blank', moa['r1'] + 1, None, None)])
        report['structure'].append([sh, 'après la ligne %d' % (moa['r1'] + 1),
                                    'Section « Maitre d\'ouvrage_SIEGE 27 » créée (tableau vide, code SIE)'])
    # réserve de 2 lignes vides par sous-section
    for t in tabs:
        if t['name'] in {x['name'] for x in tabs if x.get('code') == 'BET.CUI'}:
            continue
        secs = [r for r in range(t['r0'] + 1, t['r1'] + 1) if info[(sh, r)]['kind'] == 'section']
        bounds, start = [], t['r0'] + 1
        for sr in secs:
            if sr > start:
                bounds.append((start, sr - 1))
            start = sr + 1
        bounds.append((start, t['r1']))
        for s0, s1 in bounds:
            if s1 < s0:
                continue
            k, tmpl = 0, None
            for r in range(s1, s0 - 1, -1):
                it = info[(sh, r)]
                if it['kind'] == 'empty':
                    if not it['hidden']:
                        k += 1
                    continue
                break
            for r in range(s1, s0 - 1, -1):
                if info[(sh, r)]['kind'] == 'obs':
                    tmpl = r
                    break
            tmpl = tmpl or s1
            need = max(0, 2 - k)
            if need:
                inserts.setdefault(s1, []).extend([('data', tmpl, t['name'], None)] * need)
                report['structure'].append([sh, 'après la ligne %d' % s1, '%d ligne(s) vide(s) ajoutée(s) en fin de sous-section' % need])

    # plan des lignes
    plan = []
    for r in range(1, max_row + 1):
        if r not in deleted:
            plan.append(('src', r))
        for ins in inserts.get(r, []):
            plan.append(('new',) + ins)
    new_of, R = {}, 0
    newrows = []
    for p in plan:
        R += 1
        if p[0] == 'src':
            new_of[p[1]] = R
        newrows.append((R, p))
    new_of_all[sh] = new_of

    # construction du nouveau sheetData
    new_sd = etree.Element(N + 'sheetData')
    table_rows = {}   # nom de tableau -> [rows]
    for R, p in newrows:
        if p[0] == 'src':
            el = rows.get(p[1])
            if el is None:
                continue
            el = copy.deepcopy(el)
            it = info.get((sh, p[1]))
            if it and it.get('table'):
                table_rows.setdefault(it['table'], []).append(R)
        else:
            kind, tr, tname, payload = p[1:]
            src = rows.get(tr)
            el = etree.Element(N + 'row')
            if src is not None:
                for k2, v in src.attrib.items():
                    if k2 not in ('r', 'hidden', 'ht', 'customHeight', 'spans'):
                        el.set(k2, v)
                if kind in ('data', 'header', 'title'):
                    for c in src.findall(N + 'c'):
                        nc = etree.SubElement(el, N + 'c', r=c.get('r'))
                        if c.get('s'):
                            nc.set('s', c.get('s'))
                        if kind == 'header' or (kind == 'title' and c.get('r').startswith('A')):
                            for ch in c:
                                nc.append(copy.deepcopy(ch))
                            if c.get('t'):
                                nc.set('t', c.get('t'))
                if kind == 'title':
                    a = [c for c in el if c.get('r').startswith('A')]
                    if a:
                        set_cell_si(a[0], si_new([(None, payload)]))
                if kind == 'header' and tr in rows:
                    el.set('ht', rows[tr].get('ht', '18'))
            if tname:
                table_rows.setdefault(tname, []).append(R)
            if kind == 'data':
                new_data_rows.setdefault(sh, set()).add(R)
        el.set('r', str(R))
        el.attrib.pop('spans', None)
        for c in el.findall(N + 'c'):
            col = re.sub(r'\d+', '', c.get('r'))
            c.set('r', '%s%d' % (col_letter(col_index(col) + 1), R))
            f = c.find(N + 'f')
            if f is not None and f.text:
                f.text = remap_formula(f.text, new_of)
        new_sd.append(el)
    ws.replace(sd, new_sd)
    newrow_el = {int(r.get('r')): r for r in new_sd}
    last = max(newrow_el)

    # colonnes : A = N°, B..E = ancien A..D, F = HISTORIQUE
    cols = ws.find(N + 'cols')
    wA = None
    for c in cols:
        if int(c.get('min')) == 1:
            wA = c.get('width')
        c.set('min', str(int(c.get('min')) + 1))
        c.set('max', str(int(c.get('max')) + 1))
    for c in cols:
        if int(c.get('min')) == 2:
            c.set('width', '%.2f' % (float(wA) - W_NUM))
    cols.insert(0, etree.Element(N + 'col', min='1', max='1', width='%.2f' % W_NUM, customWidth='1'))
    etree.SubElement(cols, N + 'col', min='6', max='6', width=wA or '53', customWidth='1')

    # fusions : A{r}:D{r} -> A{r}:E{r}, contenu ramené en A
    mcs = ws.find(N + 'mergeCells')
    if mcs is not None:
        for mc in mcs:
            a, z = mc.get('ref').split(':')
            ra, rz = new_of[int(re.sub(r'\D', '', a))], new_of[int(re.sub(r'\D', '', z))]
            ca, cz = re.sub(r'\d', '', a), re.sub(r'\d', '', z)
            nca = 'A' if ca == 'A' else col_letter(col_index(ca) + 1)
            mc.set('ref', '%s%d:%s%d' % (nca, ra, col_letter(col_index(cz) + 1), rz))
            if ca == 'A':
                for rr in range(ra, rz + 1):
                    row = newrow_el.get(rr)
                    if row is None:
                        continue
                    cb = [c for c in row if c.get('r') == 'B%d' % rr]
                    if cb:
                        cb[0].set('r', 'A%d' % rr)
                        nb = etree.Element(N + 'c', r='B%d' % rr)
                        if cb[0].get('s'):
                            nb.set('s', cb[0].get('s'))
                        cb[0].addnext(nb)
    dim = ws.find(N + 'dimension')
    dim.set('ref', 'A1:F%d' % last)
    for sv in ws.iter(N + 'sheetView'):
        sv.attrib.pop('topLeftCell', None)
        for sel in list(sv):
            if etree.QName(sel).localname in ('selection', 'pane'):
                sv.remove(sel)
        etree.SubElement(sv, N + 'selection', activeCell='A1', sqref='A1')

    # tableaux : plage, colonnes N° et HISTORIQUE
    for t in tabs:
        if t.get('code') == 'BET.CUI':
            continue
        rr = sorted(table_rows[t['name']])
        tx = b.xml(t['part'])
        ref = 'A%d:F%d' % (rr[0], rr[-1])
        tx.set('ref', ref)
        af = tx.find(N + 'autoFilter')
        if af is not None:
            af.set('ref', ref)
            for fc in af.findall(N + 'filterColumn'):
                fc.set('colId', str(int(fc.get('colId')) + 1))
            for cid in (0, 5):
                e = etree.Element(N + 'filterColumn', colId=str(cid), hiddenButton='1')
                if cid == 0:
                    af.insert(0, e)
                else:
                    af.append(e)
        tcs = tx.find(N + 'tableColumns')
        ids = [int(tc.get('id')) for tc in tcs]
        e0 = etree.Element(N + 'tableColumn', id=str(max(ids) + 1), name='N°')
        e5 = etree.Element(N + 'tableColumn', id=str(max(ids) + 2), name='HISTORIQUE')
        tcs.insert(0, e0)
        tcs.append(e5)
        tcs.set('count', str(len(tcs)))
        b.touch(t['part'])
        t['new_rows'] = rr
    # suppression du tableau CREACEPT
    if sh == 'MOE-MOA':
        cre = [t for t in tabs if t['code'] == 'BET.CUI'][0]
        tp = ws.find(N + 'tableParts')
        for e in list(tp):
            if e.get('{%s}id' % RNS) == cre['rid']:
                tp.remove(e)
        tp.set('count', str(len(tp)))
        d, f = part.rsplit('/', 1)
        rels = b.xml('%s/_rels/%s.rels' % (d, f))
        for r in list(rels):
            if r.get('Id') == cre['rid']:
                rels.remove(r)
        b.touch('%s/_rels/%s.rels' % (d, f))
        del b.data[cre['part']]
        b.infos = [i for i in b.infos if i.filename != cre['part']]
        ct = b.xml('[Content_Types].xml')
        for e in list(ct):
            if e.get('PartName') == '/' + cre['part']:
                ct.remove(e)
        b.touch('[Content_Types].xml')
        # nouveau tableau SIEGE 27
        rr = sorted(table_rows['NEW_SIE'])
        tid = next_table_id()
        tname = 'TableauMOEMOA_SIE'
        tpart = 'xl/tables/table%d.xml' % (max(int(re.sub(r'\D', '', n)) for n in b.data if n.startswith('xl/tables/table')) + 1)
        moa_t = b.xml([t for t in tabs if t['code'] == 'MOA'][0]['part'])
        tx = copy.deepcopy(moa_t)
        tx.set('id', str(tid))
        for k in list(tx.attrib):
            if k.endswith('}uid'):
                tx.attrib.pop(k)
        tx.set('name', tname)
        tx.set('displayName', tname)
        tx.set('ref', 'A%d:F%d' % (rr[0], rr[-1]))
        af = tx.find(N + 'autoFilter')
        if af is not None:
            af.set('ref', tx.get('ref'))
            for k in list(af.attrib):
                if k.endswith('}uid'):
                    af.attrib.pop(k)
        for tc in tx.iter(N + 'tableColumn'):
            for k in list(tc.attrib):
                if k.endswith('}uid'):
                    tc.attrib.pop(k)
        add_part(tpart, etree.tostring(tx, xml_declaration=True, encoding='UTF-8', standalone=True),
                 'application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml')
        rid = add_rel('%s/_rels/%s.rels' % (d, f), 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/table',
                      '../tables/' + tpart.rsplit('/', 1)[1])
        etree.SubElement(tp, N + 'tablePart', {'{%s}id' % RNS: rid})
        tp.set('count', str(len(tp)))
        tabs.append(dict(part=tpart, name=tname, code='SIE', new_rows=rr, r0=None))
        info[(sh, '_sie_rows')] = rr
    b.touch(part)

# ================================================================== 3. contenu : N°, HISTORIQUE, texte condensé, statuts
greyish = lambda runs: bool(runs) and all(color_of(r) == 'GREY' for r, t in runs if t.strip())
for sh in MIGR_SHEETS:
    ws = b.ws(sh)
    new_of = new_of_all[sh]
    rowel = {int(r.get('r')): r for r in ws.find(N + 'sheetData')}
    for (s2, r), it in list(info.items()):
        if s2 != sh or not isinstance(r, int):
            continue
        R = new_of.get(r)
        if R is None or R not in rowel:
            continue
        row = rowel[R]
        cB = cell_in_row(row, 'B')
        sA = cB.get('s')
        if it['kind'] == 'header':
            for col, txt in (('A', 'N°'), ('F', 'HISTORIQUE')):
                c = cell_in_row(row, col, sA)
                set_cell_si(c, si_new([(None, txt)]), sA)
            continue
        if it['kind'] in ('section', 'empty'):
            cell_in_row(row, 'A', sA)
            cell_in_row(row, 'F', sA)
            continue
        # observation
        if cB.get('t') == 's':
            runs = runs_of_si(int(cB.find(N + 'v').text), sA)
        elif cB.find(N + 'v') is not None:
            runs = [(_rpr_from_font(b._font_of(sA)), cB.find(N + 'v').text)]
        else:
            runs = []
        fixed, fixes = fix_runs(runs)
        if fixes:
            report['anomalies_corrigees'].append([sh, '%s (ligne %d)' % (it['num'], r),
                                                  'Date mal saisie corrigée : ' + ', '.join(sorted(set(fixes)))])
        hist_idx = int(cB.find(N + 'v').text) if (cB.get('t') == 's' and not fixes) else (si_new(fixed) if fixed else None)
        # N°
        grey = greyish(runs)
        sN = xf_clone(sA, color=GREY if grey else BLACK, halign='center', valign='top', wrap=False, size=7)
        cN = cell_in_row(row, 'A', sN)
        cN.set('s', sN)
        if it['num']:
            set_cell_si(cN, si_new([(None, it['num'])]))
        # HISTORIQUE
        cF = cell_in_row(row, 'F', sA)
        if hist_idx is not None:
            set_cell_si(cF, hist_idx, sA)
        # OBSERVATIONS condensées
        if it['num'] in ABORD_FIX:
            nd = ABORD_FIX[it['num']]
            cC = cell_in_row(row, 'C')
            if not b.is_date_style(cC.get('s')):
                ref_c = [c for (s3, r3), i3 in info.items() if s3 == sh and isinstance(r3, int) and i3.get('kind') == 'obs'
                         and isinstance(i3.get('B'), datetime.datetime)][0]
                cC.set('s', b.cell(sh, 'C%d' % new_of[ref_c], False).get('s'))
            set_cell_empty(cC)
            etree.SubElement(cC, N + 'v').text = str((nd - EPOCH).days)
            report['anomalies_corrigees'].append([sh, it['num'], 'ABORDÉ LE %s → %s (décision de José ; date retirée du sujet)'
                                                  % (it['B'].strftime('%d/%m/%Y') if isinstance(it['B'], datetime.datetime) else 'vide', nd.strftime('%d/%m/%Y'))])
            it['B'] = nd
        cond = condense(fixed, it['B']) if fixed else None
        objet = attente_objet(it['Dv'])
        if cond and objet:
            cond[0].append((plain_rpr(cond[0][-1][0]), '\n→ En attente : %s' % objet))
            report.setdefault('objets_remis', []).append([sh, it['num'], objet])
        shown = ''.join(t for _, t in fixed)
        if cond:
            nr, nrel, dlast, mism = cond
            if mism:
                report.setdefault('sujet_date_a_trancher', []).append(
                    [sh, it['num'], 'sujet daté du %s, ABORDÉ LE %s : date laissée dans le sujet' % mism])
            set_cell_si(cB, si_new(nr))
            report['condenses'] += 1
            shown = ''.join(t for _, t in nr)
        elif fixes:
            set_cell_si(cB, hist_idx)
        if not it['hidden']:   # hauteur : OBSERVATIONS est plus étroite de W_NUM
            lines = sum(max(1, -(-len(x) // CPL_B)) for x in shown.split('\n'))
            ht = max(15.0, lines * 12.2 + 4)
            if not cond:
                ht = max(ht, float(row.get('ht', '0')))
            row.set('ht', '%.1f' % ht)
            row.set('customHeight', '1')
        # statut FAIT LE (E) et PM (D)
        cD, cE = cell_in_row(row, 'D'), cell_in_row(row, 'E')
        c_pm = isinstance(it['C'], str) and it['C'].strip().upper() == 'PM'
        newE, note = norm_status(it['Dv'])
        old = it['Dv']
        if c_pm:
            set_cell_empty(cD)
            if old is None:
                newE = 'PM'
                if cD.get('s'):
                    cE.set('s', cD.get('s'))
                report['pm'].append([sh, it['num'], 'POUR LE « PM » → FAIT LE « PM »'])
            else:
                report['pm'].append([sh, it['num'], 'POUR LE « PM » retiré ; FAIT LE conservé : %s' %
                                     (old.strftime('%d/%m/%Y') if isinstance(old, datetime.datetime) else '« %s »' % str(old).strip())])
        if newE is None and old is not None:
            report['hors_liste'].append([sh, it['num'], str(old).strip(), note])
        elif isinstance(newE, str) and newE != (str(old) if old is not None else None):
            oldn = ' '.join(str(old).split()) if old is not None else ''
            if not (c_pm and old is None) and oldn != newE:
                report['statuts'].append([sh, it['num'], oldn, newE, note or ''])
            set_cell_si(cE, si_new([(None, newE)]))
            need = -(-len(newE) // 10) * 12.2 + 4
            if not it['hidden'] and float(row.get('ht', '17.4')) < need:
                row.set('ht', '%.1f' % need)
                row.set('customHeight', '1')
        # anomalies de dates
        Bv, Cv = it['B'], it['Cv']
        if isinstance(Bv, datetime.datetime) and isinstance(Cv, datetime.datetime) and Cv < Bv:
            report['anomalies_signalees'].append([sh, it['num'], 'POUR LE %s antérieur à ABORDÉ LE %s' % (Cv.strftime('%d/%m/%Y'), Bv.strftime('%d/%m/%Y'))])
        if Bv is None:
            report['anomalies_signalees'].append([sh, it['num'], 'ABORDÉ LE vide' + (' (ligne masquée)' if it['hidden'] else '')])
        if isinstance(Cv, datetime.datetime) and Cv.year >= 2027:
            report['anomalies_signalees'].append([sh, it['num'], 'POUR LE en %s : %s (à confirmer)' % (Cv.year, Cv.strftime('%d/%m/%Y'))])
        if sh == '02 -GROS OEUVRE' and r == 83 and isinstance(Bv, datetime.datetime) and Bv.year == 2027:
            cC = cell_in_row(row, 'C')
            v = cC.find(N + 'v')
            v.text = str((datetime.datetime(2026, 7, 7) - EPOCH).days)
            report['anomalies_corrigees'].append([sh, it['num'], 'ABORDÉ LE 07/07/2027 corrigé en 07/07/2026 (interprétation : faute de frappe sur l\'année, le texte parle du 15/07/2026)'])
        if (sh, r) in DOUBLONS:
            for col in ('A', 'B'):
                c = cell_in_row(row, col)
                c.set('s', xf_clone(c.get('s'), fill_rgb=BLUE_FILL))
            report['doublons'].append([sh, it['num'], 'ligne %d' % r, str(it['A']).strip()[:90]])
        it['new'] = R
    b.touch(b.sheets[sh])

# lignes de données vides nouvelles (réserve, SIEGE 27) : cellules A et F stylées
for sh in MIGR_SHEETS:
    ws = b.ws(sh)
    for row in ws.find(N + 'sheetData'):
        cs = {re.sub(r'\d', '', c.get('r')): c for c in row}
        if int(row.get('r')) in new_data_rows.get(sh, ()) and 'B' in cs and 'A' not in cs:
            cell_in_row(row, 'A', cs['B'].get('s'))
            cell_in_row(row, 'F', cs['B'].get('s'))

# en-tête du tableau SIEGE 27 : N° / HISTORIQUE
sie_rows = info[('MOE-MOA', '_sie_rows')]
row = [r for r in b.ws('MOE-MOA').find(N + 'sheetData') if int(r.get('r')) == sie_rows[0]][0]
sH = cell_in_row(row, 'B').get('s')
for col, txt in (('A', 'N°'), ('F', 'HISTORIQUE')):
    set_cell_si(cell_in_row(row, col, sH), si_new([(None, txt)]), sH)

# ================================================================== 3bis. bordures, jaune manuel retiré, MFC
borders_el = b.styles.find(N + 'borders')
xfs_el = b.styles.find(N + 'cellXfs')
_bcache = {}


def border_variant(bid, drop_left=False, drop_right=False):
    key = (int(bid), drop_left, drop_right)
    if key in _bcache:
        return _bcache[key]
    bd = copy.deepcopy(borders_el[int(bid)])
    for side, drop in (('left', drop_left), ('right', drop_right)):
        if drop:
            e = bd.find(N + side)
            if e is not None:
                for ch in list(e):
                    e.remove(ch)
                e.attrib.pop('style', None)
    borders_el.append(bd)
    borders_el.set('count', str(len(borders_el)))
    _bcache[key] = str(len(borders_el) - 1)
    b.touch('xl/styles.xml')
    return _bcache[key]


_scache = {}


def restyle_cell(c, border_id=None, no_fill=False):
    key = (c.get('s', '0'), border_id, no_fill)
    if key not in _scache:
        xf = copy.deepcopy(xfs_el[int(c.get('s', '0'))])
        if border_id is not None:
            xf.set('borderId', border_id)
            xf.set('applyBorder', '1')
        if no_fill:
            xf.set('fillId', '0')
        xfs_el.append(xf)
        xfs_el.set('count', str(len(xfs_el)))
        b.touch('xl/styles.xml')
        _scache[key] = str(len(xfs_el) - 1)
    c.set('s', _scache[key])


def fill_rgb_of(s):
    fl = b.styles.find(N + 'fills')[int(xfs_el[int(s or 0)].get('fillId', 0))]
    fg = fl.find('.//' + N + 'fgColor')
    return fg.get('rgb') if fg is not None else None


dxfs = b.styles.find(N + 'dxfs')
if dxfs is None:
    dxfs = etree.SubElement(b.styles, N + 'dxfs', count='0')
d_grey = etree.SubElement(dxfs, N + 'dxf')
etree.SubElement(etree.SubElement(d_grey, N + 'font'), N + 'color', theme='0', tint='-0.499984740745262')
d_yel = etree.SubElement(dxfs, N + 'dxf')
etree.SubElement(etree.SubElement(etree.SubElement(d_yel, N + 'fill'), N + 'patternFill', patternType='solid'), N + 'bgColor', rgb='FFFFFF00')
dxfs.set('count', str(len(dxfs)))
id_grey, id_yel = len(dxfs) - 2, len(dxfs) - 1
b.touch('xl/styles.xml')

for sh in MIGR_SHEETS:
    ws = b.ws(sh)
    new_of = new_of_all[sh]
    rowel = {int(r.get('r')): r for r in ws.find(N + 'sheetData')}
    kinds = {}
    for (s2, r), it in info.items():
        if s2 == sh and isinstance(r, int) and r in new_of:
            kinds[new_of[r]] = it['kind']
    for R_ in new_data_rows.get(sh, ()):
        kinds[R_] = 'empty'
    for R_ in info.get((sh, '_sie_rows'), [])[:1]:
        kinds[R_] = 'header'
    obs_rows = sorted(R_ for R_, k in kinds.items() if k == 'obs')
    if obs_rows:
        # gabarit = une observation « courante » (précédée d'une observation) : traits fins haut et bas,
        # comme la 1re ligne d'ÉTUDES ; pas la ligne qui suit un en-tête (trait moyen en haut)
        mid = [R_ for R_ in obs_rows if kinds.get(R_ - 1) == 'obs'] or obs_rows
        tmpl = {c.get('r')[0]: xfs_el[int(c.get('s', '0'))].get('borderId', '0') for c in rowel[mid[0]] if c.get('r')[0] in 'ABCDE'}
    for R_, k in kinds.items():
        row = rowel.get(R_)
        if row is None:
            continue
        for col in 'ABCDE':
            c = cell_in_row(row, col)
            if k in ('obs', 'empty') and obs_rows and col in tmpl:
                bid = border_variant(tmpl[col], drop_left=(col == 'A'), drop_right=(col == 'E'))
            else:
                bid = border_variant(xfs_el[int(c.get('s', '0'))].get('borderId', '0'),
                                     drop_left=(col == 'A'), drop_right=(col == 'E'))
            yellow = k == 'obs' and fill_rgb_of(c.get('s')) == 'FFFFFF00'
            restyle_cell(c, bid, no_fill=yellow)
    # lignes vides des tableaux : masquées pour ne pas les imprimer (réserve conservée, 29/09/2026)
    table_last = {t['new_rows'][-1] for t in info[(sh, '_tables')] if t.get('new_rows')}
    for R_, k in kinds.items():
        if k == 'empty' and R_ in rowel and HIDE_EMPTY and R_ not in table_last:   # la dernière ligne reste visible :
            # c'est elle qui porte le trait épais de bas de tableau (style du tableau)
            rowel[R_].set('hidden', '1')
            report.setdefault('lignes_vides_masquees', 0)
            report['lignes_vides_masquees'] += 1
    # MFC : jaune si URGENT, gris si PM / soldé (décision du 29/09/2026)
    anchor = ws.find(N + 'phoneticPr')
    if anchor is None:
        anchor = ws.find(N + 'mergeCells')
    prio = 1
    for t in info[(sh, '_tables')]:
        if t.get('code') == 'BET.CUI' or len(t.get('new_rows', [])) < 2:
            continue
        r0, r1 = t['new_rows'][1], t['new_rows'][-1]
        cf = etree.Element(N + 'conditionalFormatting', sqref='A%d:E%d' % (r0, r1))
        r1_ = etree.SubElement(cf, N + 'cfRule', type='expression', dxfId=str(id_yel), priority=str(prio))
        etree.SubElement(r1_, N + 'formula').text = '$E%d="URGENT"' % r0
        r2_ = etree.SubElement(cf, N + 'cfRule', type='expression', dxfId=str(id_grey), priority=str(prio + 1))
        etree.SubElement(r2_, N + 'formula').text = ('OR($E{0}="PM",$E{0}="Annulé",$E{0}="Doublon",$E{0}="Sans objet",'
                                                     '$E{0}="Refusé",ISNUMBER($E{0}))').format(r0)
        prio += 2
        anchor.addnext(cf)
        anchor = cf
    b.touch(b.sheets[sh])

# ================================================================== 4. sauts de page, zones d'impression, validations
wbx = b.xml('xl/workbook.xml')
sheet_names = [s.get('name') for s in wbx.find(N + 'sheets')]
dns = wbx.find(N + 'definedNames')


def add_name(name, text, local=None, hidden=False):
    e = etree.SubElement(dns, N + 'definedName', name=name)
    if local is not None:
        e.set('localSheetId', str(local))
    if hidden:
        e.set('hidden', '1')
    e.text = text


for sh in MIGR_SHEETS:
    ws = b.ws(sh)
    tabs = info[(sh, '_tables')]
    if sh == 'Modele_lot':        # onglet masqué : pas de zone d'impression (LibreOffice l'imprimerait)
        continue
    last = max(t['new_rows'][-1] for t in tabs if t.get('new_rows'))      # fin du dernier tableau (pas de pages vides)
    add_name('_xlnm.Print_Area', "'%s'!$A$1:$E$%d" % (sh, last), sheet_names.index(sh))
    if sh in LOT_SHEETS:
        hdr = min(t['new_rows'][0] for t in tabs)
        add_name('_xlnm.Print_Titles', "'%s'!$%d:$%d" % (sh, hdr, hdr), sheet_names.index(sh))
        brks = []
        for (s2, r), it in info.items():
            if s2 == sh and isinstance(r, int) and it['kind'] == 'section' and str(it['A']).strip().upper() == 'TRAVAUX':
                t_end = [t for t in tabs if t['name'] == it['table']][0]['r1']
                visible = any(info[(sh, k)]['kind'] == 'obs' and not info[(sh, k)]['hidden'] for k in range(r + 1, t_end + 1))
                if visible:
                    brks.append(new_of_all[sh][r] - 1)
                else:
                    report.setdefault('sauts_supprimes', []).append([sh, 'TRAVAUX sans ligne visible : pas de saut de page'])
        if brks:
            rb = etree.Element(N + 'rowBreaks', count=str(len(brks)), manualBreakCount=str(len(brks)))
            for k in sorted(brks):
                etree.SubElement(rb, N + 'brk', id=str(k), max='16383', man='1')
            ws.find(N + 'headerFooter').addnext(rb)
    # validation de données sur FAIT LE (liste déroulante, saisie libre autorisée pour les dates)
    sq = ' '.join('E%d:E%d' % (t['new_rows'][1], t['new_rows'][-1]) for t in tabs
                  if t.get('code') != 'BET.CUI' and len(t['new_rows']) > 1)
    old = ws.find(N + 'dataValidations')
    if old is not None:
        ws.remove(old)
    dv = etree.Element(N + 'dataValidations', count='1')
    d1 = etree.SubElement(dv, N + 'dataValidation', type='list', allowBlank='1', showInputMessage='1',
                          showErrorMessage='0', sqref=sq)
    etree.SubElement(d1, N + 'formula1').text = 'ListeStatuts'
    cfs = ws.findall(N + 'conditionalFormatting')
    anchor = cfs[-1] if cfs else None
    if anchor is None:
        anchor = ws.find(N + 'phoneticPr')
    if anchor is None:
        anchor = ws.find(N + 'mergeCells')
    anchor.addnext(dv)
    b.touch(b.sheets[sh])

# ================================================================== 5. Coordonnees : fautes
FIXES_COORD = [('A23', 'DESAMIANRAGE', 'DESAMIANTAGE'), ('A20', 'AOUSTIQUE', 'ACOUSTIQUE'),
               ('H35', 'Esc', 'Exc'), ('A41', 'Esc : Excusé', 'Exc : Excusé')]
for ref, a, z in FIXES_COORD:
    rs = b.runs('Coordonnees', ref)
    rs2 = [(rpr, t.replace(a, z)) for rpr, t in rs]
    if rs2 != rs:
        b.set_runs('Coordonnees', ref, rs2)
        report['anomalies_corrigees'].append(['Coordonnees', ref, '« %s » → « %s »' % (a, z)])

# ================================================================== 6. Page de garde en formules
DD = lambda e: 'RIGHT("0"&DAY(%s),2)&"/"&RIGHT("0"&MONTH(%s),2)&"/"&YEAR(%s)' % (e, e, e)
pg = {'C22': ('"Compte rendu de la réunion de chantier du "&' + DD('B22'),
              'Compte rendu de la réunion de chantier du 22/09/2026'),
      'B26': ('" RDV chantier "&' + DD('B22+7') + '&" à "&B23', ' RDV chantier 29/09/2026 à 9H00')}
for ref, (f, v) in pg.items():
    c = b.cell('Page de garde', ref)
    for ch in list(c):
        c.remove(ch)
    c.set('t', 'str')
    etree.SubElement(c, N + 'f').text = f
    etree.SubElement(c, N + 'v').text = v
b.touch(b.sheets['Page de garde'])

# ================================================================== 7. nouveaux onglets
st_hdr = b.cell('Coordonnees', 'A5', False).get('s')
st_body = b.cell('Coordonnees', 'B6', False).get('s')
st_title = xf_clone(st_body, bold=True)
st_wrap = xf_clone(st_body, wrap=True, valign='top')
nf_date = custom_numfmt('dd/mm/yyyy;;')
st_date = xf_clone(st_body, numfmt=nf_date, halign='center')


def c_xml(ref, value=None, s=None, formula=None, cached=None):
    s_attr = ' s="%s"' % s if s else ''
    if formula is not None:
        v = '' if cached is None else '<v>%s</v>' % cached
        return '<c r="%s"%s><f>%s</f>%s</c>' % (ref, s_attr, _esc(formula), v)
    if value is None:
        return '<c r="%s"%s/>' % (ref, s_attr)
    if isinstance(value, (int, float)):
        return '<c r="%s"%s><v>%s</v></c>' % (ref, s_attr, value)
    return '<c r="%s"%s t="s"><v>%d</v></c>' % (ref, s_attr, si_new([(None, value)]))


def _esc(t):
    return t.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;').replace('"', '&quot;')


def build_sheet(cells, widths, extra_after_sheetdata=''):
    rows = {}
    for (col, r), xml in cells.items():
        rows.setdefault(r, []).append((col_index(col), xml))
    sd = ''.join('<row r="%d">%s</row>' % (r, ''.join(x for _, x in sorted(v))) for r, v in sorted(rows.items()))
    cols = ''.join('<col min="%d" max="%d" width="%s" customWidth="1"/>' % (col_index(c), col_index(c), w) for c, w in widths)
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
            '<worksheet xmlns="%s" xmlns:r="%s"><sheetViews><sheetView workbookViewId="0"/></sheetViews>'
            '<sheetFormatPr defaultRowHeight="15"/><cols>%s</cols><sheetData>%s</sheetData>%s'
            '<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/></worksheet>'
            % (NS_MAIN, RNS, cols, sd, extra_after_sheetdata)).encode('utf-8')


NS_MAIN = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'


def add_sheet(name, xml_bytes, state=None):
    nums = [int(re.sub(r'\D', '', n)) for n in b.data if re.match(r'xl/worksheets/sheet\d+\.xml$', n)]
    part = 'xl/worksheets/sheet%d.xml' % (max(nums) + 1)
    add_part(part, xml_bytes, 'application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml')
    rid = add_rel('xl/_rels/workbook.xml.rels',
                  'http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet',
                  'worksheets/' + part.rsplit('/', 1)[1])
    sheets = wbx.find(N + 'sheets')
    sid = max(int(s.get('sheetId')) for s in sheets) + 1
    e = etree.SubElement(sheets, N + 'sheet', name=name, sheetId=str(sid))
    if state:
        e.set('state', state)
    e.set('{%s}id' % RNS, rid)
    b.sheets[name] = part
    sheet_names.append(name)
    b.touch('xl/workbook.xml')
    return part


# ---- Référentiel
coord = wb_f['Coordonnees']
REF_ROWS = [('MOA', 6), ('SIE', 7), ('AMO', 8), ('CT', 9), ('SPS', 10), ('MOE', 11), ('ECO', 12),
            ('BET.ELE', 13), ('BET.CVC', 14), ('BET.THE', 15), ('BET.VRD', 16), ('BET.SBO', 17),
            ('BET.SBE', 18), ('BET.AMI', 19), ('BET.ACO', 20)] + \
           [('%02d' % k, 22 + k) for k in range(1, 12)] + \
           [('CON.EAU', 36), ('CON.ELE', 37), ('CON.ASS', 38), ('CON.RES', 39)]
sheet_of_code = {'%02d' % k: LOT_SHEETS[k - 1] for k in range(1, 12)}
section_title = {}
for (s2, key), tabs in [(k, v) for k, v in info.items() if k[1] == '_tables']:
    for t in tabs:
        if t.get('code') and s2 in ('MOE-MOA', 'Concessionnaires'):
            sheet_of_code[t['code']] = s2
sheet_of_code['SIE'] = 'MOE-MOA'
cells = {('A', 1): c_xml('A1', 'RÉFÉRENTIEL — onglet de travail masqué, non imprimé (codes, alias, vocabulaire)', st_title)}
cells[('A', 3)] = c_xml('A3', '1. Intervenants', st_title)
for i, h in enumerate(['Code', 'Onglet', 'Organisme', 'Représentants', 'Alias entendus (séparés par ;)', 'Prochain N°']):
    col = 'ABCDEF'[i]
    cells[(col, 4)] = c_xml('%s4' % col, h, st_hdr)
r = 5
for code, cr in REF_ROWS:
    org = coord.cell(cr, 2).value
    reps = coord.cell(cr, 4).value
    reps = '; '.join(x.strip(' *') for x in str(reps).split('\n') if x.strip() and '@' not in x) if reps else ''
    nxt = '%s-%03d' % (code, report['numerotation'].get(code, 0) + 1)
    vals = [code, sheet_of_code.get(code, ''), (org or '').strip() or '—', reps, None, nxt]
    for i, v in enumerate(vals):
        col = 'ABCDEF'[i]
        cells[(col, r)] = c_xml('%s%d' % (col, r), v, st_wrap)
    r += 1
STATUTS = ['Relance', 'URGENT', 'PM', 'En cours', 'En attente', 'En attente MOA', 'En attente SIEGE 27',
           'En attente CICLOP', 'En attente DEKRA', 'En attente VERITAS', 'En attente ACAU', 'En attente ECLA',
           'En attente CONCEPT NF', 'En attente ECHOS', 'En attente AHMES', 'En attente BESB', 'En attente ESGCB',
           'En attente ACCEO', 'En attente GAMBA', 'En attente DEMOLAF', 'En attente AXL', 'En attente AGC',
           'En attente GOUJON VALLEE', 'En attente AVA', 'En attente MCO', 'En attente REVNOR', 'En attente NORDEC',
           'En attente ELAIRGIE', 'En attente BEVELEC', 'En attente CFB TP', 'En attente concessionnaire',
           'En attente ENEDIS', 'En attente ORANGE', 'En attente SRPN', 'En attente SPANC',
           'Annulé', 'Doublon', 'Sans objet', 'Refusé']
cells[('H', 3)] = c_xml('H3', '2. Statuts FAIT LE (liste fermée)', st_title)
cells[('H', 4)] = c_xml('H4', 'Statut (ou une date)', st_hdr)
for i, v in enumerate(STATUTS):
    cells[('H', 5 + i)] = c_xml('H%d' % (5 + i), v, st_body)
add_name('ListeStatuts', "'Référentiel'!$H$5:$H$%d" % (4 + len(STATUTS)))
ZONES = ['préau', 'SHED', 'bâtiment A', 'chaufferie', 'restaurant scolaire', 'école existante', 'extension',
         'mairie', 'cours anglaises', 'vide sanitaire', 'base vie', 'salle polyvalente', 'maternelle',
         'circulation', 'cuisine', 'self', 'pignon']
LEX = ['RSD', 'longrines', 'hourdis', 'courettes', 'acodrains', 'MOB', 'tebopins', 'pré-isolé', 'soubassement',
       'rejingot', 'couvertine', 'précadre', 'claire-voie', 'citerneau', 'ANC', 'AEP', 'arbalétrier', 'bac',
       'noue', 'tranchée commune', 'dallage', 'sous-œuvre', 'enduit', 'cloisonnettes', 'CTA', 'DAS', 'BPE', 'TEAMS',
       'prorata']
SIGLES = [('DT', 'DICT'), ('FT', 'PT'), ('CVP', 'CVC'), ('VISA', 'VIC'), ('EXE', 'DESC'), ('PPSPS', 'PGC'),
          ('CFO-CFA', '—')]
for col0, title, items in (('J', '3. Zones et ouvrages', ZONES), ('M', '4. Lexique du chantier', LEX)):
    col1 = col_letter(col_index(col0) + 1)
    cells[(col0, 3)] = c_xml('%s3' % col0, title, st_title)
    cells[(col0, 4)] = c_xml('%s4' % col0, 'Terme exact', st_hdr)
    cells[(col1, 4)] = c_xml('%s4' % col1, 'Variantes entendues (séparées par ;)', st_hdr)
    for i, v in enumerate(items):
        cells[(col0, 5 + i)] = c_xml('%s%d' % (col0, 5 + i), v, st_body)
        cells[(col1, 5 + i)] = c_xml('%s%d' % (col1, 5 + i), None, st_body)
cells[('P', 3)] = c_xml('P3', '5. Sigles faciles à confondre à l\'oral', st_title)
cells[('P', 4)] = c_xml('P4', 'Sigle', st_hdr)
cells[('Q', 4)] = c_xml('Q4', 'Confondu avec', st_hdr)
for i, (a, z) in enumerate(SIGLES):
    cells[('P', 5 + i)] = c_xml('P%d' % (5 + i), a, st_body)
    cells[('Q', 5 + i)] = c_xml('Q%d' % (5 + i), z, st_body)
add_sheet('Référentiel', build_sheet(cells, [('A', 10), ('B', 26), ('C', 22), ('D', 42), ('E', 34), ('F', 12),
                                             ('H', 26), ('J', 22), ('K', 30), ('M', 20), ('N', 30), ('P', 10), ('Q', 16)]),
          state='hidden')

# ---- Points à traiter
all_tabs = []
for sh in OBS_SHEETS:
    for t in sorted(info[(sh, '_tables')], key=lambda t: t['new_rows'][0] if t.get('new_rows') else 0):
        if t.get('code') != 'BET.CUI':
            all_tabs.append(t['name'])
add_name('TousLesPoints', '_xlfn.VSTACK(%s)' % ','.join('%s[#Data]' % n for n in all_tabs))
cells = {('A', 1): c_xml('A1', 'POINTS À TRAITER — onglet interne masqué, exclu du PDF', st_title),
         ('A', 2): c_xml('A2', 'Comptage : formules classiques. Les 3 blocs du bas sont à compléter en collant les formules du rapport.', st_body)}
heads = ['Onglet', 'Observations', 'URGENT', 'Relance', 'En attente', 'PM', 'Soldées (date)']
for i, h in enumerate(heads):
    cells[('ABCDEFG'[i], 4)] = c_xml('%s4' % 'ABCDEFG'[i], h, st_hdr)
r = 5
for sh in OBS_SHEETS:
    q = "'%s'" % sh
    fs = ['COUNTIF(%s!$A:$A,"*-???")' % q, 'COUNTIF(%s!$E:$E,"URGENT")' % q, 'COUNTIF(%s!$E:$E,"Relance")' % q,
          'COUNTIF(%s!$E:$E,"En attente*")' % q, 'COUNTIF(%s!$E:$E,"PM")' % q, 'COUNT(%s!$E:$E)' % q]
    cells[('A', r)] = c_xml('A%d' % r, sh, st_body)
    for i, f in enumerate(fs):
        col = 'BCDEFG'[i]
        cells[(col, r)] = c_xml('%s%d' % (col, r), s=st_body, formula=f)
    r += 1
cells[('A', r)] = c_xml('A%d' % r, 'TOTAL', st_title)
for col in 'BCDEFG':
    cells[(col, r)] = c_xml('%s%d' % (col, r), s=st_title, formula='SUM(%s5:%s%d)' % (col, col, r - 1))
top = r + 3
for k, (col0, title) in enumerate((('A', 'BLOC 1 — URGENT'), ('F', 'BLOC 2 — ÉCHÉANCE DÉPASSÉE'), ('K', 'BLOC 3 — EN ATTENTE'))):
    ci0 = col_index(col0)
    cells[(col0, top)] = c_xml('%s%d' % (col0, top), title, st_title)
    for i, h in enumerate(['N°', 'Observation (120 car.)', 'Pour le', 'Fait le']):
        col = col_letter(ci0 + i)
        cells[(col, top + 1)] = c_xml('%s%d' % (col, top + 1), h, st_hdr)
    cells[(col0, top + 2)] = c_xml('%s%d' % (col0, top + 2), 'Coller ici la formule %d du rapport' % (k + 1), st_body)
    for rr in range(top + 2, top + 202):          # format date (vide si 0) pour POUR LE / FAIT LE
        for dc in (2, 3):
            col = col_letter(ci0 + dc)
            cells[(col, rr)] = c_xml('%s%d' % (col, rr), None, st_date)
add_sheet('Points à traiter', build_sheet(cells, [('A', 22), ('B', 60), ('C', 11), ('D', 18), ('F', 11), ('G', 60),
                                                  ('H', 11), ('I', 18), ('K', 11), ('L', 60), ('M', 11), ('N', 18)]),
          state='hidden')
report['points_a_traiter'] = dict(ligne_blocs=top + 2, tables=all_tabs)

# ---- Test MFC (validé le 29/09/2026 : onglet retiré de la copie finale)
ADD_TEST_MFC = False
dxfs = b.styles.find(N + 'dxfs')
d_grey = etree.SubElement(dxfs, N + 'dxf')
f_ = etree.SubElement(d_grey, N + 'font')
etree.SubElement(f_, N + 'color', theme='0', tint='-0.499984740745262')
d_yel = etree.SubElement(dxfs, N + 'dxf')
fl = etree.SubElement(d_yel, N + 'fill')
pf = etree.SubElement(fl, N + 'patternFill', patternType='solid')
etree.SubElement(pf, N + 'bgColor', rgb='FFFFFF00')
dxfs.set('count', str(len(dxfs)))
id_grey, id_yel = len(dxfs) - 2, len(dxfs) - 1
b.touch('xl/styles.xml')
body = b.cell('05 - MENUISERIES EXT.', 'B7', False).get('s')    # style réel d'une observation
blk = xf_clone(body, color=BLACK)
redrpr = _rpr_from_font(b._font_of(blk))
_set_color(redrpr, RED)
blkrpr = _rpr_from_font(b._font_of(blk))
_set_color(blkrpr, BLACK)
st_d = xf_clone(b.cell('05 - MENUISERIES EXT.', 'C7', False).get('s'), color=BLACK)
serial = lambda d: (d - EPOCH).days
TESTS = [('T-01', 'Ligne PM : historique noir. ', 'Au 22/09/2026 ajout du jour en rouge.', serial(datetime.datetime(2026, 6, 2)), None, 'PM',
          'Historique GRIS ; ajout du jour ROUGE ?'),
         ('T-02', 'Ligne soldée (date en FAIT LE). ', 'Au 22/09/2026 ajout du jour en rouge.', serial(datetime.datetime(2026, 6, 9)), None, serial(datetime.datetime(2026, 9, 22)),
          'Historique GRIS ; ajout ROUGE ?'),
         ('T-03', 'Ligne URGENT. ', 'Au 22/09/2026 Relance.', serial(datetime.datetime(2026, 5, 5)), serial(datetime.datetime(2026, 5, 19)), 'URGENT',
          'Fond JAUNE ; texte NOIR + ROUGE ?'),
         ('T-04', 'Ligne En attente (ex-PM). ', 'Au 22/09/2026 ajout en rouge.', serial(datetime.datetime(2026, 6, 16)), None, 'En attente MOA',
          'NI gris NI jaune ?'),
         ('T-05', 'Ligne PM entièrement noire, sans ajout du jour.', None, serial(datetime.datetime(2026, 5, 5)), None, 'PM',
          'Tout GRIS ?'),
         ('T-06', 'Ligne active normale. ', 'Au 22/09/2026 ajout en rouge.', serial(datetime.datetime(2026, 9, 15)), serial(datetime.datetime(2026, 9, 29)), None,
          'Aucun changement (noir + rouge)')]
cells = {('A', 1): c_xml('A1', 'TEST — mise en forme conditionnelle (gris si PM ou soldé, jaune si URGENT). À supprimer après le test.', st_title)}
for i, h in enumerate(['N°', 'OBSERVATIONS', 'ABORDÉ LE', 'POUR LE', 'FAIT LE', 'CE QUE VOUS DEVEZ VOIR']):
    cells[('ABCDEF'[i], 3)] = c_xml('%s3' % 'ABCDEF'[i], h, st_hdr)
for k, (num, h, add, d1, d2, fe, exp) in enumerate(TESTS):
    r = 4 + k
    cells[('A', r)] = c_xml('A%d' % r, num, blk)
    runs = [(blkrpr, h)] + ([(redrpr, add)] if add else [])
    cells[('B', r)] = '<c r="B%d" s="%s" t="s"><v>%d</v></c>' % (r, blk, si_new(runs))
    cells[('C', r)] = c_xml('C%d' % r, d1, st_d)
    cells[('D', r)] = c_xml('D%d' % r, d2, st_d) if d2 else c_xml('D%d' % r, None, st_d)
    cells[('E', r)] = c_xml('E%d' % r, fe, st_d if isinstance(fe, int) else blk) if fe is not None else c_xml('E%d' % r, None, blk)
    cells[('F', r)] = c_xml('F%d' % r, exp, st_wrap)
cf = ('<conditionalFormatting sqref="A4:E9"><cfRule type="expression" dxfId="%d" priority="1"><formula>$E4="URGENT"</formula></cfRule>'
      '<cfRule type="expression" dxfId="%d" priority="2"><formula>OR($E4="PM",$E4="Annulé",$E4="Doublon",$E4="Sans objet",$E4="Refusé",ISNUMBER($E4))</formula></cfRule>'
      '</conditionalFormatting>' % (id_yel, id_grey))
if ADD_TEST_MFC:
  add_sheet('Test MFC', build_sheet(cells, [('A', 9.6), ('B', 53), ('C', 11), ('D', 11), ('E', 16), ('F', 34)], cf))

wbx.find(N + 'calcPr').set('fullCalcOnLoad', '1')
b.touch('xl/workbook.xml')
b.save(DST)
json.dump(report, open(REPORT, 'w'), ensure_ascii=False, indent=1, default=str)
print('OK', DST)
