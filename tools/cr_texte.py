# -*- coding: utf-8 -*-
"""Règles de texte du CR (communes à la migration et à l'intégration hebdomadaire).

Voir docs/decisions-HONGUEMARE.md : structure « sujet / → Au JJ/MM/AAAA : … / → Relancé N fois,
dernière le … (sans réponse depuis N j) », règle A, « [...] », relances sans « Au » supprimées."""
import copy
import datetime
import re

from lxml import etree
from cr_xml import N

# ------------------------------------------------------------------ texte : typos et condensation
TYPO = re.compile(r'\b(\d{2})/(\d{2})(20\d{2})\b')
TYPO5 = re.compile(r'\b(\d{2})/(\d{2})/(202)(\d)(\d)\b')   # « 08/09/20256 » -> 2026 (interprétation)
# « Au » / « AU » seulement : un « au » minuscule est du texte (« visite au 20/07/2026 »), pas une mise à jour
AU = re.compile(r'(?:(?<=\s)|(?<=\.)|^)(?:Au|AU)\s+(\d{1,2}/\d{1,2}/(?:\d{4}|\d{2}))(?![\d])')
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


def condense(runs, abord=None, open_=True, date_cr=None):
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
        b0, b1 = en, stop                          # bornes du corps dans le texte (couleurs des runs conservées)
        while b0 < b1 and text[b0].isspace():
            b0 += 1
        while b0 < b1 and text[b0] == ':':
            b0 += 1
        while b0 < b1 and text[b0].isspace():
            b0 += 1
        while b1 > b0 and text[b1 - 1].isspace():
            b1 -= 1
        pieces.append((st, dt, body, b0, b1))
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
            st, dt, body, b0, b1 = e
            # corps repris run par run : un ajout en rouge au milieu d'une remarque reste rouge (02/10/2026)
            out.append((plain_rpr(rpr_at(runs, st)), '\n→ Au %s : ' % full_date(dt)))
            out += [(plain_rpr(rp), t) for rp, t in slice_runs(runs, b0, b1)]
        else:
            last = e[-1]
            txt = '\n→ Relancé %d fois, dernière le %s' % (len(e), full_date(last[1]))
            if open_:                              # fait objectif : jours calendaires depuis la 1re relance sans réponse
                d0 = datetime.datetime.strptime(full_date(e[0][1]), '%d/%m/%Y')
                n = (date_cr - d0).days
                if n > 0:                          # rien le jour même de la 1re relance
                    txt += ' (sans réponse depuis %d j)' % n
            out.append((plain_rpr(rpr_at(runs, last[0])), txt))
    # Sujet initial en gras ; lignes suivantes sans gras, couleurs de la source
    return out, len(bare), bare[-1][1] if bare else None, mismatch


