# -*- coding: utf-8 -*-
"""Intégration hebdomadaire d'un résumé routé dans un CR au nouveau format (colonne N°, HISTORIQUE…).

Usage : python integrer_cr.py CR_PRECEDENT.xlsx OPERATIONS.json CR_NOUVEAU.xlsx RAPPORT.json
puis    python ajuster_hauteurs.py CR_NOUVEAU.xlsx CR_FINAL.xlsx

OPERATIONS.json :
{
  "date_cr": "2026-09-29", "crc": "CRC-16",
  "operations": [
    {"type": "maj", "num": "02-066", "texte": "…", "pour_le": "2026-10-02" | "+7", "fait_le": "URGENT" | "2026-09-22"},
    {"type": "nouvelle", "onglet": "02 -GROS OEUVRE", "code": "02", "section": "ÉTUDES",
     "texte": "…", "pour_le": "+7" | "2026-10-06" | null, "fait_le": "PM" | null}
  ]
}

Règles appliquées : docs/decisions-HONGUEMARE.md. Édition XML chirurgicale (cr_xml.py) ; openpyxl en
lecture seule. Le texte affiché (OBSERVATIONS) est recalculé depuis HISTORIQUE, qui fait foi.
"""
import copy
import datetime
import json
import os
import re
import sys

from lxml import etree

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from cr_xml import Book, N, RED, BLACK, GREY, EPOCH, _set_color, _rpr_from_font, color_of, col_index  # noqa: E402
from cr_texte import condense, plain_rpr  # noqa: E402

BASE, OPS, OUT, REPORT = sys.argv[1:5]
spec = json.load(open(OPS, encoding='utf-8'))
DATE = datetime.datetime.strptime(spec['date_cr'], '%Y-%m-%d')
DSTR = DATE.strftime('%d/%m/%Y')
TERMINAUX = {'ANNULÉ', 'DOUBLON', 'SANS OBJET', 'REFUSÉ'}
CODE_RE = re.compile(r'^([A-Z0-9.]+)-(\d{3})$')
SECTION_CODES = [("maitre d'ouvrage_ville", 'MOA'), ('siege 27', 'SIE'), ('amo_', 'AMO'), ('bureau de contr', 'CT'),
                 ('csps_', 'SPS'), ("maitrise d'oeuvre", 'MOE'), ('economiste', 'ECO'), ('bet cvc', 'BET.CVC'),
                 ('bet électricité', 'BET.ELE'), ('bet thermique', 'BET.THE'), ('bet vrd', 'BET.VRD'),
                 ('structure béton', 'BET.SBE'), ('structure bois', 'BET.SBO'), ('accoustique', 'BET.ACO'),
                 ('amiante', 'BET.AMI'), ('concessionaires_réseaux', 'CON.RES'), ('concessionaires_eaux', 'CON.EAU'),
                 ('concessionaires_electricite', 'CON.ELE'), ('concessionaires_assainissement', 'CON.ASS')]
report = {'masquees': [], 'maj': [], 'nouvelles': [], 'relances_auto': [], 'reouvertures': [], 'insertions': [],
          'alertes': []}

b = Book(BASE)
b.formulas_added = True
_xf = {}


# ------------------------------------------------------------------ outils bas niveau
def restyle(c, color=None, bold=None):
    key = (c.get('s', '0'), repr(color), bold)
    if key not in _xf:
        st = b.styles
        fonts, xfs = st.find(N + 'fonts'), st.find(N + 'cellXfs')
        xf = copy.deepcopy(xfs[int(c.get('s', '0'))])
        f = copy.deepcopy(fonts[int(xf.get('fontId', 0))])
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
        xfs.append(xf)
        xfs.set('count', str(len(xfs)))
        b.touch('xl/styles.xml')
        _xf[key] = str(len(xfs) - 1)
    c.set('s', _xf[key])


def si_new(runs):
    si = etree.SubElement(b.sst, N + 'si')
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


def clear(c):
    for ch in list(c):
        c.remove(ch)
    c.attrib.pop('t', None)


def set_si(c, idx):
    clear(c)
    c.set('t', 's')
    etree.SubElement(c, N + 'v').text = str(idx)


def set_num(c, v):
    clear(c)
    etree.SubElement(c, N + 'v').text = str(v)


def set_formula(c, f, cached=None):
    clear(c)
    etree.SubElement(c, N + 'f').text = f
    if cached is not None:
        etree.SubElement(c, N + 'v').text = str(cached)


def cell(row, col):
    ref = '%s%s' % (col, row.get('r'))
    ci = col_index(ref)
    for c in row.iter(N + 'c'):
        if c.get('r') == ref:
            return c
        if col_index(c.get('r')) > ci:
            new = etree.Element(N + 'c', r=ref)
            c.addprevious(new)
            return new
    return etree.SubElement(row, N + 'c', r=ref)


def text_of(c):
    if c is None:
        return ''
    if c.get('t') == 's':
        return ''.join(b.si[int(c.find(N + 'v').text)].itertext())
    if c.get('t') == 'str':
        v = c.find(N + 'v')
        return v.text or '' if v is not None else ''
    v = c.find(N + 'v')
    return v.text if v is not None and v.text else ''


def runs_of(c):
    base = b._font_of(c.get('s'))
    if c.get('t') != 's':
        t = text_of(c)
        return [(_rpr_from_font(base), t)] if t else []
    si = b.si[int(c.find(N + 'v').text)]
    rs = si.findall(N + 'r')
    if not rs:
        t = si.find(N + 't')
        return [(_rpr_from_font(base), t.text or '')] if t is not None and t.text else []
    out = []
    for r in rs:
        rpr = r.find(N + 'rPr')
        out.append((copy.deepcopy(rpr) if rpr is not None else _rpr_from_font(base), r.find(N + 't').text or ''))
    return out


def number_of(c):
    """Valeur numérique (date = numéro de série) d'une cellule, formule comprise (valeur en cache)."""
    if c is None or c.get('t') in ('s', 'str', 'inlineStr'):
        return None
    v = c.find(N + 'v')
    try:
        return float(v.text) if v is not None and v.text else None
    except ValueError:
        return None


def serial(d):
    return (d - EPOCH).days


def from_serial(n):
    return EPOCH + datetime.timedelta(days=int(n))


def status_runs(c, value):
    rpr = plain_rpr(_rpr_from_font(b._font_of(c.get('s'))))
    _set_color(rpr, GREY if value.upper() == 'PM' else RED)
    return [(rpr, value)]


def write_fait_le(c, value):
    """Statut en texte enrichi rouge (PM gris) ; date en rouge (clôture de la semaine)."""
    if re.match(r'^\d{4}-\d{2}-\d{2}$', value):
        set_num(c, serial(datetime.datetime.strptime(value, '%Y-%m-%d')))
        restyle(c, color=RED)
    else:
        set_si(c, si_new(status_runs(c, value)))


def write_pour_le(c, value, row_r):
    if value is None:
        return
    if value.startswith('+'):
        set_num(c, serial(DATE + datetime.timedelta(days=int(value[1:]))))
    else:
        set_num(c, serial(datetime.datetime.strptime(value, '%Y-%m-%d')))
    restyle(c, color=RED)


# ------------------------------------------------------------------ lecture du modèle
def obs_sheets():
    out = []
    for name, part in b.sheets.items():
        ws = b.ws(name)
        for row in ws.find(N + 'sheetData'):
            if int(row.get('r')) > 12:
                break
            for c in row:
                if c.get('r').startswith('A') and text_of(c).strip() == 'N°':
                    out.append(name)
                    break
            if out and out[-1] == name:
                break
    return out


def resolve(base_part, target):
    d = base_part.rsplit('/', 1)[0]
    out = []
    for x in (d + '/' + target).split('/'):
        out.pop() if x == '..' else out.append(x)
    return '/'.join(out)


def tables_of(sh):
    part = b.sheets[sh]
    d, f = part.rsplit('/', 1)
    res = []
    for r in b.xml('%s/_rels/%s.rels' % (d, f)):
        if r.get('Type').endswith('/table'):
            tp = resolve(part, r.get('Target'))
            t = b.xml(tp)
            a, z = t.get('ref').split(':')
            res.append(dict(part=tp, name=t.get('name'), r0=int(re.sub(r'\D', '', a)), r1=int(re.sub(r'\D', '', z))))
    return sorted(res, key=lambda x: x['r0'])


def model(sh):
    ws = b.ws(sh)
    rows = {int(r.get('r')): r for r in ws.find(N + 'sheetData')}
    tabs = tables_of(sh)
    info = {}
    for t in tabs:
        code, sec = None, None
        for rr in range(t['r0'] + 1, t['r1'] + 1):
            a = text_of(rows[rr].find(N + 'c[@r="A%d"]' % rr)) if rr in rows else ''
            m = CODE_RE.match(a.strip())
            if m:
                code = m.group(1)
                break
        if code is None:                          # tableau sans observation : titre au-dessus du tableau
            for rr in range(t['r0'] - 1, 0, -1):
                if rr in rows:
                    tx = ' '.join(text_of(c) for c in rows[rr]).strip().lower()
                    if tx:
                        code = next((cd for k, cd in SECTION_CODES if k in tx), None)
                        break
        t['code'] = code
        for rr in range(t['r0'] + 1, t['r1'] + 1):
            row = rows.get(rr)
            a = text_of(row.find(N + 'c[@r="A%d"]' % rr)).strip() if row is not None else ''
            bb = text_of(row.find(N + 'c[@r="B%d"]' % rr)).strip().upper() if row is not None else ''
            if bb in ('ÉTUDES', 'TRAVAUX'):
                sec = bb
                info[rr] = dict(kind='section', table=t['name'], section=sec)
            elif CODE_RE.match(a):
                info[rr] = dict(kind='obs', table=t['name'], section=sec, num=a,
                                hidden=row.get('hidden') in ('1', 'true'))
            else:
                info[rr] = dict(kind='empty', table=t['name'], section=sec,
                                hidden=row is not None and row.get('hidden') in ('1', 'true'))
        info[t['r0']] = dict(kind='header', table=t['name'])
    return rows, tabs, info


# ------------------------------------------------------------------ insertion de lignes
A1 = re.compile(r"(?<![A-Za-z_\[\]!'\.])(\$?)([A-Z]{1,3})(\$?)(\d+)(?![\d(A-Za-z_])")


def shift_txt(txt, p, n):
    return A1.sub(lambda m: '%s%s%s%d' % (m.group(1), m.group(2), m.group(3),
                                         int(m.group(4)) + (n if int(m.group(4)) >= p else 0)), txt)


def insert_rows(sh, p, n, tmpl_r):
    """Insère n lignes vierges (style de la ligne tmpl_r) à la position p ; décale tout ce qui suit."""
    ws = b.ws(sh)
    sd = ws.find(N + 'sheetData')
    rows = list(sd)
    tmpl = [r for r in rows if int(r.get('r')) == tmpl_r][0]
    for row in reversed(rows):
        r = int(row.get('r'))
        if r >= p:
            row.set('r', str(r + n))
            for c in row.iter(N + 'c'):
                c.set('r', re.sub(r'\d+', '', c.get('r')) + str(r + n))
    for f in ws.iter(N + 'f'):
        if f.text:
            f.text = shift_txt(f.text, p, n)
    after = next((r for r in sd if int(r.get('r')) == p + n), None)
    for k in range(n):
        new = copy.deepcopy(tmpl)
        new.set('r', str(p + k))
        for a in ('hidden', 'ht', 'customHeight', 'spans'):
            new.attrib.pop(a, None)
        for c in new:
            c.set('r', re.sub(r'\d+', '', c.get('r')) + str(p + k))
            clear(c)
        if after is not None:
            after.addprevious(new)
        else:
            sd.append(new)
    dim = ws.find(N + 'dimension')
    if dim is not None:
        dim.set('ref', shift_txt(dim.get('ref'), p, n))
    for mc in ws.iter(N + 'mergeCell'):
        mc.set('ref', shift_txt(mc.get('ref'), p, n))
    for cf in ws.findall(N + 'conditionalFormatting'):
        cf.set('sqref', shift_txt(cf.get('sqref'), p, n))
    for dv in ws.iter(N + 'dataValidation'):
        dv.set('sqref', shift_txt(dv.get('sqref'), p, n))
    rb = ws.find(N + 'rowBreaks')
    if rb is not None:
        for brk in rb:
            if int(brk.get('id')) + 1 >= p:
                brk.set('id', str(int(brk.get('id')) + n))
    for t in tables_of(sh):
        tx = b.xml(t['part'])
        tx.set('ref', shift_txt(tx.get('ref'), p, n))
        af = tx.find(N + 'autoFilter')
        if af is not None:
            af.set('ref', tx.get('ref'))
        b.touch(t['part'])
    wbx = b.xml('xl/workbook.xml')
    for dn in wbx.iter(N + 'definedName'):
        if ("'%s'!" % sh) in (dn.text or ''):
            dn.text = shift_txt(dn.text.replace("'%s'!" % sh, '@@'), p, n).replace('@@', "'%s'!" % sh)
            b.touch('xl/workbook.xml')
    b.touch(b.sheets[sh])
    report['insertions'].append([sh, p, n])


# ================================================================== 1. page de garde
pg = b.ws('Page de garde')
for row in pg.find(N + 'sheetData'):
    if row.get('r') == '22':
        set_si(cell(row, 'A'), si_new([(_rpr_from_font(b._font_of(cell(row, 'A').get('s'))), spec['crc'])]))
        set_num(cell(row, 'B'), serial(DATE))
b.touch(b.sheets['Page de garde'])

SHEETS = obs_sheets()
ops = spec['operations']
touched = {o['num'] for o in ops if o['type'] == 'maj'}

# ================================================================== 2. place pour les lignes neuves (insertions)
for sh in SHEETS:
    news = [o for o in ops if o['type'] == 'nouvelle' and o['onglet'] == sh]
    if not news:
        continue
    groups = {}
    for o in news:
        groups.setdefault((o.get('code'), o.get('section')), []).append(o)
    rows, tabs, info = model(sh)
    plans = []
    for (code, sec), lst in groups.items():
        cand = [t for t in tabs if (code is None or t['code'] == code)]
        if not cand:
            raise SystemExit('Section introuvable : %s %s %s' % (sh, code, sec))
        t = cand[0]
        span = [r for r in range(t['r0'] + 1, t['r1'] + 1) if info[r].get('section') == sec or sec is None]
        span = [r for r in span if info[r]['kind'] != 'section']
        free = [r for r in span if info[r]['kind'] == 'empty' and r != t['r1']]
        need = len(lst) + 2 - len(free)                   # + réserve de 2 lignes vides
        if need > 0:
            last = max(span)
            p = t['r1'] if last == t['r1'] else last + 1  # avant la dernière ligne du tableau, ou avant la section suivante
            mids = [r for r in span if info[r]['kind'] == 'obs'] or span
            plans.append((p, need, mids[-1]))
    for p, n, tmpl in sorted(plans, reverse=True):
        insert_rows(sh, p, n, tmpl)

# ================================================================== 3. masquage, rouge -> noir, mises à jour, lignes neuves, relances
next_num = {}
for sh in SHEETS:
    rows, tabs, info = model(sh)
    for r, it in info.items():
        if it['kind'] == 'obs':
            m = CODE_RE.match(it['num'])
            next_num[m.group(1)] = max(next_num.get(m.group(1), 0), int(m.group(2)))

pending_new = [o for o in ops if o['type'] == 'nouvelle']
for sh in SHEETS:
    rows, tabs, info = model(sh)
    by_num = {it['num']: r for r, it in info.items() if it['kind'] == 'obs'}
    for r, it in sorted(info.items()):
        if it['kind'] != 'obs':
            continue
        row = rows[r]
        num = it['num']
        cB, cC, cD, cE, cF = (cell(row, x) for x in 'BCDEF')
        e_txt = text_of(cE).strip()
        e_num = number_of(cE)
        closed = e_num is not None or e_txt.upper() in TERMINAUX
        # -- masquage des lignes soldées lors d'un passage précédent (jamais une ligne traitée ce jour, jamais PM)
        if closed and not it['hidden'] and num not in touched:
            row.set('hidden', '1')
            report['masquees'].append([sh, num, from_serial(e_num).strftime('%d/%m/%Y') if e_num is not None else e_txt])
            continue
        # -- rouge du passage précédent -> noir (HISTORIQUE, dates ABORDÉ / POUR LE)
        # (le gris écrit dans le texte est aussi ramené au noir : c'est la MFC qui grise les lignes PM / soldées,
        #  et une ligne « En attente » ne doit jamais être grise — règle V1 ; test MFC validé le 29/09/2026)
        f_runs = runs_of(cF)
        if any(color_of(rp) in (RED, 'GREY') for rp, _ in f_runs):
            for rp, _ in f_runs:
                if color_of(rp) in (RED, 'GREY'):
                    _set_color(rp, BLACK)
            set_si(cF, si_new(f_runs))
        for c in (cC, cD):
            if c.get('s') and color_of(_rpr_from_font(b._font_of(c.get('s')))) == RED:
                restyle(c, color=BLACK)
        if e_num is not None and color_of(_rpr_from_font(b._font_of(cE.get('s')))) == RED:
            restyle(cE, color=BLACK)
        # -- mise à jour du jour
        op = next((o for o in ops if o['type'] == 'maj' and o['num'] == num), None)
        add = None
        if op:
            add = op['texte'].strip()
            if op.get('pour_le'):
                write_pour_le(cD, op['pour_le'], r)
            if op.get('fait_le'):
                write_fait_le(cE, op['fait_le'])
            elif e_num is not None:                    # réouverture (règle V1) : point annoté alors qu'il était soldé
                clear(cE)
                report['reouvertures'].append([sh, num])
            report['maj'].append([sh, num, add[:90]])
        elif not it['hidden'] and not closed:
            # -- relance automatique (échéance atteinte ; jamais sur PM, En attente, En cours, Retard)
            up = e_txt.upper()
            pl = number_of(cD)
            due = pl is not None and pl <= serial(DATE)
            if up == 'URGENT' or (up in ('', 'RELANCE') and due):
                add = 'Relance'
                if up == '':
                    write_fait_le(cE, 'Relance')
                report['relances_auto'].append([sh, num, up or '(vide)'])
        if add:
            f_runs = runs_of(cF)
            base = copy.deepcopy(f_runs[-1][0]) if f_runs else _rpr_from_font(b._font_of(cF.get('s')))
            base = plain_rpr(base)
            _set_color(base, RED)
            if f_runs and not f_runs[-1][1].rstrip().endswith(('.', ':', ';')):
                f_runs[-1] = (f_runs[-1][0], f_runs[-1][1].rstrip() + '.')     # le point prend la couleur du texte précédent
            f_runs.append((base, ' Au %s %s' % (DSTR, add)))
            set_si(cF, si_new(f_runs))
        # -- texte affiché recalculé depuis HISTORIQUE (règle A, « sans réponse depuis », objet d'attente conservé)
        old_b = runs_of(cB)
        objet = [(rp, t) for rp, t in old_b if t.startswith('\n→ En attente :')]
        ab = number_of(cC)
        e_now = text_of(cE).strip().upper()
        open_ = number_of(cE) is None and e_now not in TERMINAUX | {'PM'}
        res = condense(runs_of(cF), from_serial(ab) if ab else None, open_, date_cr=DATE)
        if res:
            runs = res[0]
            if objet and 'EN ATTENTE' in e_now:
                for rp, _ in objet:
                    if color_of(rp) in (RED, 'GREY'):
                        _set_color(rp, BLACK)
                runs += objet
            set_si(cB, si_new(runs))
    # -- lignes neuves de cet onglet
    rows, tabs, info = model(sh)
    for o in [o for o in pending_new if o['onglet'] == sh]:
        t = [t for t in tabs if o.get('code') is None or t['code'] == o.get('code')][0]
        span = [r for r in range(t['r0'] + 1, t['r1'] + 1)
                if info[r]['kind'] == 'empty' and r != t['r1'] and (o.get('section') is None or info[r].get('section') == o['section'])]
        r = span[0]
        info[r]['kind'] = 'obs'
        row = rows[r]
        row.attrib.pop('hidden', None)
        code = o.get('code') or t['code']
        next_num[code] = next_num.get(code, 0) + 1
        num = '%s-%03d' % (code, next_num[code])
        tmpl_r = max([rr for rr, it in info.items() if it['kind'] == 'obs' and it.get('table') == t['name'] and rr != r] or [r])
        for col in 'ABCDEFGHI':
            src = rows[tmpl_r].find(N + 'c[@r="%s%d"]' % (col, tmpl_r))
            c = cell(row, col)
            if src is not None and src.get('s'):
                c.set('s', src.get('s'))
            clear(c)
        cA, cB, cC, cD, cE, cF, cH, cI = (cell(row, x) for x in 'ABCDEFHI')
        set_si(cA, si_new([(_rpr_from_font(b._font_of(cA.get('s'))), num)]))
        restyle(cA, color=BLACK)
        rpr = plain_rpr(_rpr_from_font(b._font_of(cB.get('s'))))
        _set_color(rpr, RED)
        set_si(cB, si_new([(rpr, o['texte'].strip())]))
        set_si(cF, si_new([(rpr, o['texte'].strip())]))
        set_num(cC, serial(DATE))
        restyle(cC, color=RED)
        pl = o.get('pour_le')
        if pl and pl.startswith('+'):
            set_formula(cD, 'C%d+%d' % (r, int(pl[1:])), serial(DATE) + int(pl[1:]))
            restyle(cD, color=RED)
        elif pl:
            write_pour_le(cD, pl, r)
        if o.get('fait_le'):
            write_fait_le(cE, o['fait_le'])
        set_si(cH, si_new([(None, o.get('section') or code)]))
        set_formula(cI, ('IF($E{0}="Retard",IF(ISNUMBER($G{0}),MAX(0,dateCR-$G{0}),""),'
                         'IF(AND(ISNUMBER($G{0}),ISNUMBER($E{0})),MAX(0,$E{0}-$G{0}),""))').format(r))
        cI.set('t', 'str')
        etree.SubElement(cI, N + 'v').text = ''
        report['nouvelles'].append([sh, num, o.get('section') or code, o['texte'][:90]])
    # -- lignes vides des tableaux : masquées (réserve), sauf la dernière de chaque tableau ; formule JOURS RETARD
    rows, tabs, info = model(sh)
    lasts = {t['r1'] for t in tabs}
    for r, it in info.items():
        if it['kind'] == 'empty' and r in rows:
            row = rows[r]
            if r in lasts:
                row.attrib.pop('hidden', None)
            else:
                row.set('hidden', '1')
            cI = cell(row, 'I')
            if cI.find(N + 'f') is None:
                set_formula(cI, ('IF($E{0}="Retard",IF(ISNUMBER($G{0}),MAX(0,dateCR-$G{0}),""),'
                                 'IF(AND(ISNUMBER($G{0}),ISNUMBER($E{0})),MAX(0,$E{0}-$G{0}),""))').format(r))
                cI.set('t', 'str')
                etree.SubElement(cI, N + 'v').text = ''
    # -- sauts de page : avant chaque TRAVAUX qui a au moins une ligne visible
    ws = b.ws(sh)
    rb = ws.find(N + 'rowBreaks')
    if rb is not None:
        rows, tabs, info = model(sh)
        brks = []
        for r, it in sorted(info.items()):
            if it['kind'] == 'section' and it['section'] == 'TRAVAUX':
                t = [t for t in tabs if t['name'] == it['table']][0]
                if any(info[k]['kind'] == 'obs' and not info[k]['hidden'] for k in range(r + 1, t['r1'] + 1)):
                    brks.append(r - 1)
        for brk in list(rb):
            rb.remove(brk)
        for k in brks:
            etree.SubElement(rb, N + 'brk', id=str(k), max='16383', man='1')
        rb.set('count', str(len(brks)))
        rb.set('manualBreakCount', str(len(brks)))
    b.touch(b.sheets[sh])

# ================================================================== 4. Référentiel : prochain N°
if 'Référentiel' in b.sheets:
    ws = b.ws('Référentiel')
    for row in ws.find(N + 'sheetData'):
        cA = row.find(N + 'c[@r="A%s"]' % row.get('r'))
        code = text_of(cA).strip() if cA is not None else ''
        if code in next_num and int(row.get('r')) > 4:
            set_si(cell(row, 'F'), si_new([(None, '%s-%03d' % (code, next_num[code] + 1))]))
    b.touch(b.sheets['Référentiel'])

b.xml('xl/workbook.xml').find(N + 'calcPr').set('fullCalcOnLoad', '1')
b.touch('xl/workbook.xml')
b.save(OUT)
json.dump(report, open(REPORT, 'w', encoding='utf-8'), ensure_ascii=False, indent=1)
print('OK', OUT, {k: len(v) for k, v in report.items()})
