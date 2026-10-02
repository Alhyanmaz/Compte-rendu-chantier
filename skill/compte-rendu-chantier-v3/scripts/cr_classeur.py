# -*- coding: utf-8 -*-
"""Outils communs aux scripts qui modifient un CR au nouveau format (intégration hebdomadaire, relecture).

Usage : import cr_classeur as K ; K.init(Book(chemin), date_du_cr) ; puis les fonctions du module.
Règles : docs/decisions-HONGUEMARE.md."""
import copy
import datetime
import re

from lxml import etree

from cr_xml import Book, N, RED, BLACK, GREY, EPOCH, _set_color, _rpr_from_font, color_of, col_index  # noqa: F401
from cr_texte import condense, plain_rpr  # noqa: F401

b = None
DATE = None
DSTR = None
_xf = {}
TERMINAUX = {'ANNULÉ', 'DOUBLON', 'SANS OBJET', 'REFUSÉ'}
CODE_RE = re.compile(r'^([A-Z0-9.]+)-(\d{3})$')
SECTION_CODES = [("maitre d'ouvrage_ville", 'MOA'), ('siege 27', 'SIE'), ('amo_', 'AMO'), ('bureau de contr', 'CT'),
                 ('csps_', 'SPS'), ("maitrise d'oeuvre", 'MOE'), ('economiste', 'ECO'), ('bet cvc', 'BET.CVC'),
                 ('bet électricité', 'BET.ELE'), ('bet thermique', 'BET.THE'), ('bet vrd', 'BET.VRD'),
                 ('structure béton', 'BET.SBE'), ('structure bois', 'BET.SBO'), ('accoustique', 'BET.ACO'),
                 ('amiante', 'BET.AMI'), ('concessionaires_réseaux', 'CON.RES'), ('concessionaires_eaux', 'CON.EAU'),
                 ('concessionaires_electricite', 'CON.ELE'), ('concessionaires_assainissement', 'CON.ASS')]


def init(book, date):
    global b, DATE, DSTR
    b, DATE, DSTR = book, date, date.strftime('%d/%m/%Y')
    b.formulas_added = True
    _xf.clear()
    return b


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
    if c.get('t') == 'inlineStr':
        i = c.find(N + 'is')
        return ''.join(i.itertext()) if i is not None else ''
    if c.get('t') == 'str':
        v = c.find(N + 'v')
        return v.text or '' if v is not None else ''
    v = c.find(N + 'v')
    return v.text if v is not None and v.text else ''


def runs_of(c):
    base = b._font_of(c.get('s'))
    if c.get('t') not in ('s', 'inlineStr'):
        t = text_of(c)
        return [(_rpr_from_font(base), t)] if t else []
    si = c.find(N + 'is') if c.get('t') == 'inlineStr' else b.si[int(c.find(N + 'v').text)]
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
    _set_color(rpr, RED)                       # écrit ce jour : rouge, PM compris (gris au CR suivant)
    return [(rpr, value)]


def write_fait_le(c, value):
    """Statut en texte enrichi rouge ; date en rouge (clôture de la semaine)."""
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


NEUTRES = {RED, BLACK, 'GREY', 'auto', 'theme1'}   # couleurs gérées par le script ; les autres (choix de José) sont gardées


def recolor_runs(runs, color, keep_red=False):
    for rp, _ in runs:
        c = color_of(rp)
        if c in NEUTRES and not (keep_red and c == RED):
            _set_color(rp, color)
    return runs


def color_cell(c, color):
    """Couleur de police d'une cellule (date, N°, statut) : style de cellule, et runs s'il y en a."""
    if c is None:
        return
    if color_of(_rpr_from_font(b._font_of(c.get('s')))) in NEUTRES:
        restyle(c, color=color)
    if c.get('t') == 's':
        rs = runs_of(c)
        if any(color_of(rp) in NEUTRES and color_of(rp) != color_of(_color_probe(color)) for rp, _ in rs):
            set_si(c, si_new(recolor_runs(rs, color)))


def _color_probe(color):
    e = etree.Element(N + 'rPr')
    _set_color(e, color)
    return e


def is_grey_status(c):
    e = text_of(c).strip().upper()
    return number_of(c) is not None or e == 'PM' or e in TERMINAUX


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
    if target.startswith('/'):                       # cible absolue (fichiers réécrits par d'autres outils)
        return target.lstrip('/')
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




# ------------------------------------------------------------------ mise en page et formules communes
def jours_formula(c, r):
    """Formule JOURS RETARD (colonne I, masquée) : jours calendaires de retard."""
    set_formula(c, ('IF($E{0}="Retard",IF(ISNUMBER($G{0}),MAX(0,dateCR-$G{0}),""),'
                    'IF(AND(ISNUMBER($G{0}),ISNUMBER($E{0})),MAX(0,$E{0}-$G{0}),""))').format(r))
    c.set('t', 'str')
    etree.SubElement(c, N + 'v').text = ''


def finir_onglet(sh):
    """Lignes vides masquées (réserve) sauf la dernière de chaque tableau ; sauts de page avant chaque TRAVAUX
    qui a au moins une ligne visible."""
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
                jours_formula(cI, r)
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


def numeros(sheets):
    """Dernier N° utilisé par code."""
    out = {}
    for sh in sheets:
        rows, tabs, info = model(sh)
        for r, it in info.items():
            if it['kind'] == 'obs':
                m = CODE_RE.match(it['num'])
                out[m.group(1)] = max(out.get(m.group(1), 0), int(m.group(2)))
    return out


def maj_referentiel(next_num):
    if 'Référentiel' not in b.sheets:
        return
    ws = b.ws('Référentiel')
    for row in ws.find(N + 'sheetData'):
        cA = row.find(N + 'c[@r="A%s"]' % row.get('r'))
        code = text_of(cA).strip() if cA is not None else ''
        if code in next_num and int(row.get('r')) > 4:
            set_si(cell(row, 'F'), si_new([(None, '%s-%03d' % (code, next_num[code] + 1))]))
    b.touch(b.sheets['Référentiel'])


def terminer(out):
    b.xml('xl/workbook.xml').find(N + 'calcPr').set('fullCalcOnLoad', '1')
    b.touch('xl/workbook.xml')
    b.save(out)


# ------------------------------------------------------------------ statut Retard (décision du 29/09/2026)
RETARD_LIGNE = '\n→ En retard depuis le '


def appliquer_retard(row, r):
    """FAIT LE = Retard : DATE RETARD (G, masquée) = date du CR si vide ; ligne imprimée
    « → En retard depuis le JJ/MM/AAAA : N jours » (jours calendaires ; sans compteur le premier jour).
    Statut levé : la ligne imprimée est retirée (G est gardée pour le compteur JOURS RETARD)."""
    cB, cE, cG = cell(row, 'B'), cell(row, 'E'), cell(row, 'G')
    runs = [x for x in runs_of(cB) if not x[1].startswith(RETARD_LIGNE)]
    changed = len(runs) != len(runs_of(cB))
    if text_of(cE).strip().upper() == 'RETARD':
        g = number_of(cG)
        if g is None:
            g = serial(DATE)
            set_num(cG, g)
            restyle(cG, color=BLACK)
        n = serial(DATE) - int(g)
        txt = RETARD_LIGNE + from_serial(g).strftime('%d/%m/%Y') + (' : %d jour%s' % (n, 's' if n > 1 else '') if n > 0 else '')
        rpr = plain_rpr(runs[-1][0] if runs else _rpr_from_font(b._font_of(cB.get('s'))))
        _set_color(rpr, RED)
        runs.append((rpr, txt))
        changed = True
    if changed:
        set_si(cB, si_new(runs))
    return changed


STATUTS_CANON = ['Relance', 'URGENT', 'Retard', 'PM', 'En cours', 'En attente', 'Annulé', 'Doublon', 'Sans objet',
                 'Refusé']


def statut_canonique(v):
    """Casse normalisée d'un statut saisi à la main (« pm » -> « PM ») ; « en attente enedis » -> « En attente ENEDIS »."""
    t = ' '.join(v.split())
    for s in STATUTS_CANON:
        if t.upper() == s.upper():
            return s
    if t.upper().startswith('EN ATTENTE '):
        return 'En attente ' + t[11:].upper()
    return t


# ------------------------------------------------------------------ relecture : fond rose, colonne ROUTAGE, onglet « Non routé »
# (décisions du 30/09/2026 : José prend des notes de cellule en réunion ; Claude les reformule à l'appui de l'audio)
ROSE = 'FFFADADD'          # point venu de la transcription seule (sans note de José)
ROSE_DOUTE = 'FFF4A6C0'    # [?] : hypothèse de Claude, à trancher par José
ROSES = {ROSE, ROSE_DOUTE}
BLEU_DOUBLON = 'FFBDD7EE'  # repère de doublon posé par la migration (absent du classeur d'origine), retiré au remoulinage
ROUT_COL = 'J'
NON_ROUTE = 'Non routé'
_fx = {}


def fill_of(c):
    xf = b.styles.find(N + 'cellXfs')[int(c.get('s', '0'))]
    f = b.styles.find(N + 'fills')[int(xf.get('fillId', 0))]
    fg = f.find('.//' + N + 'fgColor')
    return fg.get('rgb') if fg is not None else None


def set_fill(c, rgb):
    """Fond uni rgb (None = aucun fond), sur un clone du style de la cellule."""
    key = (c.get('s', '0'), rgb)
    if key not in _fx:
        st = b.styles
        fills, xfs = st.find(N + 'fills'), st.find(N + 'cellXfs')
        xf = copy.deepcopy(xfs[int(c.get('s', '0'))])
        if rgb is None:
            xf.set('fillId', '0')
        else:
            f = etree.SubElement(fills, N + 'fill')
            pf = etree.SubElement(f, N + 'patternFill', patternType='solid')
            etree.SubElement(pf, N + 'fgColor', rgb=rgb)
            etree.SubElement(pf, N + 'bgColor', indexed='64')
            fills.set('count', str(len(fills)))
            xf.set('fillId', str(len(fills) - 1))
        xf.set('applyFill', '1')
        xfs.append(xf)
        xfs.set('count', str(len(xfs)))
        b.touch('xl/styles.xml')
        _fx[key] = str(len(xfs) - 1)
    c.set('s', _fx[key])


def colonne_routage(sh, width=55):
    """Colonne J « ROUTAGE » : visible à l'écran, hors zone d'impression (A:E), texte renvoyé à la ligne."""
    ws = b.ws(sh)
    cols = ws.find(N + 'cols')
    if cols is None:
        cols = etree.Element(N + 'cols')
        ws.find(N + 'sheetData').addprevious(cols)
    k = col_index(ROUT_COL + '1')
    if not any(int(c.get('min')) <= k <= int(c.get('max')) for c in cols):
        new = etree.Element(N + 'col', min=str(k), max=str(k), width=str(width), customWidth='1')
        after = [c for c in cols if int(c.get('min')) < k]
        (after[-1].addnext(new) if after else cols.insert(0, new))
    rows, tabs, info = model(sh)
    for t in tabs:                                   # titre sur chaque ligne d'en-tête de tableau
        row = rows.get(t['r0'])
        if row is not None:
            hF, hJ = cell(row, 'F'), cell(row, ROUT_COL)
            if hF.get('s'):
                hJ.set('s', hF.get('s'))
            set_si(hJ, si_new([(None, 'ROUTAGE (relecture, non imprimé)')]))
    b.touch(b.sheets[sh])


def ecrire_routage(row, texte, style_from='F'):
    c = cell(row, ROUT_COL)
    src = cell(row, style_from)
    if src.get('s'):
        c.set('s', src.get('s'))
    rpr = plain_rpr(_rpr_from_font(b._font_of(c.get('s'))))
    _set_color(rpr, BLACK)
    set_si(c, si_new([(rpr, texte)]))


# ---- création d'onglet (reprise de migration_cr.py)
CTNS = 'http://schemas.openxmlformats.org/package/2006/content-types'
PNS = 'http://schemas.openxmlformats.org/package/2006/relationships'
RNS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'


def add_part(name, data_bytes, content_type):
    import zipfile
    b.data[name] = data_bytes
    zi = zipfile.ZipInfo(name, date_time=(2026, 9, 30, 12, 0, 0))
    zi.compress_type = zipfile.ZIP_DEFLATED
    b.infos.append(zi)
    ct = b.xml('[Content_Types].xml')
    etree.SubElement(ct, '{%s}Override' % CTNS, PartName='/' + name, ContentType=content_type)
    b.touch('[Content_Types].xml')


def add_rel(part_rels, rtype, target):
    root = b.xml(part_rels)
    ids = {r.get('Id') for r in root}
    k = 1
    while 'rId%d' % k in ids:
        k += 1
    etree.SubElement(root, '{%s}Relationship' % PNS, Id='rId%d' % k, Type=rtype, Target=target)
    b.touch(part_rels)
    return 'rId%d' % k


def onglet_non_route(items, visible=True):
    """Onglet « Non routé » (jamais imprimé) : passages de la transcription écartés par Claude, avec la raison.
    Créé au besoin (en dernier : les index des zones d'impression ne bougent pas), contenu remplacé à chaque CR."""
    wbx = b.xml('xl/workbook.xml')
    sheets = wbx.find(N + 'sheets')
    ref = b.ws('Référentiel') if 'Référentiel' in b.sheets else None

    def style(r, col):
        if ref is None:
            return None
        for row in ref.find(N + 'sheetData'):
            if row.get('r') == str(r):
                c = row.find(N + 'c[@r="%s%d"]' % (col, r))
                return c.get('s') if c is not None else None
    st_title, st_hdr, st_wrap = style(1, 'A'), style(4, 'A'), style(5, 'D')
    data = [('A', 1, 'NON ROUTÉ — passages de la transcription non repris au CR du %s (onglet non imprimé)' % DSTR, st_title),
            ('A', 3, 'Passage de la transcription', st_hdr), ('B', 3, 'Raison', st_hdr)]
    for i, it in enumerate(items or [{'extrait': '—', 'raison': 'Aucun passage écarté.'}]):
        data.append(('A', 4 + i, it.get('extrait', ''), st_wrap))
        data.append(('B', 4 + i, it.get('raison', ''), st_wrap))
    if NON_ROUTE not in b.sheets:
        nums = [int(re.sub(r'\D', '', n)) for n in b.data if re.match(r'xl/worksheets/sheet\d+\.xml$', n)]
        part = 'xl/worksheets/sheet%d.xml' % (max(nums) + 1)
        xml = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<worksheet xmlns="%s" xmlns:r="%s">'
               '<sheetViews><sheetView workbookViewId="0"/></sheetViews><sheetFormatPr defaultRowHeight="15"/>'
               '<cols><col min="1" max="1" width="62" customWidth="1"/><col min="2" max="2" width="38" customWidth="1"/></cols>'
               '<sheetData/><pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>'
               '</worksheet>' % (N[1:-1], RNS))
        add_part(part, xml.encode('utf-8'),
                 'application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml')
        rid = add_rel('xl/_rels/workbook.xml.rels', RNS + '/worksheet', 'worksheets/' + part.rsplit('/', 1)[1])
        sid = max(int(s.get('sheetId')) for s in sheets) + 1
        e = etree.SubElement(sheets, N + 'sheet', name=NON_ROUTE, sheetId=str(sid))
        e.set('{%s}id' % RNS, rid)
        b.sheets[NON_ROUTE] = part
    ws = b.ws(NON_ROUTE)
    sd = ws.find(N + 'sheetData')
    for row in list(sd):
        sd.remove(row)
    rows = {}
    for col, r, v, s in data:
        row = rows.get(r)
        if row is None:
            row = rows[r] = etree.SubElement(sd, N + 'row', r=str(r))
        c = etree.SubElement(row, N + 'c', r='%s%d' % (col, r))
        if s:
            c.set('s', s)
        set_si(c, si_new([(None, v)]))
    etat_non_route(visible)
    b.touch(b.sheets[NON_ROUTE])
    b.touch('xl/workbook.xml')


def etat_non_route(visible):
    for s in b.xml('xl/workbook.xml').find(N + 'sheets'):
        if s.get('name') == NON_ROUTE:
            if visible:
                s.attrib.pop('state', None)
            else:
                s.set('state', 'hidden')
            b.touch('xl/workbook.xml')


def vider_non_route():
    """Édition finale : onglet « Non routé » vidé et masqué (il est rempli à nouveau à l'intégration suivante)."""
    if NON_ROUTE not in b.sheets:
        return
    sd = b.ws(NON_ROUTE).find(N + 'sheetData')
    for row in list(sd):
        sd.remove(row)
    b.touch(b.sheets[NON_ROUTE])
    etat_non_route(False)
