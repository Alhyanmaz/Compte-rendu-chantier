# -*- coding: utf-8 -*-
"""cr_xml.py — édition chirurgicale d'un classeur CR (.xlsx) sans réécriture globale.
Seules les parties XML réellement modifiées sont réécrites ; toutes les autres
restent identiques octet pour octet (Coordonnees, Généralités, images, boutons...)."""
import copy, datetime, re, zipfile
from lxml import etree

NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
RNS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
PNS = 'http://schemas.openxmlformats.org/package/2006/relationships'
N = '{%s}' % NS
RED, BLACK = 'FFFF0000', 'FF000000'
GREY = dict(theme='0', tint='-0.499984740745262')
EPOCH = datetime.datetime(1899, 12, 30)


def col_letter(c):
    s = ''
    while c:
        c, r = divmod(c - 1, 26)
        s = chr(65 + r) + s
    return s


def col_index(ref):
    m = re.match(r'([A-Z]+)', ref)
    n = 0
    for ch in m.group(1):
        n = n * 26 + ord(ch) - 64
    return n


class Book:
    def __init__(self, path):
        self.src = path
        with zipfile.ZipFile(path) as z:
            self.infos = z.infolist()
            self.data = {i.filename: z.read(i.filename) for i in self.infos}
        self.trees, self.dirty = {}, set()
        wb = self.xml('xl/workbook.xml')
        rels = self.rels('xl/workbook.xml')
        self.sheets = {}
        for s in wb.iter(N + 'sheet'):
            tgt = rels[s.get('{%s}id' % RNS)]
            self.sheets[s.get('name')] = 'xl/' + tgt.lstrip('/').replace('xl/', '', 1)
        self.sst = self.xml('xl/sharedStrings.xml') if 'xl/sharedStrings.xml' in self.data else None
        self.si = list(self.sst.iter(N + 'si')) if self.sst is not None else []
        self.styles = self.xml('xl/styles.xml')
        self._xf_cache = {}

    # ---------- bas niveau ----------
    def xml(self, name):
        if name not in self.trees:
            self.trees[name] = etree.fromstring(self.data[name])
        return self.trees[name]

    def touch(self, name):
        self.dirty.add(name)

    def rels(self, part):
        d, f = part.rsplit('/', 1)
        rp = '%s/_rels/%s.rels' % (d, f)
        if rp not in self.data:
            return {}
        return {r.get('Id'): r.get('Target') for r in self.xml(rp)}

    def rel_target(self, part, rtype):
        d, f = part.rsplit('/', 1)
        rp = '%s/_rels/%s.rels' % (d, f)
        if rp not in self.data:
            return None
        for r in self.xml(rp):
            if r.get('Type').endswith('/' + rtype):
                t = r.get('Target')
                if t.startswith('/'):
                    return t.lstrip('/')
                parts = (d + '/' + t).split('/')
                out = []
                for p in parts:
                    if p == '..':
                        out.pop()
                    else:
                        out.append(p)
                return '/'.join(out)
        return None

    # ---------- cellules ----------
    def ws(self, sheet):
        return self.xml(self.sheets[sheet])

    def row(self, sheet, r, create=True):
        sd = self.ws(sheet).find(N + 'sheetData')
        for row in sd.iter(N + 'row'):
            if int(row.get('r')) == r:
                return row
            if int(row.get('r')) > r:
                if not create:
                    return None
                new = etree.Element(N + 'row', r=str(r))
                row.addprevious(new)
                self.touch(self.sheets[sheet])
                return new
        if not create:
            return None
        new = etree.SubElement(sd, N + 'row', r=str(r))
        self.touch(self.sheets[sheet])
        return new

    def cell(self, sheet, ref, create=True):
        r = int(re.sub(r'[A-Z]', '', ref))
        row = self.row(sheet, r, create)
        if row is None:
            return None
        ci = col_index(ref)
        for c in row.iter(N + 'c'):
            if c.get('r') == ref:
                return c
            if col_index(c.get('r')) > ci:
                if not create:
                    return None
                new = etree.Element(N + 'c', r=ref)
                c.addprevious(new)
                return new
        if not create:
            return None
        return etree.SubElement(row, N + 'c', r=ref)

    def text(self, sheet, ref):
        c = self.cell(sheet, ref, False)
        if c is None:
            return ''
        if c.get('t') == 's':
            return ''.join(self.si[int(c.find(N + 'v').text)].itertext())
        if c.get('t') == 'inlineStr':
            return ''.join(c.find(N + 'is').itertext())
        v = c.find(N + 'v')
        f = c.find(N + 'f')
        if f is not None:
            return '=' + (f.text or '')
        return v.text if v is not None else ''

    def value(self, sheet, ref):
        """Valeur typée : datetime si la cellule est une date (style numFmt date), sinon texte."""
        c = self.cell(sheet, ref, False)
        if c is None:
            return None
        t = self.text(sheet, ref)
        if c.get('t') in ('s', 'inlineStr', 'str') or t.startswith('='):
            return t
        try:
            n = float(t)
        except ValueError:
            return t
        if self.is_date_style(c.get('s')):
            return EPOCH + datetime.timedelta(days=n)
        return n

    def is_date_style(self, s):
        if s is None:
            return False
        xf = self.styles.find(N + 'cellXfs')[int(s)]
        nf = int(xf.get('numFmtId', 0))
        if 14 <= nf <= 22 or 45 <= nf <= 47:
            return True
        nfs = self.styles.find(N + 'numFmts')
        if nfs is not None:
            for f in nfs:
                if int(f.get('numFmtId')) == nf:
                    code = re.sub(r'"[^"]*"|\[[^\]]*\]', '', f.get('formatCode')).lower()
                    return 'd' in code and 'y' in code
        return False

    # ---------- styles ----------
    def _font_of(self, s):
        xf = self.styles.find(N + 'cellXfs')[int(s or 0)]
        return self.styles.find(N + 'fonts')[int(xf.get('fontId', 0))]

    def restyle(self, c, color=None, bold=None, fill_rgb=None):
        """Clone le style de la cellule en changeant couleur de police / gras / remplissage."""
        s = int(c.get('s', 0))
        key = (s, repr(color), bold, fill_rgb)
        if key not in self._xf_cache:
            fonts, fills = self.styles.find(N + 'fonts'), self.styles.find(N + 'fills')
            xfs = self.styles.find(N + 'cellXfs')
            xf = copy.deepcopy(xfs[s])
            if color is not None or bold is not None:
                f = copy.deepcopy(fonts[int(xf.get('fontId', 0))])
                if color is not None:
                    _set_color(f, color)
                if bold is not None:
                    for b in f.findall(N + 'b'):
                        f.remove(b)
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
            xfs.append(xf)
            xfs.set('count', str(len(xfs)))
            self.touch('xl/styles.xml')
            self._xf_cache[key] = str(len(xfs) - 1)
        c.set('s', self._xf_cache[key])

    # ---------- texte enrichi ----------
    def runs(self, sheet, ref):
        """Liste de runs [(rPr, texte)] ; rPr toujours explicite."""
        c = self.cell(sheet, ref)
        base = self._font_of(c.get('s'))
        src = None
        if c.get('t') == 's':
            src = self.si[int(c.find(N + 'v').text)]
        elif c.get('t') == 'inlineStr':
            src = c.find(N + 'is')
        out = []
        if src is None:
            v = c.find(N + 'v')
            if v is not None and v.text:
                out.append((_rpr_from_font(base), v.text))
            return out
        rs = src.findall(N + 'r')
        if not rs:
            t = src.find(N + 't')
            if t is not None and t.text:
                out.append((_rpr_from_font(base), t.text))
            return out
        for r in rs:
            rpr = r.find(N + 'rPr')
            rpr = copy.deepcopy(rpr) if rpr is not None else _rpr_from_font(base)
            out.append((rpr, r.find(N + 't').text or ''))
        return out

    def set_runs(self, sheet, ref, runs):
        """Écrit les runs dans une NOUVELLE entrée sharedStrings (jamais d'inlineStr)."""
        c = self.cell(sheet, ref)
        was_shared = c.get('t') == 's'
        si = etree.SubElement(self.sst, N + 'si')
        for rpr, txt in runs:
            r = etree.SubElement(si, N + 'r')
            r.append(copy.deepcopy(rpr))
            t = etree.SubElement(r, N + 't')
            t.text = txt
            if txt != txt.strip() or '\n' in txt:
                t.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
        self.si.append(si)
        cnt = int(self.sst.get('count', len(self.si))) + (0 if was_shared else 1)
        self.sst.set('count', str(cnt))
        self.sst.set('uniqueCount', str(len(self.si)))
        for ch in list(c):
            c.remove(ch)
        c.set('t', 's')
        etree.SubElement(c, N + 'v').text = str(len(self.si) - 1)
        self.touch('xl/sharedStrings.xml')
        self.touch(self.sheets[sheet])

    def recolor_text(self, sheet, ref, color, keep_last=False):
        """Recolore TOUS les runs d'une cellule texte (et la police de cellule)."""
        c = self.cell(sheet, ref, False)
        if c is None:
            return
        rs = self.runs(sheet, ref)
        if rs and c.get('t') in ('s', 'inlineStr'):
            last = len(rs) - 1
            for i, (rpr, _) in enumerate(rs):
                if keep_last and i == last:
                    continue
                _set_color(rpr, color)
            self.set_runs(sheet, ref, rs)
        self.restyle(c, color=color)
        self.touch(self.sheets[sheet])

    def append_red(self, sheet, ref, text, history=BLACK):
        """Ajout du jour : historique -> `history` (noir par défaut), ajout en rouge non gras."""
        rs = self.runs(sheet, ref)
        for rpr, _ in rs:
            _set_color(rpr, history)
        base = copy.deepcopy(rs[-1][0]) if rs else _rpr_from_font(self._font_of(self.cell(sheet, ref).get('s')))
        for b in base.findall(N + 'b'):
            base.remove(b)
        _set_color(base, RED)
        rs.append((base, text))
        self.set_runs(sheet, ref, rs)

    def set_text(self, sheet, ref, text, color=RED, bold=False, style_from=None):
        c = self.cell(sheet, ref)
        if style_from:
            c.set('s', self.cell(sheet, style_from).get('s', '0'))
        rpr = _rpr_from_font(self._font_of(c.get('s')))
        for b in rpr.findall(N + 'b'):
            rpr.remove(b)
        if bold:
            rpr.insert(0, etree.Element(N + 'b'))
        _set_color(rpr, color)
        self.set_runs(sheet, ref, [(rpr, text)])
        self.restyle(c, color=color, bold=bold)

    def set_date(self, sheet, ref, date, color=RED, style_from=None):
        c = self.cell(sheet, ref)
        if style_from:
            c.set('s', self.cell(sheet, style_from).get('s', '0'))
        for ch in list(c):
            c.remove(ch)
        c.attrib.pop('t', None)
        etree.SubElement(c, N + 'v').text = str((date - EPOCH).days)
        self.restyle(c, color=color)
        self.touch(self.sheets[sheet])

    def set_formula(self, sheet, ref, formula, color=RED, style_from=None):
        c = self.cell(sheet, ref)
        if style_from:
            c.set('s', self.cell(sheet, style_from).get('s', '0'))
        for ch in list(c):
            c.remove(ch)
        c.attrib.pop('t', None)
        etree.SubElement(c, N + 'f').text = formula.lstrip('=')
        self.restyle(c, color=color)
        self.formulas_added = True
        self.touch(self.sheets[sheet])

    # ---------- lignes ----------
    def hide_row(self, sheet, r):
        self.row(sheet, r).set('hidden', '1')
        self.touch(self.sheets[sheet])

    def is_hidden(self, sheet, r):
        row = self.row(sheet, r, False)
        return row is not None and row.get('hidden') in ('1', 'true')

    def set_height(self, sheet, r, pt):
        row = self.row(sheet, r)
        row.set('ht', '%.2f' % pt)
        row.set('customHeight', '1')
        self.touch(self.sheets[sheet])

    def fill_row(self, sheet, r, rgb, cols='ABCD'):
        for col in cols:
            self.restyle(self.cell(sheet, '%s%d' % (col, r)), fill_rgb=rgb)
        self.touch(self.sheets[sheet])

    def col_width(self, sheet, col):
        cols = self.ws(sheet).find(N + 'cols')
        ci = col_index(col)
        if cols is not None:
            for c in cols:
                if int(c.get('min')) <= ci <= int(c.get('max')):
                    return float(c.get('width'))
        fmt = self.ws(sheet).find(N + 'sheetFormatPr')
        return float(fmt.get('defaultColWidth', 8.43)) if fmt is not None else 8.43

    def fit_height(self, sheet, r, cols='ABCD', font_pt=9.0, min_pt=15.0):
        """Hauteur estimée pour que tout le texte (renvoi à la ligne) soit visible."""
        lines = 1
        for col in cols:
            txt = self.text(sheet, '%s%d' % (col, r))
            if not txt or txt.startswith('='):
                continue
            cpl = max(1, int(self.col_width(sheet, col) * 1.15))   # police 9 pt ~ 1,15 car/unité
            n = sum(max(1, -(-len(p) // cpl)) for p in txt.split('\n'))
            lines = max(lines, n)
        self.set_height(sheet, r, max(min_pt, lines * font_pt * 1.35 + 4))

    # ---------- notes de cellule ----------
    def notes(self, sheet):
        part = self.sheets[sheet]
        cp = self.rel_target(part, 'comments')
        if not cp:
            return {}
        return {c.get('ref'): ''.join(c.find(N + 'text').itertext())
                for c in self.xml(cp).iter(N + 'comment')}

    def delete_note(self, sheet, ref):
        part = self.sheets[sheet]
        cp = self.rel_target(part, 'comments')
        if cp:
            root = self.xml(cp)
            for c in list(root.iter(N + 'comment')):
                if c.get('ref') == ref:
                    c.getparent().remove(c)
                    self.touch(cp)
        vp = self.notes_vml(part)
        if vp:
            txt = self.data[vp].decode('utf-8', 'replace') if vp not in self.trees else None
            r0, c0 = int(re.sub(r'[A-Z]', '', ref)) - 1, col_index(ref) - 1
            pat = re.compile(r'<v:shape\b(?:(?!</v:shape>).)*?<x:Row>\s*%d\s*</x:Row>\s*<x:Column>\s*%d\s*</x:Column>.*?</v:shape>' % (r0, c0), re.S)
            new, k = pat.subn('', txt)
            if k:
                self.data[vp] = new.encode('utf-8')
        # commentaire threadé éventuel sur la même cellule
        tp = self.rel_target(part, 'threadedComment')
        if tp:
            root = self.xml(tp)
            for c in list(root):
                if c.get('ref') == ref:
                    root.remove(c)
                    self.touch(tp)

    def notes_vml(self, part):
        """VML des bulles de notes = cible de <legacyDrawing> (PAS legacyDrawingHF = logo d'en-tête)."""
        ld = self.xml(part).find(N + 'legacyDrawing')
        if ld is None:
            return None
        rid = ld.get('{%s}id' % RNS)
        d, f = part.rsplit('/', 1)
        for r in self.xml('%s/_rels/%s.rels' % (d, f)):
            if r.get('Id') == rid:
                t = r.get('Target')
                if t.startswith('/'):
                    return t.lstrip('/')
                out = []
                for p in (d + '/' + t).split('/'):
                    out.pop() if p == '..' else out.append(p)
                return '/'.join(out)
        return None

    # ---------- enregistrement ----------
    def save(self, dst):
        if getattr(self, 'formulas_added', False):
            self._drop_calcchain()
        for name in self.dirty:
            self.data[name] = etree.tostring(self.trees[name], xml_declaration=True,
                                             encoding='UTF-8', standalone=True)
        with zipfile.ZipFile(dst, 'w', zipfile.ZIP_DEFLATED) as z:
            for i in self.infos:
                if i.filename in self.data:
                    z.writestr(i, self.data[i.filename])

    def _drop_calcchain(self):
        """Formules ajoutées sans valeur en cache : on retire calcChain et on force le recalcul."""
        if 'xl/calcChain.xml' in self.data:
            del self.data['xl/calcChain.xml']
            self.infos = [i for i in self.infos if i.filename != 'xl/calcChain.xml']
            for part in ('xl/_rels/workbook.xml.rels', '[Content_Types].xml'):
                root = self.xml(part)
                for el in list(root):
                    if 'calcChain' in (el.get('Target', '') + el.get('PartName', '')):
                        root.remove(el)
                self.touch(part)
        wb = self.xml('xl/workbook.xml')
        cp = wb.find(N + 'calcPr')
        if cp is None:
            cp = etree.SubElement(wb, N + 'calcPr')
        cp.set('fullCalcOnLoad', '1')
        self.touch('xl/workbook.xml')


# ---------- utilitaires rPr ----------
def _set_color(el, color):
    for c in el.findall(N + 'color'):
        el.remove(c)
    ce = etree.Element(N + 'color')
    if isinstance(color, dict):
        for k, v in color.items():
            ce.set(k, v)
    else:
        ce.set('rgb', color)
    # position : avant sz si présent (ordre usuel Excel), sinon en fin
    sz = el.find(N + 'sz')
    if sz is not None:
        sz.addprevious(ce)
    else:
        el.append(ce)


def _rpr_from_font(font):
    rpr = etree.Element(N + 'rPr')
    for ch in font:
        tag = etree.QName(ch).localname
        if tag == 'name':
            e = etree.SubElement(rpr, N + 'rFont')
            e.set('val', ch.get('val'))
        else:
            rpr.append(copy.deepcopy(ch))
    return rpr


def color_of(rpr):
    c = rpr.find(N + 'color')
    if c is None:
        return 'auto'
    if c.get('rgb'):
        return c.get('rgb')
    if c.get('theme') == '0' and float(c.get('tint', 0)) < -0.2:
        return 'GREY'
    return 'theme%s' % c.get('theme')