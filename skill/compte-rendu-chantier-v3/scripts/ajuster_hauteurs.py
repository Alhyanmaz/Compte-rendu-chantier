# -*- coding: utf-8 -*-
"""Ajuste la hauteur des lignes d'un CR pour que tout le texte soit visible dans Excel.

Usage : python ajuster_hauteurs.py ENTREE.xlsx SORTIE.xlsx

Principe (29/09/2026) :
1. LibreOffice, avec les polices Denim INK installées (~/.fonts), calcule le nombre de lignes
   de texte de chaque ligne du tableau (hauteur optimale, métriques « hhea » : 10,45 pt par
   ligne de texte en 9 pt).
2. Ce nombre est converti en hauteur Excel : Windows utilise les métriques « win » de la police,
   soit ~12,75 pt par ligne de texte en 9 pt (vérifié sur une ligne calculée par Excel dans le
   CRC-15 : 2 lignes = 28,8 pt).
3. Seul l'attribut ht des lignes visibles des onglets d'observations est réécrit, dans le XML
   (aucun autre octet du classeur ne change).

Prérequis : LibreOffice Calc et les polices Denim INK installées dans l'environnement.
Sans LibreOffice (session Cowork ou projet claude.ai sans installation possible) : repli automatique sur une
estimation par les métriques de la police (fontTools, police Denim INK si elle est dans ~/.fonts, sinon une largeur
moyenne de caractère). Moins fiable : le signaler à José et lui demander un contrôle dans l'aperçu avant impression.
"""
import glob
import math
import shutil
import copy
import os
import re
import subprocess
import sys
import tempfile
import zipfile

import openpyxl

LO_LINE, LO_PAD = 10.45, 1.45        # LibreOffice : pt par ligne de texte (9 pt) et marge
XL_LINE, XL_PAD = 13.2, 9.0          # Excel : pt par ligne (12,75 mesurés + marge) et marge fixe (30/09/2026 : trop juste)
SHRINK = 0.86                        # largeur utilisée pour la mesure (marge de sécurité)
MIN_HT = 17.4                        # hauteur par défaut du classeur
EST_MARGE = 0.95                     # estimation sans LibreOffice : largeur réduite en plus (prudence)
OBS_TITLES = None                    # None = onglets dont la ligne d'en-tête contient « N° » en A


def soffice(args):
    scripts = os.environ.get('XLSX_SKILL_SCRIPTS')
    if scripts:
        sys.path.insert(0, scripts)
        from office.soffice import run_soffice
        return run_soffice(args, capture_output=True, text=True, timeout=300)
    return subprocess.run(['soffice'] + args, capture_output=True, text=True, timeout=300)


def sheet_parts(z):
    wb = z.read('xl/workbook.xml').decode('utf-8')
    rels = z.read('xl/_rels/workbook.xml.rels').decode('utf-8')
    rid = dict(re.findall(r'Id="([^"]+)"[^>]*Target="([^"]+)"', rels))
    rid.update({a: b for b, a in re.findall(r'Target="([^"]+)"[^>]*Id="([^"]+)"', rels)})
    out = {}
    for name, r in re.findall(r'<sheet name="([^"]+)"[^>]*r:id="([^"]+)"', wb):
        out[name.replace('&amp;', '&')] = 'xl/' + rid[r].lstrip('/').replace('xl/', '', 1)
    return out


def lo_disponible():
    if os.environ.get('CR_SANS_LIBREOFFICE'):          # forcer l'estimation (tests)
        return False
    return bool(os.environ.get('XLSX_SKILL_SCRIPTS') or shutil.which('soffice') or shutil.which('libreoffice'))


def estimateur():
    """Nombre de lignes de texte d'une cellule, par largeur de caractères (repli sans LibreOffice)."""
    fonts = sorted(glob.glob(os.path.expanduser('~/.fonts/*DenimINK-Medium.otf')))
    adv = None
    if fonts:
        try:
            from fontTools.ttLib import TTFont
            f = TTFont(fonts[0])
            cmap, hm, upm = f.getBestCmap(), f['hmtx'], f['head'].unitsPerEm
            adv = lambda ch: hm[cmap.get(ord(ch), cmap[32])][0] / upm   # noqa: E731
        except Exception:
            adv = None
    if adv is None:
        adv = lambda ch: 0.55                                            # noqa: E731  largeur moyenne (em)
    mdw = round(adv('0') * 11 * 96 / 72)                                 # largeur de chiffre, police par défaut 11 pt

    def lignes(text, width_chars, pt=9):
        px = math.trunc(((256 * width_chars * SHRINK * EST_MARGE + math.trunc(128 / mdw)) / 256) * mdw) - 7
        n = 0
        for para in str(text).split('\n'):
            n += 1
            cur = ''
            for w in para.split(' '):
                t = (cur + ' ' + w) if cur else w
                if sum(adv(c) for c in t) * pt * 96 / 72 <= px or not cur:
                    cur = t
                else:
                    n += 1
                    cur = w
        return n
    return lignes


def main(src, dst):
    if not lo_disponible():
        print('LibreOffice absent : hauteurs ESTIMÉES (à contrôler dans l\'aperçu avant impression)')
    wbo = openpyxl.load_workbook(src)
    obs = [ws.title for ws in wbo.worksheets
           if any(str(ws.cell(r, 1).value).strip() == 'N°' for r in range(1, 12))]
    z = zipfile.ZipFile(src)
    parts = sheet_parts(z)
    targets = {parts[s]: s for s in obs}
    tmp = tempfile.mkdtemp()
    auto = os.path.join(tmp, 'auto.xlsx')
    with zipfile.ZipFile(auto, 'w', zipfile.ZIP_DEFLATED) as zo:
        for i in z.infolist():
            d = z.read(i.filename)
            if i.filename in targets:
                d = re.sub(r'<row [^>]*>', lambda m: m.group(0) if 'hidden="1"' in m.group(0) else
                           re.sub(r'\scustomHeight="[^"]*"', '', re.sub(r'\sht="[^"]*"', '', m.group(0))),
                           d.decode('utf-8')).encode('utf-8')
                # seules les colonnes imprimées (A:E) comptent : HISTORIQUE (F) et ROUTAGE (J) sont retirées de la copie
                # de mesure (sinon un historique long gonfle la hauteur imprimée — défaut constaté le 30/09/2026)
                d = re.sub(rb'<c r="(?:[F-Z]|[A-Z]{2,})\d+"(?:[^>]*/>|[^>]*>.*?</c>)', b'', d, flags=re.S)
                # mesure prudente : colonnes B à E rétrécies de 14 % (le rendu Excel coupe un peu plus tôt que LibreOffice)
                d = re.sub(rb'<col min="([2-5])" max="([2-5])" width="([0-9.]+)"',
                           lambda m: b'<col min="%s" max="%s" width="%.3f"' % (m.group(1), m.group(2), float(m.group(3)) * SHRINK), d)
            zo.writestr(copy.copy(i), d)
    if lo_disponible():
        soffice(['--headless', '--convert-to', 'xlsx:Calc MS Excel 2007 XML', '--outdir', os.path.join(tmp, 'lo'), auto])
        lo = openpyxl.load_workbook(os.path.join(tmp, 'lo', 'auto.xlsx'))
        est = None
    else:
        lo, est = None, estimateur()
    changed = 0
    with zipfile.ZipFile(dst, 'w', zipfile.ZIP_DEFLATED) as zo:
        for i in z.infolist():
            d = z.read(i.filename)
            if i.filename in targets:
                sh = targets[i.filename]
                ws_o = wbo[sh]
                ws_l = lo[sh] if lo is not None else None
                wid = {k: ws_o.column_dimensions[k].width or 8.43 for k in 'BCDE'}

                def fix(m):
                    nonlocal changed
                    t = m.group(0)
                    rr = int(m.group(1))
                    if 'hidden="1"' in t:
                        return t
                    a = ws_o.cell(rr, 1).value
                    if not (isinstance(a, str) and re.match(r'^[A-Z0-9.]+-\d{3}$', a)):
                        return t                      # seules les lignes d'observation sont ajustées
                    if ws_l is not None:
                        h = ws_l.row_dimensions[rr].height or MIN_HT
                        lines = max(1, int((h - LO_PAD) / LO_LINE + 0.05))
                    else:
                        lines = max([1] + [est(ws_o['%s%d' % (k, rr)].value, wid[k]) for k in 'BCDE'
                                           if isinstance(ws_o['%s%d' % (k, rr)].value, str)])
                    ht = max(MIN_HT, round(lines * XL_LINE + XL_PAD, 1))
                    t = re.sub(r'\sht="[^"]*"', '', t)
                    t = re.sub(r'\scustomHeight="[^"]*"', '', t)
                    changed += 1
                    return t[:-1] + ' ht="%s" customHeight="1">' % ht if not t.endswith('/>') else t[:-2] + ' ht="%s" customHeight="1"/>' % ht
                d = re.sub(r'<row [^>]*?\br="(\d+)"[^>]*>', fix, d.decode('utf-8')).encode('utf-8')
            zo.writestr(copy.copy(i), d)
    print('lignes ajustées :', changed)


if __name__ == '__main__':
    main(sys.argv[1], sys.argv[2])
