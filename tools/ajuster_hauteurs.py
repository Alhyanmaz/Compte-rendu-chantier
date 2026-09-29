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
"""
import copy
import os
import re
import subprocess
import sys
import tempfile
import zipfile

import openpyxl

LO_LINE, LO_PAD = 10.45, 1.45        # LibreOffice : pt par ligne de texte (9 pt) et marge
XL_LINE, XL_PAD = 12.75, 4.5         # Excel : pt par ligne de texte (9 pt) et marge de sécurité
MIN_HT = 17.4                        # hauteur par défaut du classeur
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


def main(src, dst):
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
            zo.writestr(copy.copy(i), d)
    r = soffice(['--headless', '--convert-to', 'xlsx:Calc MS Excel 2007 XML', '--outdir', os.path.join(tmp, 'lo'), auto])
    lo = openpyxl.load_workbook(os.path.join(tmp, 'lo', 'auto.xlsx'))
    changed = 0
    with zipfile.ZipFile(dst, 'w', zipfile.ZIP_DEFLATED) as zo:
        for i in z.infolist():
            d = z.read(i.filename)
            if i.filename in targets:
                sh = targets[i.filename]
                ws_o, ws_l = wbo[sh], lo[sh]

                def fix(m):
                    nonlocal changed
                    t = m.group(0)
                    rr = int(m.group(1))
                    if 'hidden="1"' in t:
                        return t
                    a = ws_o.cell(rr, 1).value
                    if not (isinstance(a, str) and re.match(r'^[A-Z0-9.]+-\d{3}$', a)):
                        return t                      # seules les lignes d'observation sont ajustées
                    h = ws_l.row_dimensions[rr].height or MIN_HT
                    lines = max(1, int((h - LO_PAD) / LO_LINE + 0.05))
                    ht = max(MIN_HT, round(lines * XL_LINE + XL_PAD, 1))
                    t = re.sub(r'\sht="[^"]*"', '', t)
                    t = re.sub(r'\scustomHeight="[^"]*"', '', t)
                    changed += 1
                    return t[:-1] + ' ht="%s" customHeight="1">' % ht if not t.endswith('/>') else t[:-2] + ' ht="%s" customHeight="1"/>' % ht
                d = re.sub(r'<row r="(\d+)"[^>]*>', fix, d.decode('utf-8')).encode('utf-8')
            zo.writestr(copy.copy(i), d)
    print('lignes ajustées :', changed)


if __name__ == '__main__':
    main(sys.argv[1], sys.argv[2])
