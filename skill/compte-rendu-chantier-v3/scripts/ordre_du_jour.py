# -*- coding: utf-8 -*-
"""Ordre du jour de la prochaine réunion de chantier, tiré du CR remouliné (édition finale).

Usage : python ordre_du_jour.py CR_FINAL.xlsx ODJ.docx [--date AAAA-MM-JJ] [--heure 9H00] [--lieu "…"] [--json ODJ.json]

Date et heure par défaut : Page de garde (date du CR + 7 jours, heure sous la date). Lieu : ligne « LIEU » s'il est rempli.
Points retenus (observations visibles, non soldées, hors PM) :
  1. prioritaires : FAIT LE = URGENT ou Retard ;
  2. à traiter, par intervenant dans l'ordre du CR : POUR LE au plus tard le jour de la réunion (ou Relance sans
     échéance) ;
  3. en attente d'un tiers : FAIT LE commençant par « En attente » (une ligne par organisme attendu, liste des N°).
Présence requise : intervenants ayant au moins un point en 1 ou 2 ; « convoqués » : colonne C de Coordonnees.
Le Word est écrit directement (XML), sans bibliothèque : il s'ouvre dans Word et LibreOffice.
Créé le 02/10/2026 à la demande de José (ordre du jour livré avec le remoulinage).
"""
import argparse
import datetime
import json
import os
import re
import sys
import zipfile
from xml.sax.saxutils import escape

import openpyxl

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from cr_xml import Book  # noqa: E402
import cr_classeur as K  # noqa: E402

JOURS = ['lundi', 'mardi', 'mercredi', 'jeudi', 'vendredi', 'samedi', 'dimanche']
POLICE = 'Denim INK'          # police du CR (installée chez José ; remplacée ailleurs si absente)


# ------------------------------------------------------------------ lecture du CR
def sujet_et_dernier(b_text):
    """Sujet initial (avant le premier « → » ou « [...] ») et dernier état (dernière ligne « → Au … »)."""
    lignes = b_text.split('\n')
    sujet, i = [], 0
    while i < len(lignes) and not lignes[i].startswith(('→', '[...]')):
        sujet.append(lignes[i].strip())
        i += 1
    remarques = [l[1:].strip() for l in lignes if l.startswith('→ Au ')]
    relance = next((l[1:].strip() for l in reversed(lignes) if l.startswith('→ Relancé')), None)
    return ' '.join(x for x in sujet if x), (remarques[-1] if remarques else None), relance


def court(t, n=220):
    t = ' '.join(t.split())
    return t if len(t) <= n else t[:n - 1].rsplit(' ', 1)[0] + ' …'


def lire(src, date_reunion):
    K.init(Book(src), date_reunion)
    wv = openpyxl.load_workbook(src, data_only=True)
    ref = {}
    if 'Référentiel' in wv.sheetnames:
        ws = wv['Référentiel']
        for r in range(5, ws.max_row + 1):
            code, org = ws.cell(r, 1).value, ws.cell(r, 3).value
            if code and str(code).strip() and r < 60:
                ref[str(code).strip()] = (str(org).strip() if org and str(org).strip() != '—' else '')
    groupes, points = [], []
    for sh in K.obs_sheets():
        if sh == 'Modele_lot':
            continue
        rows, tabs, info = K.model(sh)
        titre_lot = None
        for r in range(1, 6):
            for c in rows.get(r, []):
                t = K.text_of(c).strip()
                if t.upper().startswith('LOT '):
                    titre_lot = ' '.join(t.split())
        for t in tabs:
            if titre_lot:
                label = '%s — %s' % (titre_lot, ref.get(t['code'], '')) if ref.get(t['code']) else titre_lot
            else:
                tr = next((k for k in range(t['r0'] - 1, max(0, t['r0'] - 4), -1)
                           if k in rows and k not in info and any(K.text_of(c).strip() for c in rows[k])), None)
                label = ' '.join(' '.join(K.text_of(c) for c in rows[tr]).split()) if tr else (t['code'] or sh)
            g = {'onglet': sh, 'code': t['code'], 'label': label, 'points': []}
            groupes.append(g)
            for r in range(t['r0'] + 1, t['r1'] + 1):
                it = info.get(r, {})
                if it.get('kind') != 'obs' or it.get('hidden'):
                    continue
                row = rows[r]
                cC, cD, cE = (K.cell(row, x) for x in 'CDE')
                e_num, e_txt = K.number_of(cE), K.text_of(cE).strip()
                up = e_txt.upper()
                if e_num is not None or up in K.TERMINAUX or up == 'PM':
                    continue                           # soldé, statut terminal ou pour mémoire
                d = K.number_of(cD)
                sujet, dernier, relance = sujet_et_dernier(K.text_of(K.cell(row, 'B')))
                p = {'num': it['num'], 'section': it.get('section') or '', 'groupe': label,
                     'organisme': ref.get(t['code'], ''), 'sujet': sujet, 'dernier': dernier, 'relance': relance,
                     'pour_le': K.from_serial(d) if d else None, 'statut': e_txt,
                     'aborde_le': K.from_serial(K.number_of(cC)) if K.number_of(cC) else None}
                if up in ('URGENT', 'RETARD'):
                    p['rubrique'] = 'prioritaire'
                elif up.startswith('EN ATTENTE'):
                    p['rubrique'] = 'attente'
                elif (p['pour_le'] and p['pour_le'] <= date_reunion) or (up == 'RELANCE' and not p['pour_le']):
                    p['rubrique'] = 'echeance'
                else:
                    continue                           # échéance après la réunion : pas à l'ordre du jour
                g['points'].append(p)
                points.append(p)
    conv = []
    if 'Coordonnees' in wv.sheetnames:
        ws = wv['Coordonnees']
        for r in range(1, ws.max_row + 1):
            org, c = ws.cell(r, 2).value, ws.cell(r, 5).value
            if org and str(c or '').strip().upper() == 'X' and str(org).strip().upper() not in ('ORGANISME', 'ENTREPRISES'):
                conv.append(' '.join(str(org).split()))
    return groupes, points, conv


# ------------------------------------------------------------------ Word (XML écrit à la main)
def run(t, b=False, i=False, sz=None, color=None):
    pr = ''
    if b:
        pr += '<w:b/>'
    if i:
        pr += '<w:i/>'
    if color:
        pr += '<w:color w:val="%s"/>' % color
    if sz:
        pr += '<w:sz w:val="%d"/><w:szCs w:val="%d"/>' % (sz, sz)
    return '<w:r>%s<w:t xml:space="preserve">%s</w:t></w:r>' % ('<w:rPr>%s</w:rPr>' % pr if pr else '', escape(t))


def para(runs, style=None, align=None, after=None, keep=False):
    pr = ''
    if style:
        pr += '<w:pStyle w:val="%s"/>' % style
    if keep:
        pr += '<w:keepNext/>'
    if after is not None:
        pr += '<w:spacing w:after="%d"/>' % after
    if align:
        pr += '<w:jc w:val="%s"/>' % align
    return '<w:p>%s%s</w:p>' % ('<w:pPr>%s</w:pPr>' % pr if pr else '', ''.join(runs))


def table(widths, header, lignes):
    grid = ''.join('<w:gridCol w:w="%d"/>' % w for w in widths)
    bd = ''.join('<w:%s w:val="single" w:sz="4" w:space="0" w:color="808080"/>' % s
                 for s in ('top', 'left', 'bottom', 'right', 'insideH', 'insideV'))

    def tr(cells, head=False):
        out = '<w:tr>%s' % ('<w:trPr><w:tblHeader/><w:cantSplit/></w:trPr>' if head else '<w:trPr><w:cantSplit/></w:trPr>')
        for w, c in zip(widths, cells):
            shade = '<w:shd w:val="clear" w:color="auto" w:fill="D9D9D9"/>' if head else ''
            paras = c if isinstance(c, list) else [para([run(str(c), b=head, sz=17)], after=0)]
            out += '<w:tc><w:tcPr><w:tcW w:w="%d" w:type="dxa"/>%s</w:tcPr>%s</w:tc>' % (w, shade, ''.join(paras))
        return out + '</w:tr>'
    return ('<w:tbl><w:tblPr><w:tblW w:w="%d" w:type="dxa"/><w:tblLayout w:type="fixed"/><w:tblBorders>%s</w:tblBorders>'
            '<w:tblCellMar><w:left w:w="70" w:type="dxa"/><w:right w:w="70" w:type="dxa"/></w:tblCellMar></w:tblPr>'
            '<w:tblGrid>%s</w:tblGrid>%s%s</w:tbl>' % (sum(widths), bd, grid, tr(header, True), ''.join(tr(l) for l in lignes)))


def fdate(d):
    return d.strftime('%d/%m/%Y') if d else '—'


def cellule_objet(p, avec_groupe=False):
    ps = []
    if avec_groupe:
        ps.append(para([run(p['groupe'], b=True, sz=15, color='595959')], after=0))
    ps.append(para([run(court(p['sujet']), sz=17)], after=0))
    etat = p['relance'] or p['dernier']
    if etat:
        ps.append(para([run(court(etat, 160), i=True, sz=15, color='595959')], after=0))
    return ps


def docx(out, operation, crc_prec, date_prec, date_reunion, heure, lieu, groupes, points, conv):
    W = [1100, 6000, 1100, 1800]                       # 10 000 dxa utiles (A4, marges 1 cm)
    num_reunion = int(re.sub(r'\D', '', crc_prec) or 0) + 1
    body = []
    body.append(para([run('ORDRE DU JOUR', b=True, sz=32)], align='center', after=60))
    body.append(para([run('Réunion de chantier n° %d — %s %s à %s' % (num_reunion, JOURS[date_reunion.weekday()],
                                                                     fdate(date_reunion), heure), b=True, sz=24)],
                     align='center', after=60))
    body.append(para([run(operation, sz=20)], align='center', after=60))
    if lieu:
        body.append(para([run('Lieu : %s' % lieu, sz=20)], align='center', after=60))
    body.append(para([run('Établi d\'après le compte rendu %s du %s.' % (crc_prec, fdate(date_prec)), i=True, sz=18,
                          color='595959')], align='center', after=240))

    n = 0

    def titre(t):
        nonlocal n
        n += 1
        body.append(para([run('%d. %s' % (n, t))], style='Titre1', keep=True))

    titre('Approbation du compte rendu %s du %s' % (crc_prec, fdate(date_prec)))
    body.append(para([run('Observations éventuelles des intervenants, à formuler par écrit au plus tard 8 jours '
                          'calendaires après réception du compte rendu.', sz=18)]))

    prio = [p for p in points if p['rubrique'] == 'prioritaire']
    titre('Points prioritaires (urgents et retards)')
    if prio:
        body.append(table(W, ['N°', 'Objet', 'Pour le', 'Statut'],
                          [[p['num'], cellule_objet(p, True), fdate(p['pour_le']), p['statut']] for p in prio]))
    else:
        body.append(para([run('Aucun.', i=True, sz=18)]))

    titre('Points à traiter, par intervenant')
    body.append(para([run('Échéance atteinte au %s, ou point en relance. Les points prioritaires ci-dessus ne sont pas '
                          'répétés.' % fdate(date_reunion), i=True, sz=17, color='595959')]))
    vide = True
    for g in groupes:
        pts = [p for p in g['points'] if p['rubrique'] == 'echeance']
        if not pts:
            continue
        vide = False
        body.append(para([run(g['label'])], style='Titre2', keep=True))
        body.append(table(W, ['N°', 'Objet', 'Pour le', 'Statut'],
                          [[('%s' % p['num']), cellule_objet(p), fdate(p['pour_le']), p['statut'] or '—'] for p in pts]))
    if vide:
        body.append(para([run('Aucun.', i=True, sz=18)]))

    att = [p for p in points if p['rubrique'] == 'attente']
    titre('Points en attente d\'un tiers (pour mémoire)')
    if att:
        par_org = {}
        for p in att:
            o = re.sub(r'(?i)^en attente\s*', '', p['statut']).strip() or '(organisme non précisé)'
            par_org.setdefault(o, []).append(p['num'])
        W2 = [2600, 6400, 1000]
        body.append(table(W2, ['Attendu de', 'Points concernés (N°)', 'Nombre'],
                          [[o, ', '.join(v), str(len(v))] for o, v in sorted(par_org.items(), key=lambda x: x[0].upper())]))
    else:
        body.append(para([run('Aucun.', i=True, sz=18)]))

    titre('Visite de chantier')
    titre('Questions diverses')

    requis = []
    for p in points:
        if p['rubrique'] in ('prioritaire', 'echeance'):
            o = p['organisme'] or p['groupe']
            if o and o not in requis:
                requis.append(o)
    titre('Convocations')
    body.append(para([run('Présence requise (points à traiter) : ', b=True, sz=18),
                      run(', '.join(requis) if requis else '—', sz=18)]))
    body.append(para([run('Convoqués habituels (colonne C des coordonnées) : ', b=True, sz=18),
                      run(', '.join(conv) if conv else '—', sz=18)]))

    sect = ('<w:sectPr><w:headerReference w:type="default" r:id="rIdH"/><w:footerReference w:type="default" r:id="rIdF"/>'
            '<w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1000" w:right="953" w:bottom="900" w:left="953" '
            'w:header="400" w:footer="400" w:gutter="0"/></w:sectPr>')
    NS = ('xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
          'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"')
    document = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document %s><w:body>%s%s</w:body></w:document>' % (
        NS, ''.join(body), sect)
    header = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:hdr %s>%s</w:hdr>' % (
        NS, para([run('ACAU — %s — Ordre du jour de la réunion n° %d' % (operation, num_reunion), sz=15, color='595959')],
                 align='right', after=0))
    fld = lambda f: ('<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> %s </w:instrText></w:r>'  # noqa: E731
                     '<w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>1</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r>' % f)
    footer = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:ftr %s><w:p><w:pPr><w:jc w:val="right"/></w:pPr>'
              '%s%s%s</w:p></w:ftr>' % (NS, run('Page ', sz=15), fld('PAGE'), run(' / ', sz=15) + fld('NUMPAGES')))
    styles = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles %s>'
              '<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="{f}" w:hAnsi="{f}" w:cs="{f}" w:eastAsia="{f}"/>'
              '<w:sz w:val="18"/><w:szCs w:val="18"/><w:lang w:val="fr-FR"/></w:rPr></w:rPrDefault>'
              '<w:pPrDefault><w:pPr><w:spacing w:after="80" w:line="252" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults>'
              '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>'
              '<w:style w:type="paragraph" w:styleId="Titre1"><w:name w:val="heading 1"/><w:basedOn w:val="Normal"/>'
              '<w:next w:val="Normal"/><w:pPr><w:keepNext/><w:spacing w:before="280" w:after="100"/><w:outlineLvl w:val="0"/>'
              '<w:pBdr><w:bottom w:val="single" w:sz="6" w:space="1" w:color="000000"/></w:pBdr></w:pPr>'
              '<w:rPr><w:b/><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr></w:style>'
              '<w:style w:type="paragraph" w:styleId="Titre2"><w:name w:val="heading 2"/><w:basedOn w:val="Normal"/>'
              '<w:next w:val="Normal"/><w:pPr><w:keepNext/><w:spacing w:before="160" w:after="60"/><w:outlineLvl w:val="1"/></w:pPr>'
              '<w:rPr><w:b/><w:sz w:val="20"/><w:szCs w:val="20"/></w:rPr></w:style>'
              '</w:styles>' % NS).replace('{f}', POLICE)
    ct = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
          '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
          '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
          '<Default Extension="xml" ContentType="application/xml"/>'
          '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
          '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>'
          '<Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>'
          '<Override PartName="/word/footer1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/>'
          '</Types>')
    R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
    rels = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            '<Relationship Id="rId1" Type="%s/officeDocument" Target="word/document.xml"/></Relationships>' % R)
    drels = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
             '<Relationship Id="rIdS" Type="%s/styles" Target="styles.xml"/>'
             '<Relationship Id="rIdH" Type="%s/header" Target="header1.xml"/>'
             '<Relationship Id="rIdF" Type="%s/footer" Target="footer1.xml"/></Relationships>' % (R, R, R))
    with zipfile.ZipFile(out, 'w', zipfile.ZIP_DEFLATED) as z:
        z.writestr('[Content_Types].xml', ct)
        z.writestr('_rels/.rels', rels)
        z.writestr('word/document.xml', document)
        z.writestr('word/styles.xml', styles)
        z.writestr('word/header1.xml', header)
        z.writestr('word/footer1.xml', footer)
        z.writestr('word/_rels/document.xml.rels', drels)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('cr')
    ap.add_argument('out')
    ap.add_argument('--date')
    ap.add_argument('--heure')
    ap.add_argument('--lieu')
    ap.add_argument('--json')
    a = ap.parse_args()
    pg = openpyxl.load_workbook(a.cr, data_only=True)['Page de garde']
    refs = K.page_de_garde_ws(pg)                 # « CRC-NN » repéré (A22 à HONGUEMARE, A21 à DUCLAIR)
    crc, date_prec = str(pg[refs['crc']].value).strip(), pg[refs['date']].value
    date_reunion = datetime.datetime.strptime(a.date, '%Y-%m-%d') if a.date else date_prec + datetime.timedelta(days=7)
    heure = a.heure or str(pg[refs['heure']].value or '').strip() or 'heure à préciser'
    lieu = a.lieu or (' '.join(str(pg[refs['lieu']].value).split()) if refs.get('lieu') and pg[refs['lieu']].value else '')
    operation = refs['operation']
    groupes, points, conv = lire(a.cr, date_reunion)
    docx(a.out, operation, crc, date_prec, date_reunion, heure, lieu, groupes, points, conv)
    resume = {r: len([p for p in points if p['rubrique'] == r]) for r in ('prioritaire', 'echeance', 'attente')}
    print('OK', a.out, 'réunion du', fdate(date_reunion), resume)
    if a.json:
        json.dump(points, open(a.json, 'w', encoding='utf-8'), ensure_ascii=False, indent=1, default=str)


if __name__ == '__main__':
    main()
