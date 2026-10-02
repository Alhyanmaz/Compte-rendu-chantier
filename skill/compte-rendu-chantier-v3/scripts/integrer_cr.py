# -*- coding: utf-8 -*-
"""Intégration hebdomadaire dans un CR au nouveau format (colonne N°, HISTORIQUE…).

Méthode (30/09/2026) : José pose des notes de cellule sur le CR précédent pendant la réunion ; Claude les reformule
à l'appui de la transcription audio et écrit OPERATIONS.json ; ce script l'applique. Le CR produit est relu et
corrigé par José dans Excel, puis repris par relire_cr.py avant l'édition finale.

Usage : python integrer_cr.py CR_PRECEDENT.xlsx OPERATIONS.json CR_NOUVEAU.xlsx RAPPORT.json
puis    python ajuster_hauteurs.py CR_NOUVEAU.xlsx CR_FINAL.xlsx

OPERATIONS.json :
{
  "date_cr": "2026-09-29", "crc": "CRC-16",
  "operations": [
    {"type": "maj", "num": "02-066", "texte": "…", "pour_le": "2026-10-02" | "+7", "fait_le": "URGENT" | "2026-09-22"},
    {"type": "nouvelle", "onglet": "02 -GROS OEUVRE", "code": "02", "section": "ÉTUDES",
     "texte": "…", "pour_le": "+7" | "2026-10-06" | null, "fait_le": "PM" | null}
  ],
  "non_route": [{"extrait": "passage de la transcription écarté", "raison": "…", "note": "Onglet!B12" | null}]
}
Champs facultatifs de chaque opération :
  "note"    : "Onglet!B12", note de cellule de José traitée par l'opération (supprimée du classeur) ;
  "source"  : "note" (défaut) | "transcription" (point sans note de José : fond rose) ;
  "routage" : texte de la colonne ROUTAGE (note d'origine, extrait de l'audio, choix de Claude) ;
  "doute"   : question [?] à trancher par José (fond rose soutenu, écrite dans ROUTAGE).
Toute note de cellule d'un onglet d'observations doit être citée par une opération ou par « non_route ».

Règles appliquées : docs/decisions-HONGUEMARE.md. Édition XML chirurgicale (cr_xml.py) ; openpyxl en
lecture seule. Le texte affiché (OBSERVATIONS) est recalculé depuis HISTORIQUE, qui fait foi.
"""
import copy
import datetime
import json
import os
import re
import sys


sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from cr_xml import Book, _set_color, _rpr_from_font, col_letter  # noqa: E402
import cr_classeur as K  # noqa: E402

BASE, OPS, OUT, REPORT = sys.argv[1:5]
spec = json.load(open(OPS, encoding='utf-8'))
K.init(Book(BASE), datetime.datetime.strptime(spec['date_cr'], '%Y-%m-%d'))
from cr_classeur import *  # noqa: E402,F401,F403  (b, DATE, DSTR et les outils communs)
report = {'masquees': [], 'maj': [], 'nouvelles': [], 'relances_auto': [], 'reouvertures': [], 'insertions': [],
          'alertes': []}

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
        for fo in cf.iter(N + 'formula'):             # la formule suit la 1re cellule de sqref (bug du CRC-16 v1)
            fo.text = shift_txt(fo.text, p, n)
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
pg = b.ws('Page de garde')     # cellule « CRC-NN » (A22 à HONGUEMARE, A21 à DUCLAIR) ; la date est à sa droite
c_crc = next((c for c in pg.iter(N + 'c') if CRC_RE.match(text_of(c).strip())), None)
if c_crc is None:
    raise SystemExit('Page de garde : cellule « CRC-NN » introuvable')
row_pg = c_crc.getparent()
set_si(c_crc, si_new([(_rpr_from_font(b._font_of(c_crc.get('s'))), spec['crc'])]))
set_num(cell(row_pg, col_letter(col_index(c_crc.get('r')) + 1)), serial(DATE))
b.touch(b.sheets['Page de garde'])

SHEETS = obs_sheets()
# MFC de police grise retirée (30/09/2026) : elle écrasait la typo manuelle de José et le rouge des dates du jour.
# Le gris des lignes PM / soldées est écrit dans les cellules. Les fonds jaune (URGENT) et orange (Retard) restent.
for sh in SHEETS:
    ws = b.ws(sh)
    for cf in ws.findall(N + 'conditionalFormatting'):
        for rule in cf.findall(N + 'cfRule'):
            fo = rule.find(N + 'formula')
            if fo is not None and fo.text and fo.text.startswith('OR($E') and '"PM"' in fo.text:
                cf.remove(rule)
        if not cf.findall(N + 'cfRule'):
            ws.remove(cf)
    b.touch(b.sheets[sh])
ops = spec['operations']
touched = {o['num'] for o in ops if o['type'] == 'maj'}

# ================================================================== 1 bis. notes de cellule : toutes traitées, puis supprimées
# (avant toute insertion de lignes : les notes ne suivent pas les décalages)
cited = {x['note'] for x in ops + spec.get('non_route', []) if x.get('note')}
restantes = ['%s!%s' % (sh, ref) for sh in SHEETS for ref in b.notes(sh) if '%s!%s' % (sh, ref) not in cited]
if restantes:
    raise SystemExit('Notes de cellule non traitées (à citer dans une opération ou dans non_route) : %s' % restantes)
for x in sorted(cited):
    sh, ref = x.split('!', 1)
    if ref not in b.notes(sh):
        report['alertes'].append('Note citée introuvable : %s' % x)
    b.delete_note(sh, ref)
report['notes'] = sorted(cited)


def marquer(row, o, defaut=None):
    """Colonne ROUTAGE ; fond rose si le point vient de la seule transcription, rose soutenu si [?]."""
    txt = o.get('routage') or defaut or ''
    if o.get('doute'):
        txt = (txt + '\n' if txt else '') + '[?] ' + o['doute']
    if txt:
        ecrire_routage(row, txt)
    rose = ROSE_DOUTE if o.get('doute') else (ROSE if o.get('source') == 'transcription' else None)
    if rose:
        for col in 'ABCDE':
            set_fill(cell(row, col), rose)

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
        cA, cB, cC, cD, cE, cF = (cell(row, x) for x in 'ABCDEF')
        cJ = row.find(N + 'c[@r="%s%d"]' % (ROUT_COL, r))
        if cJ is not None and text_of(cJ):             # ROUTAGE d'un passage précédent (normalement vidé en relecture)
            clear(cJ)
        e_txt = text_of(cE).strip()
        e_num = number_of(cE)
        closed = e_num is not None or e_txt.upper() in TERMINAUX
        # -- masquage des lignes soldées lors d'un passage précédent (jamais une ligne traitée ce jour, jamais PM)
        if closed and not it['hidden'] and num not in touched:
            row.set('hidden', '1')
            report['masquees'].append([sh, num, from_serial(e_num).strftime('%d/%m/%Y') if e_num is not None else e_txt])
            continue
        # -- rouge du passage précédent -> noir dans HISTORIQUE (non imprimé) ; le gris y est aussi ramené au noir
        f_runs = runs_of(cF)
        if any(color_of(rp) in (RED, 'GREY') for rp, _ in f_runs):
            set_si(cF, si_new(recolor_runs(f_runs, BLACK)))
        # -- mise à jour du jour
        op = next((o for o in ops if o['type'] == 'maj' and o['num'] == num), None)
        add = None
        today = set()                                  # colonnes écrites ce jour (restent rouges)
        if op:
            add = op['texte'].strip()
            if op.get('pour_le'):
                write_pour_le(cD, op['pour_le'], r)
                today.add('D')
            if op.get('fait_le'):
                write_fait_le(cE, op['fait_le'])
                today.add('E')
            elif e_num is not None:                    # réouverture (règle V1) : point annoté alors qu'il était soldé
                clear(cE)
                report['reouvertures'].append([sh, num])
            report['maj'].append([sh, num, add[:90]])
            marquer(row, op)
        elif not it['hidden'] and not closed:
            # -- relance automatique (échéance atteinte ; jamais sur PM, En attente, En cours, Retard)
            up = e_txt.upper()
            pl = number_of(cD)
            due = pl is not None and pl <= serial(DATE)
            if up == 'URGENT' or (up in ('', 'RELANCE') and due):
                add = 'Relance'
                if up == '':
                    write_fait_le(cE, 'Relance')
                    today.add('E')
                report['relances_auto'].append([sh, num, up or '(vide)'])
                ecrire_routage(row, 'Relance automatique : %s' % ('URGENT, relancé à chaque CR' if up == 'URGENT' else
                                                                  'échéance du %s atteinte' % from_serial(pl).strftime('%d/%m/%Y')))
        if add:
            f_runs = runs_of(cF)
            base = copy.deepcopy(f_runs[-1][0]) if f_runs else _rpr_from_font(b._font_of(cF.get('s')))
            base = plain_rpr(base)
            _set_color(base, RED)
            if f_runs and not f_runs[-1][1].rstrip().endswith(('.', ':', ';')):
                f_runs[-1] = (f_runs[-1][0], f_runs[-1][1].rstrip() + '.')     # le point prend la couleur du texte précédent
            f_runs.append((base, ' Au %s %s' % (DSTR, add)))
            set_si(cF, si_new(f_runs))
        # -- couleurs explicites (pas de MFC de police : José garde la main sur la typo) :
        #    ajout du jour rouge ; ligne PM / soldée / sans objet grise ; sinon noir ; statut rouge (PM gris)
        grey = is_grey_status(cE)
        base_col = GREY if grey else BLACK
        e_now = text_of(cE).strip().upper()
        open_ = number_of(cE) is None and e_now not in TERMINAUX | {'PM'}
        old_b = runs_of(cB)
        if add:
            # texte affiché recalculé depuis HISTORIQUE (règle A, « sans réponse depuis », objet d'attente conservé)
            objet = [(rp, t) for rp, t in old_b if t.startswith('\n→ En attente :')]
            ab = number_of(cC)
            res = condense(runs_of(cF), from_serial(ab) if ab else None, open_, date_cr=DATE)
            if res:
                runs = recolor_runs(res[0], base_col, keep_red=True)
                if objet and 'EN ATTENTE' in e_now:
                    runs += recolor_runs(objet, base_col)
                set_si(cB, si_new(runs))
        elif old_b:
            # ligne non traitée ce jour : texte affiché conservé tel quel (corrections manuelles de José),
            # seuls les couleurs et le compteur « sans réponse depuis » sont mis à jour
            runs = recolor_runs(old_b, base_col)
            m = next((i for i, (_, t) in enumerate(runs) if '→ Relancé' in t), None)
            if m is not None and open_:
                ab = number_of(cC)
                res = condense(runs_of(cF), from_serial(ab) if ab else None, open_, date_cr=DATE)
                rel = [t for _, t in (res[0] if res else []) if '→ Relancé' in t]
                if rel:
                    runs[m] = (runs[m][0], rel[-1])
            set_si(cB, si_new(runs))
        appliquer_retard(row, r)
        color_cell(cA, base_col)
        for col, c in (('C', cC), ('D', cD)):
            color_cell(c, RED if col in today else base_col)
        if 'E' not in today:
            if number_of(cE) is not None:
                color_cell(cE, GREY)
            elif e_now == 'PM':
                color_cell(cE, GREY)
            elif e_now:
                color_cell(cE, RED)
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
        rpr_a = _rpr_from_font(b._font_of(cA.get('s')))
        _set_color(rpr_a, BLACK)                   # N° toujours noir (le style d'une ligne de réserve peut être gris)
        set_si(cA, si_new([(rpr_a, num)]))
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
        jours_formula(cI, r)
        report['nouvelles'].append([sh, num, o.get('section') or code, o['texte'][:90]])
        marquer(row, o)
    colonne_routage(sh)
    finir_onglet(sh)

# ================================================================== 4. Référentiel : prochain N°, enregistrement
maj_referentiel(next_num)
onglet_non_route(spec.get('non_route', []), visible=True)
terminer(OUT)
json.dump(report, open(REPORT, 'w', encoding='utf-8'), ensure_ascii=False, indent=1)
print('OK', OUT, {k: len(v) for k, v in report.items()})
