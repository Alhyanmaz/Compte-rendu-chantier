# -*- coding: utf-8 -*-
"""Mode relecture : reprend un CR corrigé à la main par José avant diffusion, sans logique de nouvelle semaine.

Usage : python relire_cr.py CR_LIVRE.xlsx CR_CORRIGE.xlsx CR_SORTIE.xlsx RAPPORT.json
puis    python ajuster_hauteurs.py CR_SORTIE.xlsx CR_FINAL.xlsx

CR_LIVRE   : la version livrée par Claude (référence, pour savoir ce que José a changé).
CR_CORRIGE : la même version, corrigée par José dans Excel.

Ce que fait le script (docs/decisions-HONGUEMARE.md, « Mode relecture ») :
- ni masquage, ni relance, ni remise au noir du rouge du jour ;
- HISTORIQUE (F) modifié : les « Au JJ/MM/AAAA » à la date du CR passent en rouge et OBSERVATIONS (B) est
  recalculé depuis F (règle A, compteur de relances). B et F modifiés tous les deux : F l'emporte, signalé ;
- B seul modifié : B conservé, signalé (à reporter dans HISTORIQUE, sinon perdu à la prochaine mise à jour) ;
- ABORDÉ LE / POUR LE / FAIT LE modifiés : en rouge ; statut saisi à la main remis à la casse de la liste ;
- ligne passée en PM / soldée : texte gris, sauf l'ajout du jour ; compteur « sans réponse depuis » mis à jour ;
- ligne nouvelle sans N° (ligne de réserve remplie) : N° attribué, style du tableau, SECTION, formule
  JOURS RETARD, texte en rouge, ABORDÉ LE = date du CR s'il est vide ;
- Retard : DATE RETARD et ligne « → En retard depuis le … » ; lignes vides et sauts de page remis à jour.
"""
import copy
import datetime
import json
import os
import re
import sys

import openpyxl

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from cr_xml import Book, _set_color, _rpr_from_font  # noqa: E402
from cr_texte import AU, full_date, slice_runs  # noqa: E402
import cr_classeur as K  # noqa: E402

LIVRE, CORRIGE, OUT, REPORT = sys.argv[1:5]

# ------------------------------------------------------------------ date du CR : page de garde B22 du fichier corrigé
_pg = openpyxl.load_workbook(CORRIGE, data_only=True)['Page de garde']['B22'].value
if not isinstance(_pg, datetime.datetime):
    raise SystemExit('Date du CR introuvable en Page de garde!B22 : %r' % _pg)
K.init(Book(CORRIGE), _pg)
from cr_classeur import *  # noqa: E402,F401,F403

report = {'date_cr': DSTR, 'modifiees': [], 'nouvelles': [], 'alertes': []}


# ------------------------------------------------------------------ lecture comparée (valeurs et formules)
def norm_txt(v):
    if v is None:
        return ''
    return '\n'.join(' '.join(x.split()) for x in str(v).replace('\r\n', '\n').split('\n')).strip()


def norm_f(v):
    return re.sub(r'\d+', '#', v) if isinstance(v, str) and v.startswith('=') else v


def snapshot(path):
    wv = openpyxl.load_workbook(path, data_only=True)
    wf = openpyxl.load_workbook(path, data_only=False)
    out = {}
    for ws in wv.worksheets:
        if not any(str(ws.cell(r, 1).value).strip() == 'N°' for r in range(1, 12)):
            continue
        wsf = wf[ws.title]
        for r in range(1, ws.max_row + 1):
            a = ws.cell(r, 1).value
            if isinstance(a, str) and CODE_RE.match(a.strip()):
                d = {'row': r}
                for col in 'BCDEF':
                    v, f = ws['%s%d' % (col, r)].value, wsf['%s%d' % (col, r)].value
                    d[col] = (norm_txt(v) if col in 'BF' or isinstance(v, str) else v, norm_f(f))
                out.setdefault((ws.title, a.strip()), d)
    return out


def differe(p, q):
    """p, q : (valeur, formule). Valeur en cache absente d'un côté : on compare les formules."""
    if p[0] is None or q[0] is None:
        return p[1] != q[1]
    return p[0] != q[0]


ref = snapshot(LIVRE)
cur = snapshot(CORRIGE)
for (sh, num) in sorted(set(ref) - set(cur)):
    report['alertes'].append('%s %s : ligne absente du fichier corrigé (supprimée ?)' % (sh, num))


def today_red(runs):
    """Les segments « Au JJ/MM/AAAA » datés du jour du CR passent en rouge (jusqu'au segment suivant)."""
    text = ''.join(t for _, t in runs)
    segs = [(m.start(), full_date(m.group(1))) for m in AU.finditer(text)]
    spans = []
    for k, (st, d) in enumerate(segs):
        if d == DSTR:
            spans.append((st, segs[k + 1][0] if k + 1 < len(segs) else len(text)))
    if not spans:
        return runs, False
    cuts = sorted({0, len(text)} | {x for s in spans for x in s})
    out = []
    for a, z in zip(cuts, cuts[1:]):
        for rpr, t in slice_runs(runs, a, z):
            rpr = copy.deepcopy(rpr)
            if any(s <= a < e for s, e in spans):
                _set_color(rpr, RED)
            out.append((rpr, t))
    return out, True


def set_statut(c, v):
    rpr = plain_rpr(_rpr_from_font(b._font_of(c.get('s'))))
    _set_color(rpr, RED)
    set_si(c, si_new([(rpr, v)]))


# ------------------------------------------------------------------ traitement
SHEETS = obs_sheets()
next_num = numeros(SHEETS)
seen = {}
for sh in SHEETS:
    rows, tabs, info = model(sh)
    for r, it in sorted(info.items()):
        row = rows.get(r)
        if row is None or it.get('hidden') or it['kind'] not in ('obs', 'empty'):
            continue
        cA, cB, cC, cD, cE, cF = (cell(row, x) for x in 'ABCDEF')
        # ---------------------------------------------------------- ligne nouvelle sans N° (réserve remplie)
        if it['kind'] == 'empty':
            tb, tf = text_of(cB).strip(), text_of(cF).strip()
            if not (tb or tf) or tb.upper() in ('ÉTUDES', 'TRAVAUX'):
                continue
            t = [t for t in tabs if t['name'] == it['table']][0]
            if r == t['r1']:
                report['alertes'].append('%s ligne %d : écrite dans la dernière ligne du tableau (trait épais de fin '
                                         'de tableau) ; prévoir une ligne vide après' % (sh, r))
            code = t['code']
            next_num[code] = next_num.get(code, 0) + 1
            num = '%s-%03d' % (code, next_num[code])
            tmpl = max([rr for rr, x in info.items() if x['kind'] == 'obs' and x['table'] == t['name']] or [r])
            for col in 'ABCDEFGHI':
                src = rows[tmpl].find(N + 'c[@r="%s%d"]' % (col, tmpl))
                if src is not None and src.get('s'):
                    cell(row, col).set('s', src.get('s'))
            set_si(cA, si_new([(_rpr_from_font(b._font_of(cA.get('s'))), num)]))
            restyle(cA, color=BLACK)
            runs = K.runs_of(cB) or K.runs_of(cF)
            for rp, _ in runs:
                _set_color(rp, RED)
            set_si(cB, si_new(runs))
            set_si(cF, si_new([(copy.deepcopy(rp), x) for rp, x in runs]))
            if number_of(cC) is None:
                set_num(cC, serial(DATE))
            for c in (cC, cD):
                if number_of(c) is not None:
                    restyle(c, color=RED)
            e = text_of(cE).strip()
            if e and number_of(cE) is None:
                set_statut(cE, statut_canonique(e))
            elif number_of(cE) is not None:
                restyle(cE, color=RED)
            set_si(cell(row, 'H'), si_new([(None, it.get('section') or code)]))
            jours_formula(cell(row, 'I'), r)
            appliquer_retard(row, r)
            report['nouvelles'].append([sh, num, it.get('section') or code, (tb or tf)[:90]])
            continue
        # ---------------------------------------------------------- ligne existante
        num = it['num']
        if num in seen:
            report['alertes'].append('%s : N° %s en double (lignes %d et %d)' % (sh, num, seen[num], r))
        seen[num] = r
        p, q = ref.get((sh, num)), cur.get((sh, num))
        if p is None:                                  # N° saisi à la main : ligne traitée comme nouvelle
            ch = set('BCDEF')
            report['alertes'].append('%s %s : N° absent de la version livrée, ligne traitée comme ajout du jour'
                                     % (sh, num))
        else:
            ch = {col for col in 'BCDEF' if differe(p[col], q[col])}
        if not ch:
            continue
        report['modifiees'].append([sh, num, ''.join(sorted(ch))])
        e = text_of(cE).strip()
        if 'E' in ch and e and number_of(cE) is None and statut_canonique(e) != e:
            set_statut(cE, statut_canonique(e))
        e_now = text_of(cE).strip().upper()
        if 'E' in ch and e_now == 'RELANCE' and 'Au %s Relance' % DSTR not in text_of(cF):
            # relance posée à la main : inscrite dans HISTORIQUE comme une relance automatique
            f_runs = K.runs_of(cF)
            base = plain_rpr(copy.deepcopy(f_runs[-1][0]) if f_runs else _rpr_from_font(b._font_of(cF.get('s'))))
            _set_color(base, RED)
            if f_runs and not f_runs[-1][1].rstrip().endswith(('.', ':', ';')):
                f_runs[-1] = (f_runs[-1][0], f_runs[-1][1].rstrip() + '.')
            set_si(cF, si_new(f_runs + [(base, ' Au %s Relance' % DSTR)]))
            ch.add('F')
        grey = is_grey_status(cE)
        base_col = GREY if grey else BLACK
        open_ = number_of(cE) is None and e_now not in TERMINAUX | {'PM'}
        old_b = K.runs_of(cB)
        if 'F' in ch:
            f_runs, _ = today_red(K.runs_of(cF))
            set_si(cF, si_new(f_runs))
            if 'B' in ch:
                report['alertes'].append('%s %s : OBSERVATIONS et HISTORIQUE modifiés tous les deux ; le texte '
                                         'imprimé est recalculé depuis HISTORIQUE' % (sh, num))
            objet = [(rp, x) for rp, x in old_b if x.startswith('\n→ En attente :')]
            ab = number_of(cC)
            res = condense(f_runs, from_serial(ab) if ab else None, open_, date_cr=DATE)
            if res:
                runs = recolor_runs(res[0], base_col, keep_red=True)
                if objet and 'EN ATTENTE' in e_now:
                    runs += recolor_runs(objet, base_col, keep_red=True)
                set_si(cB, si_new(runs))
        else:
            if 'B' in ch:
                report['alertes'].append('%s %s : OBSERVATIONS modifié sans HISTORIQUE ; texte conservé, mais à '
                                         'reporter dans HISTORIQUE (sinon perdu à la prochaine mise à jour de la '
                                         'ligne)' % (sh, num))
            runs = recolor_runs(old_b, base_col, keep_red=True)
            m = next((i for i, (_, x) in enumerate(runs) if '→ Relancé' in x), None)
            if m is not None:                          # compteur « sans réponse depuis » : retiré si la ligne est soldée
                ab = number_of(cC)
                res = condense(K.runs_of(cF), from_serial(ab) if ab else None, open_, date_cr=DATE)
                rel = [x for _, x in (res[0] if res else []) if '→ Relancé' in x]
                if rel:
                    runs[m] = (runs[m][0], rel[-1])
            set_si(cB, si_new(runs))
        appliquer_retard(row, r)
        color_cell(cA, base_col)
        for col, c in (('C', cC), ('D', cD)):
            if col in ch:
                color_cell(c, RED)
            elif K.color_of(_rpr_from_font(b._font_of(c.get('s')))) != RED:
                color_cell(c, base_col)
        if 'E' in ch:
            color_cell(cE, RED)
        elif K.color_of(_rpr_from_font(b._font_of(cE.get('s')))) != RED and not any(
                K.color_of(rp) == RED for rp, _ in K.runs_of(cE)):
            color_cell(cE, GREY if (number_of(cE) is not None or e_now == 'PM') else RED)
    finir_onglet(sh)

maj_referentiel(next_num)
terminer(OUT)
json.dump(report, open(REPORT, 'w', encoding='utf-8'), ensure_ascii=False, indent=1)
print('OK', OUT, {k: (len(v) if isinstance(v, list) else v) for k, v in report.items()})
