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
- Retard : DATE RETARD et ligne « → En retard depuis le … » ; lignes vides et sauts de page remis à jour ;
- édition finale : fonds roses retirés (un [?] laissé tel quel est signalé), colonne ROUTAGE vidée,
  onglet « Non routé » vidé et masqué ; repère bleu des doublons de la migration retiré.
"""
import copy
import datetime
import difflib
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

# ------------------------------------------------------------------ date du CR : page de garde du fichier corrigé
_ws_pg = openpyxl.load_workbook(CORRIGE, data_only=True)['Page de garde']
_pg = _ws_pg[K.page_de_garde_ws(_ws_pg)['date']].value       # à droite de « CRC-NN » (B22 ou B21 selon l'opération)
if not isinstance(_pg, datetime.datetime):
    raise SystemExit('Date du CR introuvable en Page de garde : %r' % _pg)
K.init(Book(CORRIGE), _pg)
from cr_classeur import *  # noqa: E402,F401,F403

report = {'date_cr': DSTR, 'modifiees': [], 'nouvelles': [], 'alertes': [], 'dates_posterieures': []}


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
                    if col in 'BF':
                        d['raw' + col] = '' if v is None else str(v).replace('\r\n', '\n')
                out.setdefault((ws.title, a.strip()), d)
    return out


def differe(p, q):
    """p, q : (valeur, formule). Formule identique (au décalage de ligne près) : pas une modification de José,
    même si Excel a recalculé une autre valeur. Valeur en cache absente d'un côté : on compare les formules."""
    if isinstance(p[1], str) and p[1].startswith('=') and p[1] == q[1]:
        return False
    if p[0] is None or q[0] is None:
        return p[1] != q[1]
    return p[0] != q[0]


ref = snapshot(LIVRE)
cur = snapshot(CORRIGE)
for (sh, num) in sorted(set(ref) - set(cur)):
    report['alertes'].append('%s %s : ligne supprimée dans le fichier corrigé (N° non réutilisé ; historique perdu)'
                             % (sh, num))


def red_spans(runs, old_text=None):
    """Plages à mettre en rouge : segments « Au JJ/MM/AAAA » datés du jour du CR, et tout texte ajouté par José
    par rapport à la version livrée (quelle que soit la date qu'il a écrite : 25/09, 30/09, 02/10…)."""
    text = ''.join(t for _, t in runs)
    segs = [(m.start(), full_date(m.group(1))) for m in AU.finditer(text)]
    spans = [(st, segs[k + 1][0] if k + 1 < len(segs) else len(text)) for k, (st, d) in enumerate(segs) if d == DSTR]
    if old_text is not None:
        sm = difflib.SequenceMatcher(None, old_text, text, autojunk=False)
        spans += [(j1, j2) for tag, i1, i2, j1, j2 in sm.get_opcodes()
                  if tag in ('insert', 'replace') and text[j1:j2].strip()]
    return spans


def paint(runs, spans, color=RED):
    if not spans:
        return runs
    text = ''.join(t for _, t in runs)
    cuts = sorted({0, len(text)} | {x for sp in spans for x in sp})
    out = []
    for a, z in zip(cuts, cuts[1:]):
        for rpr, t in slice_runs(runs, a, z):
            rpr = copy.deepcopy(rpr)
            if any(s0 <= a < e0 for s0, e0 in spans):
                _set_color(rpr, color)
            out.append((rpr, t))
    return out


def replace_span(runs, a, z, new, color=RED):
    """Remplace text[a:z] par new (texte ajouté en rouge) en gardant la mise en forme du reste."""
    text = ''.join(t for _, t in runs)
    out = slice_runs(runs, 0, a)
    if new:
        rpr = copy.deepcopy(runs[0][0] if not out else out[-1][0])
        _set_color(rpr, color)
        out.append((rpr, new))
    return out + slice_runs(runs, z, len(text))


def _ctx(s, a, z, side, n):
    """Contexte d'une modification, limité au segment (pas de « → » ni de saut de ligne : B et F diffèrent là)."""
    c = s[max(0, a - n):a] if side == 'g' else s[z:z + n]
    cut = [i for i, ch in enumerate(c) if ch in '\n→']
    if side == 'g':
        return c[cut[-1] + 1:] if cut else c
    return c[:cut[0]] if cut else c


def trouver(f, pat):
    """Occurrences de pat dans f, les suites d'espaces comptant pour un seul (José tape souvent deux espaces)."""
    toks = pat.split()
    if not toks:
        return []
    rx = r'\s+'.join(re.escape(t) for t in toks)
    if pat[:1].isspace():
        rx = r'\s+' + rx
    if pat[-1:].isspace():
        rx += r'\s+'
    return [(m.start(), m.end()) for m in re.finditer(rx, f)]


def propager(old_b, new_b, f_runs):
    """Reporte dans HISTORIQUE les corrections faites par José dans OBSERVATIONS (le texte imprimé est recalculé
    depuis HISTORIQUE : sans report, elles seraient perdues). Renvoie (runs, nb reportées, nb impossibles)."""
    sm = difflib.SequenceMatcher(None, old_b, new_b, autojunk=False)
    ok = ko = 0
    for tag, i1, i2, j1, j2 in reversed(sm.get_opcodes()):
        if tag == 'equal':
            continue
        o, n = old_b[i1:i2], new_b[j1:j2]
        if not o.strip() and not n.strip():
            continue                                   # mise en page seule (espaces, sauts de ligne)
        if any(x in o + n for x in ('→', '[...]')):
            continue                                   # lignes calculées (relances, « [...] ») : rien à reporter
        if not n.strip():
            n = ''
        f = ''.join(t for _, t in f_runs)
        done = False
        for k in (25, 12, 6):
            gn, dn = _ctx(new_b, j1, j2, 'g', k), _ctx(new_b, j1, j2, 'd', k)
            if n and (gn.strip() or dn.strip()) and len(trouver(f, gn + n + dn)) == 1:
                done = True                            # déjà reporté par José dans HISTORIQUE
                break
            if not n and gn.strip() and dn.strip() and len(trouver(f, gn + dn)) == 1:
                done = True                            # suppression déjà faite dans HISTORIQUE
                break
            g, d = _ctx(old_b, i1, i2, 'g', k), _ctx(old_b, i1, i2, 'd', k)
            if o and not (g.strip() or d.strip() or len(o) >= 8):
                continue
            if not o and not (g.strip() or d.strip()):
                continue
            occ = trouver(f, g + o + d) if o.strip() else trouver(f, g + d)
            if len(occ) == 1:
                a0, z0 = occ[0]
                if o.strip():                          # remplacement / suppression : bornes de o dans l'occurrence
                    sub = f[a0:z0]
                    oo = [x for x in trouver(sub, o)]
                    pos0, pos1 = a0 + oo[0][0], a0 + oo[0][1]
                else:                                  # insertion : après le contexte gauche
                    pos0 = pos1 = a0 + (trouver(f[a0:z0], g)[0][1] if g.strip() else 0)
                f_runs = replace_span(f_runs, pos0, pos1, n)
                done = True
                break
        ok += done
        ko += not done
    return f_runs, ok, ko


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
            rpr_a = _rpr_from_font(b._font_of(cA.get('s')))
            _set_color(rpr_a, BLACK)
            set_si(cA, si_new([(rpr_a, num)]))
            restyle(cA, color=BLACK)
            runs = K.runs_of(cB) or K.runs_of(cF)
            for rp, _ in runs:
                _set_color(rp, RED)
            ab0 = from_serial(number_of(cC)) if number_of(cC) is not None else DATE
            res0 = condense(runs, ab0, True, date_cr=DATE)     # une ligne neuve ne commence pas par « Au [date] »
            if res0 and res0[0]:
                runs = res0[0]
            set_si(cB, si_new(runs))
            set_si(cF, si_new([(copy.deepcopy(rp), x) for rp, x in runs]))
            if number_of(cC) is None:
                set_num(cC, serial(DATE))
            for c in (cC, cD):
                if number_of(c) is not None:
                    restyle(c, color=RED)
            if text_of(cD).strip().upper() == 'PM':       # PM saisi dans POUR LE : c'est un statut de FAIT LE
                clear(cD)
                if not text_of(cE).strip() and number_of(cE) is None:
                    set_si(cE, si_new([(None, 'PM')]))
                report['alertes'].append('%s ligne %d : PM déplacé de POUR LE vers FAIT LE' % (sh, r))
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
        if text_of(cD).strip().upper() == 'PM':           # PM saisi dans POUR LE (habitude V1) : c'est un statut
            clear(cD)
            if not text_of(cE).strip() and number_of(cE) is None:
                set_statut(cE, 'PM')
            ch |= {'D', 'E'}
            report['alertes'].append('%s %s : PM déplacé de POUR LE vers FAIT LE' % (sh, num))
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
        b_ok = True
        if 'B' in ch and p is not None:
            f_new, nok, nko = propager(p['rawB'], q['rawB'], K.runs_of(cF))
            if nok:
                set_si(cF, si_new(f_new))
                ch.add('F')
            if nko:
                b_ok = False
                report['alertes'].append('%s %s : correction d\'OBSERVATIONS impossible à reporter dans HISTORIQUE ; '
                                         'texte imprimé conservé tel que corrigé, à reporter à la main dans HISTORIQUE'
                                         % (sh, num))
        if 'F' in ch and b_ok:
            f_runs = paint(K.runs_of(cF), red_spans(K.runs_of(cF), p['rawF'] if p else None))
            set_si(cF, si_new(f_runs))
            ftxt = ''.join(t for _, t in f_runs)
            apres = [m.group(1) for m in AU.finditer(ftxt)
                     if datetime.datetime.strptime(full_date(m.group(1)), '%d/%m/%Y') > DATE]
            if apres:
                report['dates_posterieures'].append('%s %s : %s' % (sh, num, ', '.join(sorted(set(apres)))))
            objet = [(rp, x) for rp, x in old_b if x.startswith('\n→ En attente :')]
            ab = number_of(cC)
            res = condense(f_runs, from_serial(ab) if ab else None, open_, date_cr=DATE)
            if res:
                runs = recolor_runs(res[0], base_col, keep_red=True)
                if objet and 'EN ATTENTE' in e_now:
                    runs += recolor_runs(objet, base_col, keep_red=True)
                set_si(cB, si_new(runs))
        else:
            if 'F' in ch:
                set_si(cF, si_new(paint(K.runs_of(cF), red_spans(K.runs_of(cF), p['rawF'] if p else None))))
            if 'B' in ch and p is not None:
                old_b = paint(old_b, red_spans(old_b, p['rawB']))     # texte ajouté par José dans B : rouge
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
    # ---------------------------------------------------------- édition finale : rose retiré, ROUTAGE vidée
    rows, tabs, info = model(sh)
    for r, row in rows.items():
        for c in row:                                  # repère bleu des doublons de la migration : traité par José
            if c.get('r')[0] in 'ABCDE' and c.get('s') and fill_of(c) == BLEU_DOUBLON:
                set_fill(c, None)
                report.setdefault('bleu_retire', set()).add('%s ligne %d' % (sh, r))
        roses = {fill_of(c) for c in row if c.get('r')[0] in 'ABCDE' and c.get('s')} & ROSES
        if roses:
            it = info.get(r, {})
            if ROSE_DOUTE in roses and it.get('kind') == 'obs' and it['num'] not in {m[1] for m in report['modifiees'] if m[0] == sh}:
                report['alertes'].append('%s %s : point [?] laissé tel quel, considéré comme validé' % (sh, it['num']))
            for c in row:
                if c.get('r')[0] in 'ABCDE' and c.get('s') and fill_of(c) in ROSES:
                    set_fill(c, None)
        cJ = row.find(N + 'c[@r="%s%d"]' % (ROUT_COL, r))
        if cJ is not None:
            clear(cJ)
    finir_onglet(sh)

vider_non_route()
maj_referentiel(next_num)
terminer(OUT)
report['bleu_retire'] = sorted(report.get('bleu_retire', []))
json.dump(report, open(REPORT, 'w', encoding='utf-8'), ensure_ascii=False, indent=1)
print('OK', OUT, {k: (len(v) if isinstance(v, list) else v) for k, v in report.items()})
