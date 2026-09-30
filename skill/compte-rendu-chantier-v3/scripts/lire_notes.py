# -*- coding: utf-8 -*-
"""Liste les notes de cellule posées par José sur le CR, avec le contexte utile pour rédiger OPERATIONS.json.

Usage : python lire_notes.py CR_ANNOTE.xlsx [NOTES.json]

Pour chaque note : onglet, cellule, et ce qu'elle désigne
- une observation (N°, texte affiché, ABORDÉ LE / POUR LE / FAIT LE) -> opération « maj » ;
- une ligne ÉTUDES / TRAVAUX ou l'en-tête d'un tableau -> opération(s) « nouvelle » (un paragraphe = une ligne) ;
- autre chose -> à interpréter, et à signaler à José.
Rien n'est modifié dans le classeur.
"""
import datetime
import json
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from cr_xml import Book  # noqa: E402
import cr_classeur as K  # noqa: E402


def val(c):
    n = K.number_of(c)
    if n is not None and n > 20000:
        return K.from_serial(n).strftime('%d/%m/%Y')
    return K.text_of(c).strip()


def auteurs(sh):
    cp = K.b.rel_target(K.b.sheets[sh], 'comments')
    return [a.text or '' for a in K.b.xml(cp).iter(K.N + 'author')] if cp else []


def sans_auteur(txt, noms):
    """Excel écrit « Auteur: » en tête de chaque note : retiré."""
    for n in noms:
        if n and txt.lstrip().startswith(n + ':'):
            return txt.lstrip()[len(n) + 1:].strip()
    return txt.strip()


def main(src, out=None):
    K.init(Book(src), datetime.datetime.now())
    res = []
    for sh in K.obs_sheets():
        notes = K.b.notes(sh)
        if not notes:
            continue
        noms = auteurs(sh)
        notes = {ref: sans_auteur(t, noms) for ref, t in notes.items()}
        rows, tabs, info = K.model(sh)
        for ref, txt in sorted(notes.items(), key=lambda x: (int(''.join(ch for ch in x[0] if ch.isdigit())), x[0])):
            r = int(''.join(ch for ch in ref if ch.isdigit()))
            it = info.get(r, {'kind': 'hors tableau'})
            if it['kind'] == 'hors tableau':            # titre de section juste au-dessus d'un tableau
                t_sous = next((t for t in tabs if 0 < t['r0'] - r <= 3), None)
                if t_sous:
                    it = info[t_sous['r0']]
            t = next((t for t in tabs if t['name'] == it.get('table')), None)
            d = {'note': '%s!%s' % (sh, ref), 'texte_note': txt.strip(), 'ligne': r, 'type': it['kind'],
                 'code': t['code'] if t else None, 'section': it.get('section')}
            row = rows.get(r)
            if it['kind'] == 'obs' and row is not None:
                d.update(num=it['num'], masquee=it.get('hidden'),
                         observations=K.text_of(K.cell(row, 'B'))[:400],
                         aborde_le=val(K.cell(row, 'C')), pour_le=val(K.cell(row, 'D')), fait_le=val(K.cell(row, 'E')))
                d['operation'] = 'maj %s' % it['num']
            elif it['kind'] == 'section':
                d['operation'] = 'nouvelle(s) : %s, %s' % (d['code'], it['section'])
            elif it['kind'] == 'header':
                d['operation'] = 'nouvelle(s) : %s (tableau %s)' % (d['code'], it['table'])
            elif it['kind'] == 'empty' and t:
                d['operation'] = 'nouvelle(s) : %s%s (ligne vide du tableau)' % (d['code'], ', ' + it['section'] if it.get('section') else '')
            else:
                d['operation'] = 'à interpréter (hors tableau)'
            res.append(d)
    for d in res:
        print('%-14s %-45s %s' % (d['note'], d['operation'], d['texte_note'].replace('\n', ' / ')[:120]))
    print('%d note(s)' % len(res))
    if out:
        json.dump(res, open(out, 'w', encoding='utf-8'), ensure_ascii=False, indent=1)


if __name__ == '__main__':
    main(*sys.argv[1:3])
