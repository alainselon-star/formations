# Génère data.js du test de positionnement Power BI depuis la banque Excel
import openpyxl, json, sys
src = sys.argv[1]; out = sys.argv[2]
wb = openpyxl.load_workbook(src, read_only=True)
rows = [r for r in list(wb['Banque'].iter_rows(values_only=True))[1:] if r[0]]
Q = []
for r in rows:
    ans = [a for a in r[5:10] if a]
    assert ans[-1] == 'Je ne sais pas', r[0]
    Q.append(dict(id=r[0], d=r[1][:2], l=r[2][1:3], q=r[4], a=ans, c='ABCDE'.index(r[10].strip()), fb=r[11],
                  err=[x if x and not str(x).startswith('—') else '' for x in r[12:16]]))
T, draws, D = {}, {}, {}
for t in list(wb['Tiroirs'].iter_rows(values_only=True))[1:]:
    if not t[0] or not str(t[0]).startswith('L'): continue
    l = t[0][1:3]; T[l] = t[2]; draws[l] = t[3]
    D.setdefault(t[1][:2], {'name': t[1][3:], 'lessons': []})['lessons'].append(l)
open(out, 'w').write('window.POS=' + json.dumps({'Q': Q, 'draws': draws, 'titles': T, 'domains': D}, ensure_ascii=False) + ';')
print(len(Q), 'questions ;', sum(draws.values()), 'tirées ;', {k: (v['name'], v['lessons']) for k, v in D.items()})
