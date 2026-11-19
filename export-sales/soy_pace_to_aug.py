"""Soybeans: what China and rest of world need from week N to the end of the marketing year (end-August) for USDA (China 25 Mt),
split into week N+1-26 (to end-Feb) and week 27-end (Mar-Aug). Both Unknown treatments. Usage: python3 soy_pace_to_aug.py <esr_soy.csv> [usda=1720]"""
import sys, numpy as np, pandas as pd
src = sys.argv[1]; USDA = float(sys.argv[2]) if len(sys.argv) > 2 else 1720
exec(open('soy_pace_to_feb.py').read().split('NW = W1 - WK')[0].replace("src = sys.argv[1]; USDA = float(sys.argv[2]) if len(sys.argv) > 2 else 1720; W1 = int(sys.argv[3]) if len(sys.argv) > 3 else 26", "W1 = 26"))
bar = USDA * K
LASTW = int(d[d.Y == CY - 1].WeekNumber.max())   # weeks in a season (incl. stub week)
WE = 53                                           # last week of the marketing year (43 weeks from week 10, as in the analysts note)
n_a, n_b = W1 - WK, WE - W1
rows = []
for lab, kind, yrs in [('China, Unknown 60% to China', 'cn60', CN_YRS), ('China, named only', 'cn0', CN_YRS),
                       ('Rest of world, 40% of Unknown', 'row40', ROW_YRS), ('Rest of world, incl. all Unknown', 'row100', ROW_YRS)]:
    now = series(CY, WK, kind); tgt = CN if kind.startswith('cn') else bar - CN; rem = tgt - now
    sh = shape(kind, yrs); a = rem * sh; b = rem - a
    norm_a = np.mean([series(y, W1, kind) - series(y, WK, kind) for y in yrs]); norm_all = np.mean([final(y, kind) - series(y, WK, kind) for y in yrs]); norm_b = norm_all - norm_a
    l4 = (series(CY, WK, kind) - series(CY, WK - 4, kind)) * 1000 / 4
    rows.append(dict(group=lab, booked=now, rem=rem, need_a=a, need_a_wk=a * 1000 / n_a, norm_a_wk=norm_a * 1000 / n_a, need_b=b, need_b_wk=b * 1000 / n_b,
                     norm_b_wk=norm_b * 1000 / n_b, need_all_wk=rem * 1000 / (n_a + n_b), norm_all=norm_all, norm_all_wk=norm_all * 1000 / (n_a + n_b), pct=rem / norm_all, l4=l4))
R = pd.DataFrame(rows); pd.set_option('display.width', 260)
print(f'wk{WK} -> wk{W1}: {n_a} weeks (to {(LAST + pd.Timedelta(weeks=W1 - WK)):%d %b}); wk{W1+1} -> end Aug: {n_b} weeks; total {n_a + n_b} weeks')
print(R.round(2).to_string(index=False))
