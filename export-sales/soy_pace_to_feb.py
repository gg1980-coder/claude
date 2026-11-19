"""Soybeans: what China and rest of world need to book from week N to week 26 (end-Feb) to stay on pace for USDA (China 25 Mt),
with Unknown split 60/40 (ours) or all Unknown in rest of world (Mike's way).
Usage: python3 soy_pace_to_feb.py <esr_soy.csv> [usda_mbu=1720] [end_week=26]"""
import sys, numpy as np, pandas as pd
src = sys.argv[1]; USDA = float(sys.argv[2]) if len(sys.argv) > 2 else 1720; W1 = int(sys.argv[3]) if len(sys.argv) > 3 else 26
K = 0.0272155; CN = 25.0
d = pd.read_csv(src); d['dt'] = pd.to_datetime(d.WeekendingDate); d['Y'] = d.ReportingMarketingYearName.str.extract(r'(\d{4})/').astype(int)
d = d[(d.dt >= pd.to_datetime(d.Y.astype(str) + '-08-31')) & (d.dt <= pd.to_datetime((d.Y + 1).astype(str) + '-09-07'))]
d['g'] = 'RoW'; d.loc[d.CountryName.str.strip() == 'CHINA, PEOPLES REPUBLIC OF', 'g'] = 'China'; d.loc[d.CountryName.str.contains('UNKNOWN'), 'g'] = 'Unk'
CY = int(d.loc[d.dt.idxmax(), 'Y']); WK = int(d[d.Y == CY].WeekNumber.max()); LAST = d[d.Y == CY].dt.max()
tc = d.groupby(['Y', 'WeekNumber', 'g']).TotalCommitment.sum().unstack('g').fillna(0) / 1e6
fin = {y: d[(d.Y == y) & (d.WeekNumber == d[d.Y == y].WeekNumber.max())].groupby('g').AccumulatedExports.sum() / 1e6 for y in range(CY - 10, CY)}
def series(y, w, kind):
    t = tc.loc[(y, w)]
    return {'cn60': t['China'] + .6 * t['Unk'], 'cn0': t['China'], 'row40': t['RoW'], 'row100': t['RoW'] + t['Unk']}[kind]
def final(y, kind):
    f = fin[y]; return {'cn60': f.get('China', 0), 'cn0': f.get('China', 0), 'row40': f.get('RoW', 0), 'row100': f.get('RoW', 0)}[kind]
YRS = int(sys.argv[4]) if len(sys.argv) > 4 else 10            # history window (seasons) for the seasonal pattern and 'normal' pace
CN_YRS = list(range(CY - YRS, CY))
ROW_YRS = list(range(CY - YRS, CY))
def shape(kind, yrs):   # pooled share of the remaining season (week N -> final) booked by week W1
    num = sum(min(max(series(y, W1, kind) - series(y, WK, kind), 0), max(final(y, kind) - series(y, WK, kind), 1e-9)) for y in yrs)
    den = sum(max(final(y, kind) - series(y, WK, kind), 1e-9) for y in yrs)
    return num / den
def hist_w(kind, yrs): return np.mean([series(y, W1, kind) - series(y, WK, kind) for y in yrs])
NW = W1 - WK; bar = USDA * K
rows = []
for lab, kind, target_fn, yrs in [('China, Unknown 60% to China (ours)', 'cn60', lambda: CN, CN_YRS), ('China, named only (Mike)', 'cn0', lambda: CN, CN_YRS),
                                  ('Rest of world, 40% of Unknown (ours)', 'row40', lambda: bar - CN, ROW_YRS), ('Rest of world incl. all Unknown (Mike)', 'row100', lambda: bar - CN, ROW_YRS)]:
    now = series(CY, WK, kind); tgt = target_fn()
    if kind == 'row40': tgt = bar - CN - .6 * tc.loc[(CY, WK)]['Unk'] * 0  # rest of world target = USDA - 25 (China incl. its Unknown share)
    rem = tgt - now; sh = shape(kind, yrs); need = rem * sh
    l4 = (series(CY, WK, kind) - series(CY, WK - 4, kind)) * 1000 / 4
    rows.append({'Group': lab, 'Full-year target (Mt)': tgt, 'Booked wk %d (Mt)' % WK: now, 'Still to book, full season (Mt)': rem,
                 'Normal share of remaining done by wk %d' % W1: sh, 'Needed by end-Feb (Mt)': need, 'Needed per week (kt)': need * 1000 / NW,
                 'Normal wk %d-%d, %d-yr avg (kt/wk)' % (WK + 1, W1, YRS): hist_w(kind, yrs) * 1000 / NW, 'Last 4 weeks (kt/wk)': l4,
                 'Booked by end-Feb on pace (Mt)': now + need})
R = pd.DataFrame(rows)
pd.set_option('display.width', 250)
print(f'Data through w/e {LAST:%d %b %Y} (week {WK}); end week {W1} = w/e {(LAST + pd.Timedelta(weeks=W1 - WK)):%d %b %Y}; {NW} weeks. USDA {USDA:.0f} = {bar:.2f} Mt')
print(R.round(2).to_string(index=False))
print('China pattern years', CN_YRS, ' RoW pattern years', ROW_YRS)
