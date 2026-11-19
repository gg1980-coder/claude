"""Regional pace vs USDA for soybeans (China capped at 25 Mt incl. 60% of Unknown) or corn (no cap; Unknown not assigned).
Needed from week N to week 26 and full season; 10-yr normal; last 4 weeks; full year implied by USDA vs 10-season history.
Usage: python3 region_pace.py soy|corn <esr.csv> [usda] [out.xlsx|-] [cut_week=26]"""
import sys, numpy as np, pandas as pd
crop, src = sys.argv[1:3]; MIKE = crop == 'soy_mike'; crop = 'soy' if MIKE else crop; USDA = float(sys.argv[3]) if len(sys.argv) > 3 else (1720 if crop == 'soy' else 3275); OUT = sys.argv[4] if len(sys.argv) > 4 and sys.argv[4] != '-' else None
W1 = int(sys.argv[5]) if len(sys.argv) > 5 else 26; WE, YRS = 53, 10
d = pd.read_csv(src); d['C'] = d.CountryName.str.strip(); d['dt'] = pd.to_datetime(d.WeekendingDate)
d['Y'] = d.ReportingMarketingYearName.str.extract(r'(\d{4})/').astype(int)
d = d[(d.dt >= pd.to_datetime(d.Y.astype(str) + '-08-31')) & (d.dt <= pd.to_datetime((d.Y + 1).astype(str) + '-09-07'))]
CY = int(d.loc[d.dt.idxmax(), 'Y']); WK = int(d[d.Y == CY].WeekNumber.max()); LAST = d[d.Y == CY].dt.max(); P = list(range(CY - YRS, CY)); LY = CY - 1
M = {}
if crop == 'soy':
    K, RATIO = 0.0272155, 1.0
    groups = [(['CHINA, PEOPLES REPUBLIC OF'], 'China'), (['MEXICO'], 'Mexico'), (['JAPAN', 'TAIWAN', 'KOREA, REPUBLIC OF', 'HONG KONG'], 'East Asia ex-China'),
              (['INDONESIA', 'VIETNAM', 'MALAYSIA', 'THAILAND', 'PHILIPPINES', 'BURMA', 'CAMBODIA', 'LAOS', 'SINGAPORE'], 'Southeast Asia'),
              (['BANGLADESH', 'PAKISTAN', 'NEPAL', 'SRI LANKA', 'INDIA'], 'South Asia'),
              (['EGYPT', 'ALGERIA', 'TURKEY', 'TUNISIA', 'MOROCCO', 'IRAN', 'IRAQ', 'ISRAEL', 'JORDAN', 'LEBANON', 'LIBYA', 'SAUDI ARABIA', 'QATAR', 'UNITED ARAB EMIRATES'], 'Middle East & N. Africa'),
              (['SPAIN', 'NETHERLANDS', 'PORTUGAL', 'ITALY', 'GERMANY', 'UNITED KINGDOM', 'GREECE', 'BELGIUM-LUXEMBOURG', 'FRANCE', 'IRELAND', 'NORWAY', 'ROMANIA', 'SWITZERLAND', 'UKRAINE', 'GEORGIA, REPUBLIC OF'], 'Europe'),
              (['UNKNOWN'], 'Unknown')]
    OTHER = 'Other (Americas, Africa)'
else:
    K, RATIO = 0.025401, 0.967
    groups = [(['MEXICO'], 'Mexico'), (['JAPAN'], 'Japan'), (['KOREA, REPUBLIC OF'], 'South Korea'), (['COLOMBIA'], 'Colombia'),
              (['TAIWAN', 'HONG KONG', 'VIETNAM', 'INDONESIA', 'PHILIPPINES', 'MALAYSIA', 'THAILAND', 'BANGLADESH', 'SINGAPORE', 'BURMA', 'CAMBODIA', 'SRI LANKA', 'PAKISTAN', 'INDIA', 'NEPAL'], 'Other Asia (Taiwan, SE Asia)'),
              (['GUATEMALA', 'HONDURAS', 'COSTA RICA', 'DOMINICAN REPUBLIC', 'EL SALVADOR', 'PANAMA', 'NICARAGUA', 'JAMAICA', 'CUBA', 'HAITI', 'BARBADOS', 'LEEWARD-WINDWARD ISLANDS', 'TRINID', 'BELIZE', 'BAHAMAS'], 'Central America & Caribbean'),
              (['PERU', 'VENEZUELA', 'CHILE', 'ECUADOR', 'GUYANA', 'SURINAM'], 'Other South America'),
              (['SPAIN', 'PORTUGAL', 'NETHERLANDS', 'IRELAND', 'UNITED KINGDOM', 'ITALY', 'GERMANY', 'FRANCE', 'BELGIUM-LUXEMBOURG', 'GREECE'], 'Europe'),
              (['SAUDI ARABIA', 'MOROCCO', 'EGYPT', 'ISRAEL', 'TUNISIA', 'ALGERIA', 'JORDAN', 'IRAQ', 'UNITED ARAB EMIRATES', 'OMAN', 'LIBYA', 'TURKEY', 'LEBANON', 'KUWAIT', 'QATAR', 'YEMEN'], 'Middle East & N. Africa'),
              (['CHINA, PEOPLES REPUBLIC OF'], 'China'), (['UNKNOWN'], 'Unknown')]
    OTHER = 'Canada, Africa & other'
for lst, nm in groups:
    for c_ in lst: M[c_] = nm
d['R'] = d.C.map(M).fillna(OTHER)
REGS = [g[1] for g in groups if g[1] not in ('Unknown',)] + [OTHER] + (['Unknown'] if MIKE else [])
cum = d.pivot_table(index=['R', 'WeekNumber'], columns='Y', values='TotalCommitment', aggfunc='sum').fillna(0) / 1e6
fin = d[d.Y < CY].sort_values('dt').groupby(['Y', 'C']).tail(1).groupby(['R', 'Y']).AccumulatedExports.sum().unstack().fillna(0) / 1e6
def c(r, y, w):
    try: return float(cum.loc[(r, w), y])
    except KeyError: return 0.0
def f(r, y): return float(fin.loc[r, y]) if (r in fin.index and y in fin.columns) else 0.0
BAR = USDA * K * RATIO; unk = c('Unknown', CY, WK)
if crop == 'soy':   # China fixed at 25 incl. 60% of Unknown; rest of world shares the rest
    UC = 0.0 if MIKE else 0.6                          # Mike: China = named only, all Unknown in rest of world
    china_now = c('China', CY, WK) + UC * unk; china_need = 25.0 - china_now
    named = [r for r in REGS if r != 'China']; named_now = sum(c(r, CY, WK) for r in named); pool = BAR - 25.0 - named_now
else:
    named = REGS; named_now = sum(c(r, CY, WK) for r in named); pool = BAR - named_now
rem = {r: np.mean([f(r, y) - c(r, y, WK) for y in P]) for r in named}; scale = pool / sum(rem.values())
rows = []
def row(r, now, need, remhist, a_hist, finals, l4, ly_now):
    remy = [max(x, 1e-9) for x in remhist]
    sh = sum(min(max(a, 0), rr) for a, rr in zip(a_hist, remy)) / sum(remy)
    na = need * sh; norm_a = np.mean(a_hist); norm_full = np.mean(remhist)
    fy = now + need; nb = need - na; norm_b = norm_full - norm_a
    return {'Region': r, 'Booked (Mt)': now, 'vs last season (Mt)': now - ly_now, 'Needed to cut (kt/wk)': na * 1000 / (W1 - WK),
            'Normal to cut, 10-yr (kt/wk)': norm_a * 1000 / (W1 - WK), 'Last 4 weeks (kt/wk)': l4,
            'Needed after cut (kt/wk)': nb * 1000 / (WE - W1), 'Normal after cut, 10-yr (kt/wk)': norm_b * 1000 / (WE - W1),
            'Needed full season (kt/wk)': need * 1000 / (WE - WK), 'Still to book, full season (Mt)': need,
            'Normally booked from here (Mt)': norm_full, 'Needed % of normal': need / norm_full if norm_full > 0.01 else np.nan,
            'Full year implied by USDA (Mt)': fy, '10-yr avg full year (Mt)': np.mean(finals), 'Best of 10 seasons (Mt)': max(finals),
            'Rank (1 = highest)': sum(x > fy for x in finals) + 1,
            'Read (last 4 wks vs needed to cut)': 'ahead' if l4 > na * 1000 / (W1 - WK) * 1.1 else ('behind' if l4 < na * 1000 / (W1 - WK) * 0.9 else 'on pace')}
for r in named:
    rows.append(row(r, c(r, CY, WK), rem[r] * scale, [f(r, y) - c(r, y, WK) for y in P], [c(r, y, W1) - c(r, y, WK) for y in P],
                    [f(r, y) for y in P], (c(r, CY, WK) - c(r, CY, WK - 4)) * 1000 / 4, c(r, LY, WK)))
if crop == 'soy':
    cs = lambda y, w: c('China', y, w) + UC * c('Unknown', y, w)
    rows.insert(0, row('China (named only), capped at 25 Mt' if MIKE else 'China (+60% of Unknown), capped at 25 Mt', china_now, china_need, [f('China', y) - cs(y, WK) for y in P], [cs(y, W1) - cs(y, WK) for y in P],
                       [f('China', y) for y in P], (cs(CY, WK) - cs(CY, WK - 4)) * 1000 / 4, cs(LY, WK)))
T = pd.DataFrame(rows)
tot_named = T[~T.Region.str.startswith('China (')]
tot = {'Region': ('Rest of world incl. Unknown' if MIKE else 'Rest of world') if crop == 'soy' else 'TOTAL (named destinations)'}
for k in ('Booked (Mt)', 'vs last season (Mt)', 'Needed to cut (kt/wk)', 'Normal to cut, 10-yr (kt/wk)', 'Last 4 weeks (kt/wk)', 'Needed after cut (kt/wk)', 'Normal after cut, 10-yr (kt/wk)', 'Needed full season (kt/wk)', 'Still to book, full season (Mt)',
          'Normally booked from here (Mt)', 'Full year implied by USDA (Mt)', '10-yr avg full year (Mt)'):
    tot[k] = tot_named[k].sum()
tot['Needed % of normal'] = tot['Still to book, full season (Mt)'] / tot['Normally booked from here (Mt)']
T = pd.concat([T, pd.DataFrame([tot])], ignore_index=True)
print(f'{crop.upper()} w/e {LAST:%d %b %Y} wk {WK}; cut wk {W1} = w/e {(LAST + pd.Timedelta(weeks=W1 - WK)):%d %b %Y} ({W1-WK} wks), then {WE-W1} wks; USDA {USDA:.0f} -> {BAR:.2f} Mt; Unknown {unk:.2f} Mt; pool {pool:.2f}; scale {scale:.2f}')
pd.set_option('display.width', 300); print(T.round(2).to_string(index=False))
if OUT:
    T.round(3).to_excel(OUT, index=False, sheet_name=crop)
