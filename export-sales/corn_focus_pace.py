"""Corn: what Mexico, Japan, South Korea and Europe need to book from week N to end-Feb (wk 26) and to end-Aug for USDA, vs their 10-yr pace.
USDA bar on weekly-sales basis (x ESR/USDA ratio, as in the corn workbook). Unknown is not assigned: the requirement for named destinations is
the USDA bar minus named bookings, shared by each region's usual (10-yr) bookings from here (which already include Unknown later named).
Usage: python3 corn_focus_pace.py <esr_corn.csv> [usda_mbu=3275] [years=10]"""
import sys, numpy as np, pandas as pd
src = sys.argv[1]; USDA = float(sys.argv[2]) if len(sys.argv) > 2 else 3275; YRS = int(sys.argv[3]) if len(sys.argv) > 3 else 10
K, RATIO = 0.025401, 0.967; W1, WE = 26, 53
d = pd.read_csv(src); d['C'] = d.CountryName.str.strip(); d['dt'] = pd.to_datetime(d.WeekendingDate)
d['Y'] = d.ReportingMarketingYearName.str.extract(r'(\d{4})/').astype(int)
d = d[(d.dt >= pd.to_datetime(d.Y.astype(str) + '-08-31')) & (d.dt <= pd.to_datetime((d.Y + 1).astype(str) + '-09-07'))]
CY = int(d.loc[d.dt.idxmax(), 'Y']); WK = int(d[d.Y == CY].WeekNumber.max()); LAST = d[d.Y == CY].dt.max(); P = list(range(CY - YRS, CY)); LY = CY - 1
EU = ['SPAIN', 'PORTUGAL', 'NETHERLANDS', 'IRELAND', 'UNITED KINGDOM', 'ITALY', 'GERMANY', 'FRANCE', 'BELGIUM-LUXEMBOURG', 'GREECE']
REG = {'MEXICO': 'Mexico', 'JAPAN': 'Japan', 'KOREA, REPUBLIC OF': 'South Korea', 'UNKNOWN': 'Unknown', **{c: 'Europe' for c in EU}}
d['R'] = d.C.map(REG).fillna('All other destinations')
FOCUS = ['Mexico', 'Japan', 'South Korea', 'Europe']; NAMED = FOCUS + ['All other destinations']
cum = d.pivot_table(index=['R', 'WeekNumber'], columns='Y', values='TotalCommitment', aggfunc='sum').fillna(0) / 1e6
fin = d[d.Y < CY].sort_values('dt').groupby(['Y', 'C']).tail(1).groupby(['R', 'Y']).AccumulatedExports.sum().unstack().fillna(0) / 1e6
def c(r, y, w):
    try: return float(cum.loc[(r, w), y])
    except KeyError: return 0.0
def f(r, y): return float(fin.loc[r, y]) if (r in fin.index and y in fin.columns) else 0.0
BAR = USDA * K * RATIO
now = {r: c(r, CY, WK) for r in NAMED + ['Unknown']}; named_now = sum(now[r] for r in NAMED)
rem = {r: np.mean([f(r, y) - c(r, y, WK) for y in P]) for r in NAMED}
scale = (BAR - named_now) / sum(rem.values())
rows = []
for r in NAMED:
    remy = {y: max(f(r, y) - c(r, y, WK), 1e-9) for y in P}
    sh = sum(min(max(c(r, y, W1) - c(r, y, WK), 0), remy[y]) for y in P) / sum(remy.values())
    need = rem[r] * scale; na = need * sh; nb = need - na
    norm_a = np.mean([c(r, y, W1) - c(r, y, WK) for y in P]); norm_b = rem[r] - norm_a
    l4 = (c(r, CY, WK) - c(r, CY, WK - 4)) * 1000 / 4
    hist_full = [f(r, y) for y in P]
    rows.append(dict(region=r, booked=now[r], booked_ly=c(r, LY, WK), need_full=need, norm_full=rem[r], pct=need / rem[r],
                     need_a=na, need_a_wk=na * 1000 / (W1 - WK), norm_a_wk=norm_a * 1000 / (W1 - WK), need_b_wk=nb * 1000 / (WE - W1), norm_b_wk=norm_b * 1000 / (WE - W1),
                     need_full_wk=need * 1000 / (WE - WK), l4=l4, fy_usda=now[r] + need, fy_best=max(hist_full), fy_avg=np.mean(hist_full),
                     fy_rank=sum(h > now[r] + need for h in hist_full) + 1))
R = pd.DataFrame(rows); pd.set_option('display.width', 280)
print(f'Data w/e {LAST:%d %b %Y} (wk {WK}); USDA {USDA:.0f} mbu -> bar {BAR:.2f} Mt (ESR basis); named booked {named_now:.2f}; Unknown {now["Unknown"]:.2f}; '
      f'named still to book {BAR - named_now:.2f}; scale vs {YRS}-yr usual {scale:.2f}; total booked {named_now + now["Unknown"]:.2f}')
print(R.round(2).to_string(index=False))
