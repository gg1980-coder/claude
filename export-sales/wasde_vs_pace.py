"""Does the week-10 sales pace predict USDA's later export revisions?
Usage: python3 wasde_vs_pace.py <us_exports_by_report.csv> <esr_corn.csv> <esr_soy.csv>
For each season with a November WASDE: USDA Nov forecast, USDA final (latest value), and the week-10 pace estimate
(leave-one-out 'sold so far + usual still to come', all destinations, ESR -> USDA basis), plus the monthly revision steps."""
import sys, numpy as np, pandas as pd
w, ec, es = sys.argv[1:4]
x = pd.read_csv(w, parse_dates=['ReleaseDate'])
CONV = {'Corn': (0.025401, 0.967, ec), 'Soybeans': (0.0272155, 0.977, es)}
def esr(p):
    d = pd.read_csv(p); d['Y'] = d.ReportingMarketingYearName.str.extract(r'(\d{4})/').astype(int); d['dt'] = pd.to_datetime(d.WeekendingDate)
    d = d[(d.dt >= pd.to_datetime(d.Y.astype(str) + '-08-31')) & (d.dt <= pd.to_datetime((d.Y + 1).astype(str) + '-09-07'))]
    tc = d.groupby(['Y', 'WeekNumber']).TotalCommitment.sum() / 1e6
    fin = {y: d[(d.Y == y) & (d.WeekNumber == d[d.Y == y].WeekNumber.max())].AccumulatedExports.sum() / 1e6 for y in d.Y.unique() if y < d.Y.max()}
    return tc, fin
rows = []
for c, (K, R, p) in CONV.items():
    tc, fin = esr(p); P = sorted(fin)
    s = x[x.Commodity == c]
    for my, g in s.groupby('MarketYear'):
        y = int(my[:4]); g = g.sort_values('ReleaseDate')
        nov = g[(g.ReleaseDate.dt.month == 11) & (g.ReleaseDate.dt.year == y)]
        if nov.empty: continue
        final = g.Value.iloc[-1]; final_date = g.ReleaseDate.iloc[-1]
        o = [q for q in P if q != y]
        pace = (tc[(y, 10)] + np.mean([fin[q] - tc[(q, 10)] for q in o])) / R / K
        after = g[g.ReleaseDate > nov.ReleaseDate.iloc[0]]
        steps = after.assign(chg=after.Value.diff().fillna(after.Value.iloc[0] - nov.Value.iloc[0]) if len(after) else None)
        steps = ', '.join(f"{r.ReleaseDate:%b%y} {r.chg:+.0f}" for r in steps.itertuples() if abs(r.chg) > 0.5) if len(after) else ''
        done = final_date.year >= y + 2 or (final_date.year == y + 1 and final_date.month >= 10)
        rows.append(dict(commodity=c, season=my, usda_nov=nov.Value.iloc[0], pace_wk10=round(pace), gap=round(pace - nov.Value.iloc[0]),
                         usda_latest=final, nov_to_latest=final - nov.Value.iloc[0], complete=done, steps=steps))
r = pd.DataFrame(rows)
r['same_direction'] = np.where(r.complete & (r.nov_to_latest.abs() > 0), np.sign(r.gap) == np.sign(r.nov_to_latest), None)
pd.set_option('display.width', 250); pd.set_option('display.max_colwidth', 120)
print(r.to_string(index=False))
