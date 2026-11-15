"""USDA (WASDE) US export revisions after November, 2010/11-2025/26, and whether the week-10 sales pace predicted them.
Usage: python3 wasde_revisions.py <wasde_us_exports_by_report.csv> <esr_corn.csv> <esr_soy.csv> [out.xlsx]"""
import sys, numpy as np, pandas as pd
from scipy import stats
w, ec, es = sys.argv[1:4]; OUT = sys.argv[4] if len(sys.argv) > 4 else None
x = pd.read_csv(w, parse_dates=['ReleaseDate'])
CONV = {'Corn': (0.025401, 0.967, ec), 'Soybeans': (0.0272155, 0.977, es)}
CUR = 2026
def esr(p):
    d = pd.read_csv(p); d['Y'] = d.ReportingMarketingYearName.str.extract(r'(\d{4})/').astype(int); d['dt'] = pd.to_datetime(d.WeekendingDate)
    d = d[(d.dt >= pd.to_datetime(d.Y.astype(str) + '-08-31')) & (d.dt <= pd.to_datetime((d.Y + 1).astype(str) + '-09-07'))]
    tc = d.groupby(['Y', 'WeekNumber']).TotalCommitment.sum() / 1e6
    fin = {y: d[(d.Y == y) & (d.WeekNumber == d[d.Y == y].WeekNumber.max())].AccumulatedExports.sum() / 1e6 for y in d.Y.unique() if y < CUR}
    return tc, fin
season_rows, step_rows = [], []
for c, (K, R, p) in CONV.items():
    tc, fin = esr(p); P = sorted(fin)
    for my, g in x[x.Commodity == c].groupby('MarketYear'):
        y = int(my[:4])
        if y < 2010: continue
        g = g.sort_values('ReleaseDate').reset_index(drop=True)
        ref = g[(g.ReleaseDate.dt.month == 11) & (g.ReleaseDate.dt.year == y)]; rl = 'Nov'
        if ref.empty: ref = g[(g.ReleaseDate.dt.month == 12) & (g.ReleaseDate.dt.year == y)]; rl = 'Dec (proxy)'
        if ref.empty: continue
        sep = g[(g.ReleaseDate.dt.month == 9) & (g.ReleaseDate.dt.year == y)]
        nov_v, nov_d = ref.Value.iloc[0], ref.ReleaseDate.iloc[0]
        final = g.Value.iloc[-1]; complete = y < CUR
        pace = gap = None
        if y in P or y == CUR:
            if (y, 10) in tc.index:
                o = [q for q in P if q != y]
                pace = (tc[(y, 10)] + np.mean([fin[q] - tc[(q, 10)] for q in o])) / R / K; gap = pace - nov_v
        season_rows.append(dict(commodity=c, season=my, ref=rl, usda_sep=sep.Value.iloc[0] if len(sep) else np.nan, usda_nov=nov_v,
                                final=final if complete else np.nan, rev=final - nov_v if complete else np.nan,
                                pace_wk10=pace, gap=gap, complete=complete))
        if complete:
            a = g[g.ReleaseDate >= nov_d].copy(); a['chg'] = a.Value.diff()
            for r in a.iloc[1:].itertuples():
                k = (r.ReleaseDate.year - nov_d.year) * 12 + r.ReleaseDate.month - nov_d.month   # months after November
                if 1 <= k <= 12: step_rows.append(dict(commodity=c, season=my, k=k, month=r.ReleaseDate.strftime('%b'), chg=r.chg))
S = pd.DataFrame(season_rows); T = pd.DataFrame(step_rows)
pd.set_option('display.width', 250)
for c in CONV:
    s = S[(S.commodity == c) & S.complete]
    print(f'\n===== {c}: Nov -> final, {len(s)} seasons {s.season.iloc[0]}-{s.season.iloc[-1]}')
    print(s[['season', 'ref', 'usda_sep', 'usda_nov', 'final', 'rev', 'pace_wk10', 'gap']].round(0).to_string(index=False))
    print(f'  final above Nov: {(s.rev > 0).sum()}  below: {(s.rev < 0).sum()}  mean |rev| {s.rev.abs().mean():.0f}  median |rev| {s.rev.abs().median():.0f}  max |rev| {s.rev.abs().max():.0f}')
    mom = s.dropna(subset=['usda_sep']); mom = mom[(mom.usda_nov - mom.usda_sep).abs() > 0]
    print(f'  USDA momentum (Sep->Nov sign = Nov->final sign): {(np.sign(mom.usda_nov - mom.usda_sep) == np.sign(mom.rev)).sum()} of {len(mom)}')
    t = T[T.commodity == c]
    tm = t.groupby('k').agg(month=('month', 'first'), mean_abs=('chg', lambda v: v.abs().mean()), share_changed=('chg', lambda v: (v.abs() > 0.5).mean()))
    tm['share_of_total_abs'] = t.groupby('k').chg.apply(lambda v: v.abs().sum()) / t.chg.abs().sum()
    print('  timing after Nov:'); print(tm.round(2).to_string())
    steps = t[t.chg.abs() > 0.5].chg.abs(); print(f'  step size when changed: median {steps.median():.0f}, 75th pct {steps.quantile(.75):.0f}, max {steps.max():.0f}')
    pz = s.dropna(subset=['gap'])
    hit = (np.sign(pz.gap) == np.sign(pz.rev)).sum()
    lr = stats.linregress(pz.gap, pz.rev); resid = pz.rev - (lr.intercept + lr.slope * pz.gap)
    print(f'  PACE TEST {len(pz)} seasons: direction hit {hit}/{len(pz)}; rev = {lr.intercept:.0f} + {lr.slope:.2f} x gap, R2 {lr.rvalue**2:.2f}, p {lr.pvalue:.3f}, resid sd {resid.std(ddof=2):.0f}')
    cur = S[(S.commodity == c) & (S.season == f'{CUR}/{str(CUR+1)[-2:]}')].iloc[0]
    pred = lr.intercept + lr.slope * cur.gap; se = resid.std(ddof=2) * np.sqrt(1 + 1 / len(pz) + (cur.gap - pz.gap.mean()) ** 2 / ((pz.gap - pz.gap.mean()) ** 2).sum())
    print(f'  2026/27: Nov {cur.usda_nov:.0f}, pace {cur.pace_wk10:.0f}, gap {cur.gap:.0f} -> regression revision {pred:+.0f} (se {se:.0f}); P(cut) {stats.t.cdf(-pred / se, len(pz) - 2):.0%}')
if OUT:
    with pd.ExcelWriter(OUT) as xw:
        S.round(1).to_excel(xw, 'Seasons', index=False); T.to_excel(xw, 'Monthly steps', index=False)
