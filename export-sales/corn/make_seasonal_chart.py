"""Seasonal pattern of US corn weekly sales (net new sales = change in total commitments, all destinations), and what
USDA 3,275 vs the 5-yr pace need week by week from here. 4-week centred averages for readability."""
import sys, os, numpy as np, pandas as pd
HERE = os.path.dirname(os.path.abspath(__file__)); ROOT = os.path.dirname(HERE); os.chdir(ROOT)
sys.argv = ['build_corn.py', 'data/esr_corn_2016-2027.csv', os.path.join(HERE, '_tmp.xlsx'), '3275']
exec(open('build_corn.py').read()); os.remove(os.path.join(HERE, '_tmp.xlsx'))
import matplotlib; matplotlib.use('Agg'); import matplotlib.pyplot as plt
BLK, DGR, LGR, GRID = '#000000', '#555555', '#c8c8c8', '#e6e6e6'
plt.rcParams.update({'font.family': 'Carlito', 'font.size': 11, 'axes.edgecolor': DGR, 'figure.facecolor': 'white', 'axes.facecolor': 'white'})
W = list(range(2, 53))
def ws(y): return pd.Series({w: (tc.get((y, w), np.nan) - tc.get((y, w - 1), np.nan)) * 1000 for w in W})
sm = lambda s: s.rolling(4, center=True, min_periods=2).mean()
hist = pd.DataFrame({y: sm(ws(y)) for y in PRIOR})
a5 = hist[B5].mean(axis=1); lo, hi = hist.min(axis=1), hist.max(axis=1)
ly = hist[LY]; cur = sm(ws(CY)).loc[:WK]
fw = list(range(WK + 1, 54))
raw5 = pd.Series({w: np.mean([(tc.get((y, w), np.nan) - tc.get((y, w - 1), np.nan)) * 1000 for y in B5]) for w in fw})
SC_NEED = (BAR - now) * 1000 / raw5.sum(); SC_BASE = (BASE - now) * 1000 / raw5.sum()
need = raw5 * SC_NEED; base = raw5 * SC_BASE
rawly = pd.Series({w: (tc.get((LY, w), np.nan) - tc.get((LY, w - 1), np.nan)) * 1000 for w in fw}); SC_LY = rawly.sum() / raw5.sum()
need_s = sm(need.loc[:52])
fig, a = plt.subplots(figsize=(9, 4.9))
a.fill_between(W, lo, hi, color='#ededed', lw=0, label='Range, last 10 seasons')
a.plot(W, ly, color=DGR, lw=1.5, label=f'{lab(LY)} (record year)')
a.plot(W, a5, color=DGR, lw=1.4, ls=(0, (5, 3)), label='5-yr average')
a.plot(need_s.index, need_s.values, color=BLK, lw=1.9, ls=(0, (1.5, 2)), label=f'Needed for USDA {USDA:,.0f} (5-yr pattern × {SC_NEED:.2f})')
a.plot(cur.index, cur.values, color=BLK, lw=2.6, label=f'{CUR} (this season, to 5 Nov)')
a.axvline(WK, color=DGR, lw=0.8); a.text(WK + 0.4, a.get_ylim()[1] * 0.02 + 50, 'today', fontsize=9, color=DGR)
mt = {}
for w in range(2, 53):
    dt_ = LAST + pd.Timedelta(weeks=w - WK)
    if dt_.day <= 7: mt[w] = dt_.strftime('%b')
a.set_xticks(list(mt.keys()), list(mt.values())); a.set_xlim(2, 52); a.set_ylim(0, None)
a.set_ylabel('Weekly sales (kt, 4-week average)'); a.legend(frameon=False, loc='upper right', fontsize=9.5, ncol=2)
a.spines[['top', 'right']].set_visible(False); a.grid(axis='y', color=GRID, lw=0.8)
a.set_title(f'US corn: seasonal pattern of weekly sales, and what USDA {USDA:,.0f} needs from here\nUSDA needs {SC_NEED - 1:.0%} more than a normal week, every week; last season (the record) ran {SC_LY - 1:.0%} above normal', loc='left', fontsize=12)
fig.text(0.01, 0.01, 'Source: USDA FAS weekly export sales, net change in total commitments, all destinations. Month labels at the first week ending in that month.', fontsize=8.5, color=DGR)
fig.tight_layout(rect=(0, 0.04, 1, 1)); fig.savefig(os.path.join(HERE, 'corn_seasonal_weekly_sales.png'), dpi=200); plt.close(fig)
# monthly table: by month of week-ending date
dates = {w: LAST + pd.Timedelta(weeks=w - WK) for w in fw}
mon = pd.DataFrame({'m': [dates[w].strftime('%b') for w in fw], 'need': need.values, 'base': base.values, 'avg5': raw5.values, 'ly': rawly.values}, index=fw)
g = mon.groupby('m', sort=False).agg(weeks=('need', 'size'), need=('need', 'sum'), base=('base', 'sum'), avg5=('avg5', 'sum'), ly=('ly', 'sum'))
for k_ in ('need', 'avg5', 'ly'): g[k_ + '_wk'] = g[k_] / g.weeks
print(g.round(0).to_string())
print('scales need', round(SC_NEED, 3), 'base', round(SC_BASE, 3), 'LY', round(SC_LY, 3), 'sum5', raw5.sum(), 'sumLY', rawly.sum())
