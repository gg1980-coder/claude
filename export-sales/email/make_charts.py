"""Charts for the desk email (week 9, 2026/27). Desk number 1,800 mln bu (China 25 Mt, Brazil competing harder than normal)."""
import numpy as np, pandas as pd, matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
from matplotlib.colors import LinearSegmentedColormap, TwoSlopeNorm
from matplotlib.patches import FancyBboxPatch
from scipy import stats
OUT = '/home/user/claude/export-sales/email/'
DATA = '/home/user/claude/export-sales/data/esr_soybeans_2016-2027.csv'
INK, INK2, GRID, SURF = '#0b0b0b', '#52514e', '#e4e3df', '#fcfcfb'
BLUE, RED, GREY, ORANGE, AQUA = '#2a78d6', '#e34948', '#c3c2b7', '#eb6834', '#1baf7a'
plt.rcParams.update({'font.family': 'DejaVu Sans', 'font.size': 10, 'axes.edgecolor': GRID, 'axes.labelcolor': INK2,
                     'xtick.color': INK2, 'ytick.color': INK2, 'figure.facecolor': SURF, 'axes.facecolor': SURF})
USDA, PACE, BASE, DESK = 1720, 1859, 1825, 1800
CONV = 0.0272155 * 0.97729741421324          # Mt (ESR basis) per mln bu
PER_MT = 1 / CONV
RISK = [0, -15, -29, -44, -57, -67, -77]    # extra Brazil supply 0..6 Mt -> mln bu (switching workbook)

# ---------- data for milestones: rest-of-world (named, ex China & Unknown) cumulative commitments ----------
d = pd.read_csv(DATA); d['dt'] = pd.to_datetime(d.WeekendingDate); d['Y'] = d.ReportingMarketingYearName.str.extract(r'Sep (\d{4})').astype(int)
d = d[(d.dt >= pd.to_datetime(d.Y.astype(str) + '-08-31')) & (d.dt <= pd.to_datetime((d.Y + 1).astype(str) + '-09-07'))]
cy = int(d.loc[d.dt.idxmax(), 'Y']); WK = int(d[d.Y == cy].WeekNumber.max()); LAST = d[d.Y == cy].dt.max()
rw = d[~d.CountryName.isin(['CHINA, PEOPLES REPUBLIC OF', 'UNKNOWN'])]
cum = rw.groupby(['Y', 'WeekNumber']).TotalCommitment.sum().unstack(0) / 1e6
fin = rw.sort_values('dt').groupby(['Y', 'CountryName']).tail(1).groupby('Y').AccumulatedExports.sum() / 1e6
B5 = [cy - 5 + i for i in range(5)]
now = cum.loc[WK, cy]
shape = {w: np.mean([(cum.loc[w, y] - cum.loc[WK, y]) / (fin[y] - cum.loc[WK, y]) for y in B5]) for w in range(WK, 53)}
def path(X): still = X * CONV - 25 - now; return {w: now + still * shape[w] for w in shape}
paths = {X: path(X) for X in (USDA, DESK, BASE)}
MIL = [13, 17, 22, 26, 35]
date = lambda w: (LAST + pd.Timedelta(weeks=w - WK)).strftime('%d %b')
pd.DataFrame({f'{X}': [paths[X][w] for w in MIL] for X in paths}, index=[f'wk{w} {date(w)}' for w in MIL]).round(1).to_csv(OUT + 'milestones.csv')

# 1) Simple bridge
steps = [('USDA\nNov WASDE', USDA, 'total'), ('China\nat 25 Mt', 92, 'd'), ('Rest of world\non pace', 47, 'd'), ('Pace\nestimate', PACE, 'total'),
         ('Prudence\n(median season)', -34, 'd'), ('Brazil competes\nharder', -25, 'd'), ('DESK\nNUMBER', DESK, 'total')]
fig, ax = plt.subplots(figsize=(10, 5)); run = 0
for i, (lab, v, kind) in enumerate(steps):
    if kind == 'total':
        col = INK2 if i == 0 else (BLUE if lab.startswith('DESK') else '#7fa9de')
        ax.bar(i, v - 1600, bottom=1600, color=col, width=0.6, zorder=3); run = v
        ax.text(i, v + 5, f'{v:,.0f}', ha='center', va='bottom', fontsize=12, fontweight='bold', color=INK)
    else:
        ax.bar(i, v, bottom=run, color=BLUE if v > 0 else RED, width=0.6, zorder=3, alpha=0.85)
        ax.text(i, max(run, run + v) + 5, f'{v:+,.0f}', ha='center', va='bottom', fontsize=11, color=INK); run += v
ax.axhline(USDA, color=INK2, lw=1, ls=(0, (4, 3)), zorder=2)
ax.set_xticks(range(len(steps)), [s[0] for s in steps]); ax.set_ylim(1600, 1900); ax.set_ylabel('US soybean exports 2026/27 (mln bu)')
ax.spines[['top', 'right']].set_visible(False); ax.grid(axis='y', color=GRID, lw=0.8, zorder=0)
ax.set_title('From USDA 1,720 to our desk number 1,800', loc='left', fontsize=13, color=INK)
fig.text(0.01, 0.01, f'Dashed line = USDA. Rest of world = Middle East/N. Africa +28, SE Asia +25, Europe +5, others +3, East Asia −8, South Asia −6. ESR data to {LAST:%d %b %Y} (wk {WK}).',
         fontsize=7.5, color=INK2)
fig.tight_layout(rect=(0, 0.03, 1, 1)); fig.savefig(OUT + '1_bridge.png', dpi=200); plt.close(fig)

# 2) Milestone path: rest-of-world cumulative commitments needed for 1,720 / 1,800 / 1,825
fig, ax = plt.subplots(figsize=(10, 5.4))
wk_hist = range(1, WK + 1)
avg5 = cum.loc[1:52, B5].mean(axis=1)
ax.plot(list(avg5.index), avg5.values, color=GREY, lw=1.5, label='5-yr average season (for shape)')
ax.plot(list(wk_hist), [cum.loc[w, cy] for w in wk_hist], color=INK, lw=2.6, label=f'{cy}/{str(cy+1)[-2:]} actual')
for X, col, lw, ls in ((BASE, AQUA, 1.6, (0, (4, 2))), (DESK, BLUE, 2.6, '-'), (USDA, RED, 1.6, (0, (4, 2)))):
    xs = list(paths[X].keys()); ys = list(paths[X].values())
    ax.plot(xs, ys, color=col, lw=lw, ls=ls, label=f'path to {X:,}' + (' (desk)' if X == DESK else (' (USDA)' if X == USDA else '')))
for w in MIL:
    ax.axvline(w, color=GRID, lw=1, zorder=0)
    ax.plot(w, paths[DESK][w], 'o', color=BLUE, ms=7, mec=SURF, mew=2, zorder=5)
    ax.text(w, paths[DESK][w] + 0.6, f'{paths[DESK][w]:.1f}', ha='center', fontsize=9, fontweight='bold', color=BLUE)
    ax.text(w, paths[USDA][w] - 1.1, f'{paths[USDA][w]:.1f}', ha='center', fontsize=8.5, color=RED)
    ax.text(w, 1.0, date(w), ha='center', fontsize=8.5, color=INK2)
ax.set_xlim(1, 52); ax.set_ylim(0, 27); ax.set_xlabel('Week of the season (Sep–Aug)'); ax.set_ylabel('Rest-of-world commitments, Mt (excl. China & Unknown)')
ax.spines[['top', 'right']].set_visible(False); ax.grid(axis='y', color=GRID, lw=0.8); ax.legend(frameon=False, loc='upper left', fontsize=9)
ax.set_title('Milestones: rest-of-world bookings needed (China fixed at 25 Mt)', loc='left', fontsize=13, color=INK)
fig.text(0.01, 0.01, 'Each Thursday compare the export sales total for all destinations except China and Unknown with the blue path. Above blue = 1,800+ on track; near red = drifting to USDA.',
         fontsize=7.5, color=INK2)
fig.tight_layout(rect=(0, 0.03, 1, 1)); fig.savefig(OUT + '2_milestones.png', dpi=200); plt.close(fig)

# 3) Scenario map (optional attachment)
chinas = np.arange(21, 29); brz = np.arange(0, 7)
grid = np.array([[BASE + (c - 25) * PER_MT + RISK[b] for b in brz] for c in chinas])
cmap = LinearSegmentedColormap.from_list('div', [RED, '#f0efec', BLUE])
fig, ax = plt.subplots(figsize=(9.5, 6))
im = ax.imshow(grid - USDA, cmap=cmap, norm=TwoSlopeNorm(vcenter=0, vmin=-180, vmax=180), aspect='auto', origin='lower')
for i, c in enumerate(chinas):
    for j, b in enumerate(brz):
        v = grid[i, j]; ax.text(j, i, f'{v:,.0f}\n({v-USDA:+.0f})', ha='center', va='center', fontsize=8.5, color=INK, fontweight='bold' if (c == 25 and b == 2) else 'normal')
ax.set_xticks(range(len(brz)), [f'+{b}' for b in brz]); ax.set_yticks(range(len(chinas)), [f'{c} Mt' for c in chinas])
ax.set_xlabel('Extra Brazilian soybeans to non-China buyers vs 5-yr avg (Mt)'); ax.set_ylabel('China full-year purchases of US beans')
ax.set_title('Scenario map: US exports (mln bu) and gap to USDA 1,720', loc='left', fontsize=12, color=INK)
for s in ax.spines.values(): s.set_visible(False)
ax.add_patch(plt.Rectangle((1.5, 3.5), 1, 1, fill=False, ec=INK, lw=2.2))
fig.text(0.01, 0.01, 'Box = desk number (China 25 Mt, Brazil ~+2 Mt to non-China buyers). Red = below USDA.', fontsize=7.5, color=INK2)
cb = fig.colorbar(im, ax=ax, fraction=0.035, pad=0.02); cb.set_label('vs USDA (mln bu)'); cb.outline.set_visible(False)
fig.tight_layout(rect=(0, 0.03, 1, 1)); fig.savefig(OUT + '3_scenario_map.png', dpi=200); plt.close(fig)

# 4) Mind map
need = USDA * CONV; rwm = 24.445 - (PACE - DESK) * CONV; se = 2.769 * np.sqrt(1.1)
P = 1 - stats.t.cdf((need - (25 + rwm)) / se, df=9); BE = need - rwm
m = {w: paths[DESK][w] for w in MIL}
fig, ax = plt.subplots(figsize=(13, 7.6)); ax.set_xlim(0, 13); ax.set_ylim(0, 7.7); ax.axis('off')
def box(x, y, w, h, title, lines, ec):
    ax.add_patch(FancyBboxPatch((x, y), w, h, boxstyle='round,pad=0.04,rounding_size=0.12', fc='white', ec=ec, lw=2))
    ax.text(x + 0.15, y + h - 0.22, title, fontsize=11, fontweight='bold', color=INK, va='top')
    for k, l in enumerate(lines): ax.text(x + 0.15, y + h - 0.62 - k * 0.33, l, fontsize=8.8, color=INK2, va='top')
def arrow(x1, y1, x2, y2, c): ax.annotate('', xy=(x2, y2), xytext=(x1, y1), arrowprops=dict(arrowstyle='-|>', color=c, lw=1.6))
ax.add_patch(FancyBboxPatch((4.55, 3.0), 3.9, 1.45, boxstyle='round,pad=0.05,rounding_size=0.15', fc=BLUE, ec=BLUE))
ax.text(6.5, 4.12, 'Desk number: 1,800 mln bu', ha='center', fontsize=12.5, fontweight='bold', color='white')
ax.text(6.5, 3.72, f'USDA 1,720  ·  P(above USDA) ≈ {P:.0%}', ha='center', fontsize=10.5, color='white')
ax.text(6.5, 3.33, 'China fixed at 25 Mt', ha='center', fontsize=9.5, color='white')
box(0.2, 4.55, 4.0, 2.6, '1 · CHINA (fixed at 25 Mt)', ['Booked 19.3 Mt (China + 60% of Unknown)', f'USDA only right if China ≤ ~{BE:.0f} Mt',
    '±1 Mt China = ±38 mln bu', 'WATCH: China + Unknown sales should', 'fade as they approach 25 Mt.', 'Still buying above 25 Mt = upside'], BLUE)
box(8.8, 4.55, 4.0, 2.6, '2 · REST OF WORLD', [f'Booked {now:.1f} Mt (excl. China & Unknown)', f'Milestones for 1,800: {m[17]:.1f} Mt by {date(17)},',
    f'{m[26]:.1f} Mt by {date(26)}, {m[35]:.1f} Mt by {date(35)}', 'Weekly need: ~650 kt (Nov) → ~520 (Dec–Jan)', '→ ~400 (Feb). Last 4 weeks: 689 kt',
    'WATCH: weekly total vs the milestone path'], BLUE)
box(0.2, 0.25, 4.0, 2.6, '3 · BRAZIL', ['Non-China exports: 5-yr avg 26.3 Mt,', 'record 29.9 Mt; provider forecasts 33.5 Mt', 'Each +1 Mt ≈ −15 mln bu US exports',
    'Most exposed: Thailand, Bangladesh, Taiwan,', 'S. Korea, Netherlands, Turkey', 'WATCH: Brazil line-ups from Feb'], ORANGE)
box(8.8, 0.25, 4.0, 2.6, '4 · EUROPE (EU-27 + UK)', ['Market: 4.895 Mt imports — on pace', 'Booked 1.18 Mt; ≈2.8 Mt needed by end-Dec', 'Nov–Feb normal ≈160 kt/week',
    'Germany / NL: no forward book yet', 'WATCH: EUDR (30 Dec 2026?) — could keep', 'US beans in EU after Feb = upside'], ORANGE)
arrow(4.2, 5.4, 4.95, 4.45, BLUE); arrow(8.8, 5.4, 8.05, 4.45, BLUE); arrow(4.2, 2.0, 4.95, 3.0, ORANGE); arrow(8.8, 2.0, 8.05, 3.0, ORANGE)
ax.text(6.5, 2.75, 'UP to 1,825+:', ha='center', va='top', fontsize=9, fontweight='bold', color=BLUE)
ax.text(6.5, 2.48, 'China > 25 Mt · RoW above path\nBrazil non-China < 30 Mt', ha='center', va='top', fontsize=8.5, color=INK)
ax.text(6.5, 1.75, 'DOWN to 1,720:', ha='center', va='top', fontsize=9, fontweight='bold', color=RED)
ax.text(6.5, 1.48, 'RoW near red path · China < 23 Mt\nBrazil non-China > 33 Mt', ha='center', va='top', fontsize=8.5, color=INK)
ax.text(6.5, 7.5, 'What moves the US export number — and what to check each Thursday', ha='center', fontsize=13, fontweight='bold', color=INK)
fig.savefig(OUT + '4_mental_map.png', dpi=200, bbox_inches='tight'); plt.close(fig)

# 5) Probability vs China (optional attachment), centred on the desk number
cs = np.linspace(18, 29, 111); pr = 1 - stats.t.cdf((need - (cs + rwm)) / se, df=9)
fig, ax = plt.subplots(figsize=(9.5, 4.6)); ax.plot(cs, pr, color=BLUE, lw=2.4)
ax.axhline(0.5, color=GREY, lw=1, ls=(0, (4, 3))); ax.axvline(BE, color=RED, lw=1.2, ls=(0, (4, 3)))
ax.text(BE + 0.1, 0.06, f'break-even ≈{BE:.1f} Mt', color=RED, fontsize=9)
for c in (22, 23, 25, 28):
    pv = 1 - stats.t.cdf((need - (c + rwm)) / se, df=9); ax.plot(c, pv, 'o', color=BLUE, ms=7, mec=SURF, mew=2)
    ax.text(c, pv - 0.08, f'{pv:.0%}', ha='center', fontsize=10, fontweight='bold', color=INK)
ax.set_xlim(18, 29); ax.set_ylim(0, 1.02); ax.yaxis.set_major_formatter(matplotlib.ticker.PercentFormatter(1.0))
ax.set_xlabel('China full-year purchases of US soybeans (Mt)'); ax.set_ylabel('Probability US exports > 1,720')
ax.spines[['top', 'right']].set_visible(False); ax.grid(color=GRID, lw=0.8)
ax.set_title('Probability of beating USDA by China outcome (centred on the 1,800 desk number)', loc='left', fontsize=12, color=INK)
fig.tight_layout(); fig.savefig(OUT + '5_probability_vs_china.png', dpi=200); plt.close(fig)
print('P', round(P, 3), 'BE', round(BE, 2)); print(open(OUT + 'milestones.csv').read())
