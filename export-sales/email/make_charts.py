"""Charts for the desk email (week 9, 2026/27). Numbers come from the analysis workbooks (see email text)."""
import numpy as np, matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
from matplotlib.colors import LinearSegmentedColormap, TwoSlopeNorm
from matplotlib.patches import FancyBboxPatch
from scipy import stats
OUT = '/home/user/claude/export-sales/email/'
INK, INK2, GRID, SURF = '#0b0b0b', '#52514e', '#e4e3df', '#fcfcfb'
BLUE, RED, GREY, ORANGE = '#2a78d6', '#e34948', '#c3c2b7', '#eb6834'
plt.rcParams.update({'font.family': 'DejaVu Sans', 'font.size': 10, 'axes.edgecolor': GRID, 'axes.labelcolor': INK2,
                     'xtick.color': INK2, 'ytick.color': INK2, 'figure.facecolor': SURF, 'axes.facecolor': SURF})
USDA = 1720; BASE = 1825; PER_MT = 1 / 0.97729741421324 / 0.0272155   # mln bu per 1 Mt (ESR basis)
RISK = [0, -15, -29, -44, -57, -67, -77]                                # Brazil extra supply 0..6 Mt -> mln bu (switching workbook)

# 1) Bridge: USDA -> our pace -> base case -> planning case
steps = [('USDA\nNov WASDE', USDA, 'total'), ('China\nat 25 Mt', 92, 'd'), ('Mid East &\nN. Africa', 28, 'd'), ('Southeast\nAsia', 25, 'd'),
         ('Europe', 5, 'd'), ('Other\nAmericas/Africa', 3, 'd'), ('East Asia\nex-China', -8, 'd'), ('South\nAsia', -6, 'd'),
         ('Pace\nestimate', 1859, 'total'), ('Prudence\n(median season)', -34, 'd'), ('BASE CASE', BASE, 'total'),
         ('Brazil competes\n(+3 Mt)', -44, 'd'), ('Planning\ncase', 1781, 'total')]
fig, ax = plt.subplots(figsize=(13, 5.6))
run = 0
for i, (lab, v, kind) in enumerate(steps):
    if kind == 'total':
        col = INK2 if lab.startswith('USDA') else (BLUE if lab == 'BASE CASE' else '#7fa9de')
        ax.bar(i, v - 1600, bottom=1600, color=col, width=0.62, zorder=3); run = v
        ax.text(i, v + 6, f'{v:,.0f}', ha='center', va='bottom', fontsize=10, fontweight='bold', color=INK)
    else:
        col = BLUE if v > 0 else RED
        ax.bar(i, v, bottom=run, color=col, width=0.62, zorder=3, alpha=0.85)
        ax.text(i, max(run, run + v) + 6, f'{v:+,.0f}', ha='center', va='bottom', fontsize=9.5, color=INK)
        run += v
ax.axhline(USDA, color=INK2, lw=1, ls=(0, (4, 3)), zorder=2)
ax.set_xticks(range(len(steps)), [s[0] for s in steps], fontsize=8.5)
ax.set_ylim(1600, 1900); ax.set_ylabel('US soybean exports 2026/27 (mln bu)')
ax.spines[['top', 'right']].set_visible(False); ax.grid(axis='y', color=GRID, lw=0.8, zorder=0)
ax.set_title('From USDA to our number: where the extra ~105 mln bu comes from (and what could take it away)', loc='left', fontsize=12, color=INK)
fig.text(0.01, 0.01, 'Dashed line = USDA 1,720. Regional bars = each region\'s pace-based full year minus its 5-yr share of USDA\'s total (China fixed at 25 Mt). Data: USDA FAS weekly export sales through w/e 29-Oct-26 (wk 9).',
         fontsize=7.5, color=INK2)
fig.tight_layout(rect=(0, 0.03, 1, 1)); fig.savefig(OUT + '1_bridge.png', dpi=200); plt.close(fig)

# 2) Scenario grid: China x Brazil competition -> US exports vs USDA
chinas = np.arange(21, 29); brz = np.arange(0, 7)
grid = np.array([[BASE + (c - 25) * PER_MT + RISK[b] for b in brz] for c in chinas])
cmap = LinearSegmentedColormap.from_list('div', [RED, '#f0efec', BLUE])
fig, ax = plt.subplots(figsize=(9.5, 6))
im = ax.imshow(grid - USDA, cmap=cmap, norm=TwoSlopeNorm(vcenter=0, vmin=-180, vmax=180), aspect='auto', origin='lower')
for i, c in enumerate(chinas):
    for j, b in enumerate(brz):
        v = grid[i, j]
        ax.text(j, i, f'{v:,.0f}\n({v-USDA:+.0f})', ha='center', va='center', fontsize=8.5,
                color=INK, fontweight='bold' if (c == 25 and b in (0, 3)) else 'normal')
ax.set_xticks(range(len(brz)), [f'+{b}' for b in brz]); ax.set_yticks(range(len(chinas)), [f'{c} Mt' for c in chinas])
ax.set_xlabel('Extra Brazilian soybeans to non-China buyers vs 5-yr avg (Mt)  →  more competition for US beans')
ax.set_ylabel('China full-year purchases of US beans')
ax.set_title('Scenario map: US exports (mln bu) and gap to USDA 1,720', loc='left', fontsize=12, color=INK)
for s in ax.spines.values(): s.set_visible(False)
ax.add_patch(plt.Rectangle((-0.5, 3.5), 1, 1, fill=False, ec=INK, lw=2)); ax.add_patch(plt.Rectangle((2.5, 3.5), 1, 1, fill=False, ec=INK, lw=2, ls='--'))
fig.text(0.01, 0.01, 'Solid box = base case (China 25 Mt, normal Brazil). Dashed = planning case (Brazil sells to non-China buyers at last season\'s record pace). Red = below USDA.',
         fontsize=7.5, color=INK2)
cb = fig.colorbar(im, ax=ax, fraction=0.035, pad=0.02); cb.set_label('vs USDA (mln bu)'); cb.outline.set_visible(False)
fig.tight_layout(rect=(0, 0.03, 1, 1)); fig.savefig(OUT + '2_scenario_map.png', dpi=200); plt.close(fig)

# 3) Probability of beating USDA vs China outcome (rest of world on pace, 10-season error)
need = USDA / PER_MT; rw = 24.445 - (1859 - BASE) / PER_MT; se = 2.769 * np.sqrt(1 + 1 / 10)   # centred on the base case
cs = np.linspace(18, 29, 111); pr = 1 - stats.t.cdf((need - (cs + rw)) / se, df=9)
fig, ax = plt.subplots(figsize=(9.5, 4.6))
ax.plot(cs, pr, color=BLUE, lw=2.4)
ax.axhline(0.5, color=GREY, lw=1, ls=(0, (4, 3))); be = need - rw
ax.axvline(be, color=RED, lw=1.2, ls=(0, (4, 3))); ax.text(be + 0.1, 0.06, f'break-even\n≈{be:.1f} Mt', color=RED, fontsize=9)
ax.axvline(19.3, color=INK2, lw=1); ax.text(19.35, 0.82, 'China booked\nso far 19.3 Mt\n(incl. 60% of\nUnknown)', fontsize=8.5, color=INK2)
for c in (22, 23, 25, 28):
    pv = 1 - stats.t.cdf((need - (c + rw)) / se, df=9); ax.plot(c, pv, 'o', color=BLUE, ms=7, mec=SURF, mew=2)
    ax.text(c, pv - 0.08, f'{pv:.0%}', ha='center', fontsize=10, fontweight='bold', color=INK)
ax.set_xlim(18, 29); ax.set_ylim(0, 1.02); ax.yaxis.set_major_formatter(matplotlib.ticker.PercentFormatter(1.0))
ax.set_xlabel('China full-year purchases of US soybeans (Mt)'); ax.set_ylabel('Probability US exports > 1,720')
ax.spines[['top', 'right']].set_visible(False); ax.grid(color=GRID, lw=0.8)
ax.set_title('China is the swing factor: probability of beating USDA by China outcome', loc='left', fontsize=12, color=INK)
fig.text(0.01, 0.01, 'Centred on the 1,825 base case (rest of world at its 5-yr pace less a prudence haircut); uncertainty = 10-season back-tested error (±2.8 Mt). Normal Brazil competition.', fontsize=7.5, color=INK2)
fig.tight_layout(rect=(0, 0.03, 1, 1)); fig.savefig(OUT + '3_probability_vs_china.png', dpi=200); plt.close(fig)

# 4) Mental map: what moves the number and what to watch weekly
fig, ax = plt.subplots(figsize=(13, 7.4)); ax.set_xlim(0, 13); ax.set_ylim(0, 7.7); ax.axis('off')
def box(x, y, w, h, title, lines, ec, tsz=11):
    ax.add_patch(FancyBboxPatch((x, y), w, h, boxstyle='round,pad=0.04,rounding_size=0.12', fc='white', ec=ec, lw=2))
    ax.text(x + 0.15, y + h - 0.22, title, fontsize=tsz, fontweight='bold', color=INK, va='top')
    for k, l in enumerate(lines): ax.text(x + 0.15, y + h - 0.62 - k * 0.33, l, fontsize=8.6, color=INK2, va='top')
def arrow(x1, y1, x2, y2, c): ax.annotate('', xy=(x2, y2), xytext=(x1, y1), arrowprops=dict(arrowstyle='-|>', color=c, lw=1.6))
ax.add_patch(FancyBboxPatch((4.55, 3.0), 3.9, 1.45, boxstyle='round,pad=0.05,rounding_size=0.15', fc=BLUE, ec=BLUE))
ax.text(6.5, 4.12, 'US soybean exports 2026/27', ha='center', fontsize=12, fontweight='bold', color='white')
ax.text(6.5, 3.72, 'Our base 1,825  |  USDA 1,720', ha='center', fontsize=11, color='white')
ax.text(6.5, 3.33, 'P(> USDA) ≈ 82%  ·  10/10 seasons above', ha='center', fontsize=9.5, color='white')
box(0.2, 4.55, 4.0, 2.6, '1 · CHINA  (biggest lever)', ['Booked 19.3 of 25 Mt (incl. 60% of Unknown)', 'Break-even for USDA ≈ 22 Mt', '±1 Mt China = ±38 mln bu',
    'WATCH: China + Unknown weekly sales;', 'slowdown once ~25 Mt booked = cap holds;', 'continued buying = upside'], BLUE)
box(8.8, 4.55, 4.0, 2.6, '2 · REST OF WORLD PACE', ['Booked 8.7 Mt; usually books ~15.8 Mt more', 'USDA needs only ~76% of that (12.0 Mt)', 'Nov–Feb: normal ≈ 580 kt/wk,',
    'USDA needs ≈ 440 kt/wk; weakest season 465', 'Last 4 wks: 689 kt/wk (on pace)', 'WATCH: weekly RoW net sales vs 440/580'], BLUE)
box(0.2, 0.25, 4.0, 2.6, '3 · BRAZIL COMPETITION', ['Brazil to non-China buyers: 5-yr avg 26.3 Mt,', 'record 29.9 Mt (2025/26)', 'Each +1 Mt ≈ −15 mln bu US exports',
    'Switchers: Thailand, Bangladesh, Taiwan,', 'S. Korea, NL, Turkey; captive: Mexico, Egypt,', 'Indonesia, Japan · WATCH: Brazil line-ups'], ORANGE)
box(8.8, 0.25, 4.0, 2.6, '4 · EUROPE  (EU-27 + UK)', ['Market: 4.895 Mt imports (≈4.23 Mt our basis)', 'Booked 1.18 Mt: on pace for both', 'Nov–Feb need ≈155 kt/wk (normal 163)',
    'Checkpoint: ≈2.8 Mt booked by end-Dec', 'Germany/NL: zero forward book', 'WATCH: EUDR timing, NL/DE forward sales'], ORANGE)
arrow(4.2, 5.4, 4.95, 4.45, BLUE); arrow(8.8, 5.4, 8.05, 4.45, BLUE); arrow(4.2, 2.0, 4.95, 3.0, ORANGE); arrow(8.8, 2.0, 8.05, 3.0, ORANGE)
ax.text(6.5, 2.45, 'Upside if: China > 25 Mt · RoW > 580 kt/wk\nDownside if: China < 22 Mt · RoW < 440 kt/wk ·\nBrazil non-China exports > 31 Mt',
        ha='center', va='top', fontsize=9, color=INK)
ax.text(6.5, 7.45, 'What moves the US export number — and what to watch each Thursday', ha='center', fontsize=13, fontweight='bold', color=INK)
fig.savefig(OUT + '4_mental_map.png', dpi=200, bbox_inches='tight'); plt.close(fig)
print('done; break-even', round(need - rw, 2), 'P25', round(float(1 - stats.t.cdf((need - (25 + rw)) / se, 9)), 3))
