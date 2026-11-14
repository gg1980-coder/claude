"""Black-and-grey charts for the corn research email. Runs build_corn.py logic for the numbers (data/esr_corn csv)."""
import sys, os, numpy as np
HERE = os.path.dirname(os.path.abspath(__file__)); ROOT = os.path.dirname(HERE)
os.chdir(ROOT)
sys.argv = ['build_corn.py', 'data/esr_corn_2016-2027.csv', os.path.join(HERE, '_tmp.xlsx'), '3275', HERE]
exec(open('build_corn.py').read())
os.remove(os.path.join(HERE, '_tmp.xlsx'))
os.replace(os.path.join(HERE, 'corn_chart1_after_wk_vs_usda.png'), os.path.join(HERE, 'chart1_after_wk10_vs_usda.png'))
import matplotlib.pyplot as plt
W = list(range(1, 36))
fig, a = plt.subplots(figsize=(8, 4.6))
a.plot(W, [tc[(LY, w)] for w in W], color=DGR, lw=1.6, label=f'{lab(LY)} (record year, {WASDE_FINALS[LY]:,} mln bu)')
a.plot(W, [np.mean([tc[(y, w)] for y in B5]) for w in W], color=DGR, lw=1.4, ls=(0, (5, 3)), label='5-yr average')
fw = list(range(WK, 36))
a.plot(fw, [now + (BAR - now) * (shape[w] if w > WK else 0) for w in fw], color=BLK, lw=1.8, ls=(0, (1.5, 2)), label=f'Path needed for USDA {USDA:,.0f}')
a.plot(list(range(1, WK + 1)), [tc[(CY, w)] for w in range(1, WK + 1)], color=BLK, lw=2.6, label=f'{CUR} (this season)')
a.annotate(f'{now:.1f} Mt\n(5 Nov)', (WK, now), xytext=(WK + 1.2, now - 9), fontsize=10, arrowprops=dict(arrowstyle='-', color=DGR, lw=0.8))
a.legend(frameon=False, loc='upper left', fontsize=10); a.set_xlim(1, 35); a.set_ylim(0, 90)
a.set_xlabel('Week of the marketing year (week 13 = late Nov, 17 = late Dec, 22 = late Jan, 35 = end Apr)'); a.set_ylabel('Mt sold (cumulative)')
a.spines[['top', 'right']].set_visible(False); a.grid(axis='y', color=GRID, lw=0.8)
a.set_title('Chart 2. US corn: cumulative sales this season vs last season, the 5-yr average\nand the path USDA 3,275 needs from here', loc='left', fontsize=12)
fig.text(0.01, 0.01, 'Source: USDA FAS weekly export sales (total commitments, all destinations), through w/e 5 Nov.', fontsize=8.5, color=DGR)
fig.tight_layout(rect=(0, 0.04, 1, 1)); fig.savefig(os.path.join(HERE, 'chart2_cumulative_paths.png'), dpi=200); plt.close(fig)
print('ok')
