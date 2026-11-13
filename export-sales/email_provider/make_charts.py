"""Black-and-grey charts for the provider email (Outlook style). ESR weekly data, week 9."""
import matplotlib; matplotlib.use('Agg')
import matplotlib.pyplot as plt
OUT = '/home/user/claude/export-sales/email_provider/'
BLK, DGR, LGR, GRID = '#000000', '#555555', '#c8c8c8', '#e6e6e6'
plt.rcParams.update({'font.family': 'Carlito', 'font.size': 11, 'axes.edgecolor': DGR, 'axes.labelcolor': BLK, 'xtick.color': BLK,
                     'ytick.color': BLK, 'figure.facecolor': 'white', 'axes.facecolor': 'white'})
S = ['16/17', '17/18', '18/19', '19/20', '20/21', '21/22', '22/23', '23/24', '24/25', '25/26']

def bars(ax, vals, fmt, lab_y):
    order = sorted(range(10), key=lambda i: vals[i])
    ax.bar(range(10), [vals[i] for i in order], color=LGR, edgecolor=DGR, linewidth=0.6, width=0.62, zorder=3)
    for k, i in enumerate(order):
        ax.text(k, lab_y(vals[i]), fmt(vals[i]), ha='center', fontsize=10, color=BLK, zorder=6)
    ax.set_xticks(range(10), [S[i] for i in order])
    ax.spines[['top', 'right']].set_visible(False); ax.grid(axis='y', color=GRID, lw=0.8, zorder=0)

# Chart 1: rest of world still to book after week 9
rw = [15.09, 20.52, 18.71, 17.99, 14.09, 18.94, 13.47, 13.08, 18.11, 15.22]
fig, a = plt.subplots(figsize=(8, 4.6))
bars(a, rw, lambda v: f'{v:.1f}', lambda v: v - 1.4 if v > 15.5 else 10.6)
a.axhline(12.6, color=BLK, lw=2.0, zorder=4, label='1,700 (yours): needs 12.6 Mt, below all 10 seasons')
a.axhline(13.1, color=BLK, lw=1.8, ls=(0, (6, 3)), zorder=4, label='1,720 (USDA): needs 13.1 Mt, equal to the weakest season')
a.axhline(14.2, color=DGR, lw=1.6, ls=(0, (1.5, 2)), zorder=4, label='1,760 (our base case): needs 14.2 Mt')
a.legend(frameon=False, loc='upper left', fontsize=10)
a.set_ylim(0, 27); a.set_ylabel('Mt booked after week 9')
a.set_title('Chart 1. Rest of world: Mt booked after week 9, last 10 seasons\nvs what each export number needs from here (China at 25 Mt)', loc='left', fontsize=12)
fig.text(0.01, 0.01, 'Source: USDA FAS weekly export sales. Rest of world = all except China and Unknown. Seasons sorted low to high.', fontsize=8.5, color=DGR)
fig.tight_layout(rect=(0, 0.04, 1, 1)); fig.savefig(OUT + 'chart1_row_vs_history.png', dpi=200); plt.close(fig)

# Chart 2: share of the year already sold by week 9
sh = [62.2, 55.9, 46.1, 46.5, 80.0, 56.0, 61.7, 54.5, 55.7, 41.9]
fig, a = plt.subplots(figsize=(8, 4.6))
bars(a, sh, lambda v: f'{v:.0f}%', lambda v: v - 6 if v > 70 else v / 2)
a.axhline(65.3, color=BLK, lw=2.0, zorder=4, label='1,700 and 1,720 imply 64-65% already sold')
a.axhline(63.1, color=DGR, lw=1.6, ls=(0, (1.5, 2)), zorder=4, label='1,760 (our base case) implies 63%')
a.legend(frameon=False, loc='upper left', fontsize=10)
a.set_ylim(0, 100); a.set_ylabel('% of final exports sold by week 9')
a.set_title('Chart 2. Share of the year already sold by week 9, last 10 seasons\nOnly 2020/21 (record China front-loading) was higher than 1,700/1,720 imply', loc='left', fontsize=12)
fig.text(0.01, 0.01, 'Source: USDA FAS weekly export sales (commitments at week 9 / final exports). Seasons sorted low to high.', fontsize=8.5, color=DGR)
fig.tight_layout(rect=(0, 0.04, 1, 1)); fig.savefig(OUT + 'chart2_share_sold.png', dpi=200); plt.close(fig)
print('ok')
