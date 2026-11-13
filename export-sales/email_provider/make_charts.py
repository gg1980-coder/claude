"""Chart for the provider email: what 1,700 requires vs 10 seasons of history (ESR weekly data, week 9)."""
import matplotlib; matplotlib.use('Agg')
import matplotlib.pyplot as plt
OUT = '/home/user/claude/export-sales/email_provider/'
INK, INK2, GRID, SURF, BLUE, RED, GREY = '#0b0b0b', '#52514e', '#e4e3df', '#fcfcfb', '#2a78d6', '#e34948', '#c3c2b7'
plt.rcParams.update({'font.family': 'DejaVu Sans', 'font.size': 10, 'axes.edgecolor': GRID, 'axes.labelcolor': INK2, 'xtick.color': INK2,
                     'ytick.color': INK2, 'figure.facecolor': SURF, 'axes.facecolor': SURF})
S = ['16/17', '17/18', '18/19', '19/20', '20/21', '21/22', '22/23', '23/24', '24/25', '25/26']
rw_still = [15.09, 20.52, 18.71, 17.99, 14.09, 18.94, 13.47, 13.08, 18.11, 15.22]      # RoW (excl. China & Unknown) booked after week 9, Mt
share = [62.2, 55.9, 46.1, 46.5, 80.0, 56.0, 61.7, 54.5, 55.7, 41.9]                  # % of final exports already committed at week 9
fig, (a, b) = plt.subplots(1, 2, figsize=(13, 5))
a.bar(range(10), rw_still, color=GREY, width=0.62, zorder=3)
for i, v in enumerate(rw_still): a.text(i, v / 2, f'{v:.1f}', ha='center', fontsize=9, color=INK2)
a.axhline(11.5, color=RED, lw=2.2, zorder=4, label='needed for 1,700: 11.5 Mt')
a.axhline(14.2, color=BLUE, lw=2.2, ls=(0, (5, 3)), zorder=4, label='needed for 1,800: 14.2 Mt')
a.legend(frameon=False, loc='upper right', fontsize=9.5)
a.set_xticks(range(10), S, fontsize=9); a.set_ylim(0, 26); a.set_ylabel('Mt booked after week 9')
a.set_title('Rest of world: bookings after week 9 in each season\n(China fixed at 25 Mt)', loc='left', fontsize=11.5, color=INK)
a.spines[['top', 'right']].set_visible(False); a.grid(axis='y', color=GRID, lw=0.8, zorder=0)
b.bar(range(10), share, color=GREY, width=0.62, zorder=3)
for i, v in enumerate(share): b.text(i, v / 2, f'{v:.0f}%', ha='center', fontsize=9, color=INK2)
b.axhline(66.9, color=RED, lw=2.2, zorder=4, label='implied by 1,700: 67%')
b.axhline(63.1, color=BLUE, lw=2.2, ls=(0, (5, 3)), zorder=4, label='implied by 1,800: 63%')
b.legend(frameon=False, loc='upper right', fontsize=9.5)
b.set_xticks(range(10), S, fontsize=9); b.set_ylim(0, 100); b.set_ylabel('% of final exports already committed at week 9')
b.set_title('Share of the year already sold by week 9:\n1,700 implies one of the highest in 10 years', loc='left', fontsize=11.5, color=INK)
b.spines[['top', 'right']].set_visible(False); b.grid(axis='y', color=GRID, lw=0.8, zorder=0)
fig.text(0.01, 0.01, 'USDA FAS weekly export sales, week 9 (end-Oct). 2020/21 = record China front-loading; US supplies ran short in spring. Export-sales basis; '
         '1,700 mln bu = 45.2 Mt on this basis.', fontsize=7.5, color=INK2)
fig.tight_layout(rect=(0, 0.04, 1, 1)); fig.savefig(OUT + 'chart_1700_vs_history.png', dpi=200); plt.close(fig)
print('ok')
