"""Chart for the provider email: rest-of-world bookings needed after week 9 for 1,700 / 1,720 / 1,760 vs 10 seasons (ESR weekly data)."""
import matplotlib; matplotlib.use('Agg')
import matplotlib.pyplot as plt
OUT = '/home/user/claude/export-sales/email_provider/'
INK, INK2, GRID, SURF, BLUE, RED, ORANGE, GREY = '#0b0b0b', '#52514e', '#e4e3df', '#fcfcfb', '#2a78d6', '#e34948', '#eb6834', '#c3c2b7'
plt.rcParams.update({'font.family': 'DejaVu Sans', 'font.size': 10, 'axes.edgecolor': GRID, 'axes.labelcolor': INK2, 'xtick.color': INK2,
                     'ytick.color': INK2, 'figure.facecolor': SURF, 'axes.facecolor': SURF})
S = ['16/17', '17/18', '18/19', '19/20', '20/21', '21/22', '22/23', '23/24', '24/25', '25/26']
rw = [15.09, 20.52, 18.71, 17.99, 14.09, 18.94, 13.47, 13.08, 18.11, 15.22]   # RoW (excl. China & Unknown) booked after week 9, Mt
order = sorted(range(10), key=lambda i: rw[i])
fig, a = plt.subplots(figsize=(10, 5))
a.bar(range(10), [rw[i] for i in order], color=GREY, width=0.62, zorder=3)
for k, i in enumerate(order): a.text(k, (rw[i] - 1.4 if rw[i] > 15.5 else 10.6), f'{rw[i]:.1f}', ha='center', fontsize=10, color=INK, zorder=5)
a.axhline(12.6, color=RED, lw=2.2, zorder=4, label='needed for 1,700 (provider): 12.6 Mt, below all 10 seasons')
a.axhline(13.1, color=ORANGE, lw=2.2, zorder=4, label='needed for 1,720 (USDA): 13.1 Mt, equal to the weakest season')
a.axhline(14.2, color=BLUE, lw=2.2, ls=(0, (5, 3)), zorder=4, label='needed for our base case 1,760: 14.2 Mt')
a.legend(frameon=False, loc='upper left', fontsize=9.5)
a.set_xticks(range(10), [S[i] for i in order], fontsize=9.5); a.set_ylim(0, 27); a.set_ylabel('Mt booked after week 9')
a.set_title('Rest of world: what each export number needs from here vs what it booked in the last 10 seasons\n(China fixed at 25 Mt; seasons sorted low to high)', loc='left', fontsize=11.5, color=INK)
a.spines[['top', 'right']].set_visible(False); a.grid(axis='y', color=GRID, lw=0.8, zorder=0)
fig.text(0.01, 0.01, 'USDA FAS weekly export sales, week 9 (end-Oct). Rest of world = all destinations except China and Unknown.', fontsize=7.5, color=INK2)
fig.tight_layout(rect=(0, 0.04, 1, 1)); fig.savefig(OUT + 'chart_1700_vs_history.png', dpi=200); plt.close(fig)
print('ok')
