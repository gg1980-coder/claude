"""US corn export sales pace vs USDA: same approach as the soybean study, adapted to corn.
Usage: python3 build_corn.py <esr_corn.csv> <out.xlsx> [usda_mbu=3275] [chart_dir]
- No China anchor (China is a small corn buyer now): the model is all destinations, "sold so far + usual still to come".
- Unknown is not assigned to any country: by the end of the season it resolves into named destinations, and the
  historical "still to come" (final exports - commitments at week N, both including Unknown) already captures that.
- USDA (Census/balance-sheet) basis vs ESR basis: ratio from the WASDE finals available in the November 2026 WASDE
  (2024/25 2,873 and 2025/26 3,429 mln bu) vs ESR final accumulated exports. Only two seasons: treated as indicative.
All values are computed here and written as numbers (no live formulas)."""
import sys, os
import numpy as np, pandas as pd
from scipy.stats import t as T
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter as L

src, out = sys.argv[1], sys.argv[2]
USDA = float(sys.argv[3]) if len(sys.argv) > 3 else 3275
CHART_DIR = sys.argv[4] if len(sys.argv) > 4 else None
K = 0.025401                         # Mt per mln bu of corn (56 lb)
WASDE_FINALS = {2024: 2873, 2025: 3429}   # mln bu, USDA basis (Nov 2026 WASDE)

d = pd.read_csv(src)
d['C'] = d.CountryName.str.strip(); d['dt'] = pd.to_datetime(d.WeekendingDate)
d['Y'] = d.ReportingMarketingYearName.str.extract(r'(\d{4})/').astype(int)
d = d[(d.dt >= pd.to_datetime(d.Y.astype(str) + '-08-31')) & (d.dt <= pd.to_datetime((d.Y + 1).astype(str) + '-09-07'))]
CY = int(d.loc[d.dt.idxmax(), 'Y']); WK = int(d[d.Y == CY].WeekNumber.max()); LAST = d[d.Y == CY].dt.max()
PRIOR = list(range(CY - 10, CY)); B5 = list(range(CY - 5, CY)); LY = CY - 1
lab = lambda y: f'{y}/{str(y + 1)[-2:]}'
CUR = lab(CY)

REG = {}
for lst, nm in [(['MEXICO'], 'Mexico'), (['JAPAN'], 'Japan'), (['KOREA, REPUBLIC OF'], 'South Korea'), (['COLOMBIA'], 'Colombia'),
                (['CHINA, PEOPLES REPUBLIC OF'], 'China'),
                (['TAIWAN', 'HONG KONG', 'VIETNAM', 'INDONESIA', 'PHILIPPINES', 'MALAYSIA', 'THAILAND', 'BANGLADESH', 'SINGAPORE', 'BURMA',
                  'CAMBODIA', 'SRI LANKA', 'PAKISTAN', 'INDIA', 'NEPAL'], 'Other Asia (Taiwan, SE Asia)'),
                (['GUATEMALA', 'HONDURAS', 'COSTA RICA', 'DOMINICAN REPUBLIC', 'EL SALVADOR', 'PANAMA', 'NICARAGUA', 'JAMAICA', 'CUBA', 'HAITI',
                  'BARBADOS', 'LEEWARD-WINDWARD ISLANDS', 'TRINID', 'BELIZE', 'BAHAMAS'], 'Central America & Caribbean'),
                (['PERU', 'VENEZUELA', 'CHILE', 'ECUADOR', 'GUYANA', 'SURINAM'], 'Other South America'),
                (['SPAIN', 'PORTUGAL', 'NETHERLANDS', 'IRELAND', 'UNITED KINGDOM', 'ITALY', 'GERMANY', 'FRANCE', 'BELGIUM-LUXEMBOURG', 'GREECE'], 'Europe'),
                (['SAUDI ARABIA', 'MOROCCO', 'EGYPT', 'ISRAEL', 'TUNISIA', 'ALGERIA', 'JORDAN', 'IRAQ', 'UNITED ARAB EMIRATES', 'OMAN', 'LIBYA',
                  'TURKEY', 'LEBANON', 'KUWAIT', 'QATAR', 'YEMEN'], 'Middle East & N. Africa'),
                (['UNKNOWN'], 'Unknown')]:
    for c_ in lst: REG[c_] = nm
d['R'] = d.C.map(REG).fillna('Canada, Africa & other')
REGS = ['Mexico', 'Japan', 'South Korea', 'Colombia', 'Central America & Caribbean', 'Other South America', 'Other Asia (Taiwan, SE Asia)',
        'Europe', 'Middle East & N. Africa', 'China', 'Canada, Africa & other', 'Unknown']

tc = d.groupby(['Y', 'WeekNumber']).TotalCommitment.sum() / 1e6
ae = d.groupby(['Y', 'WeekNumber']).AccumulatedExports.sum() / 1e6
fin = {y: d[(d.Y == y) & (d.WeekNumber == d[d.Y == y].WeekNumber.max())].AccumulatedExports.sum() / 1e6 for y in PRIOR}
rtc = d.groupby(['Y', 'R', 'WeekNumber']).TotalCommitment.sum() / 1e6
rfin = {(y, r): v / 1e6 for (y, r), v in d[d.Y < CY].sort_values('dt').groupby(['Y', 'C']).tail(1).groupby(['Y', 'R']).AccumulatedExports.sum().items()}
def rt(y, r, w): return float(rtc.get((y, r, w), 0.0))
def rf(y, r): return float(rfin.get((y, r), 0.0))

now, shipped = tc[(CY, WK)], ae[(CY, WK)]
ratio_by = {y: fin[y] / (WASDE_FINALS[y] * K) for y in WASDE_FINALS if y in fin}
RATIO = float(np.mean(list(ratio_by.values())))
BAR_RAW, BAR = USDA * K, USDA * K * RATIO        # USDA in Mt, and on ESR basis
to_mbu = lambda mt: mt / RATIO / K                 # ESR Mt -> mln bu, USDA basis

# ---------- pace methods + backtest ----------
after = {y: fin[y] - tc[(y, WK)] for y in PRIOR}; mult = {y: fin[y] / tc[(y, WK)] for y in PRIOR}
def loo(w, kind):
    e = []
    for y in PRIOR:
        o = [x for x in PRIOR if x != y]
        p = tc[(y, w)] + np.mean([fin[x] - tc[(x, w)] for x in o]) if kind == 'rem' else tc[(y, w)] * np.mean([fin[x] / tc[(x, w)] for x in o])
        e.append(p - fin[y])
    return np.array(e)
BTW = sorted(set([5, 9, 13, 17, 22, 26, 30, 35, 40, WK]))
bt = {w: (np.sqrt((loo(w, 'rem') ** 2).mean()), np.sqrt((loo(w, 'mul') ** 2).mean())) for w in BTW}
err_wk = loo(WK, 'rem'); SE_R = bt[WK][0] * np.sqrt(1 + 1 / len(PRIOR)); SE_M = bt[WK][1] * np.sqrt(1 + 1 / len(PRIOR))
pab = lambda mean, se, bar: float(1 - T.cdf((bar - mean) / se, len(PRIOR) - 1))
t80 = T.ppf(0.9, len(PRIOR) - 1)
methods = []
for nm, yrs, kind in [('Remaining sales, 5-yr (RECOMMENDED)', B5, 'rem'), ('Remaining sales, 10-yr', PRIOR, 'rem'),
                      ('Multiplier ("share sold by now"), 5-yr', B5, 'mul'), ('Multiplier, 10-yr', PRIOR, 'mul'),
                      (f'Remaining sales, last season only ({lab(LY)}, a record year)', [LY], 'rem')]:
    m = now + np.mean([after[y] for y in yrs]) if kind == 'rem' else now * np.mean([mult[y] for y in yrs])
    se = SE_R if kind == 'rem' else SE_M
    methods.append((nm, m, to_mbu(m), se, to_mbu(m - t80 * se), to_mbu(m + t80 * se), pab(m, se, BAR)))
BASE = methods[0][1]
patterns = [(lab(y), now + after[y], to_mbu(now + after[y]), now * mult[y], to_mbu(now * mult[y])) for y in PRIOR]

# ---------- regions ----------
reg_rows = []
NAMED = [r for r in REGS if r != 'Unknown']
rem5_r = {r: np.mean([rf(y, r) - rt(y, r, WK) for y in B5]) for r in REGS}
named_now = sum(rt(CY, r, WK) for r in NAMED); named_rem = sum(rem5_r[r] for r in NAMED)
need_scale = (BAR - named_now) / named_rem          # named regions must deliver the USDA bar in full (Unknown resolves to 0)
base_scale = (BASE - named_now) / named_rem         # same split for our 5-yr pace estimate
for r in REGS:
    n = rt(CY, r, WK); ly = rt(LY, r, WK); a5 = np.mean([rt(y, r, WK) for y in B5]); a10 = np.mean([rt(y, r, WK) for y in PRIOR])
    f5 = np.mean([rf(y, r) for y in B5]); f10 = np.mean([rf(y, r) for y in PRIOR]); fly = rf(LY, r)
    if r == 'Unknown':
        rem5, exp5, need, fy_usda, rank = -n, 0.0, -n, 0.0, None
    else:
        rem5 = rem5_r[r]; exp5 = n + rem5 * base_scale; need = rem5 * need_scale; fy_usda = n + need
    hist = sorted(rf(y, r) for y in PRIOR)
    if r != 'Unknown': rank = sum(h > fy_usda for h in hist) + 1
    reg_rows.append((r, n, ly, a5, a10, n / ly - 1 if ly > 0.01 else None, n / a5 - 1 if a5 > 0.01 else None, fly, f5, f10, rem5, exp5, need, fy_usda, max(hist), rank))

# ---------- destinations ----------
dw = d[d.WeekNumber == WK].groupby(['Y', 'C']).TotalCommitment.sum().unstack(0).fillna(0) / 1e3
dfin = d[d.Y < CY].sort_values('dt').groupby(['Y', 'C']).tail(1).groupby(['C', 'Y']).AccumulatedExports.sum().unstack().fillna(0) / 1e3
dest = pd.DataFrame({'now': dw.get(CY, 0), 'LY': dw.get(LY, 0), 'a5': dw[B5].mean(axis=1), 'a10': dw[PRIOR].mean(axis=1)})
dest['R'] = [REG.get(c_, 'Canada, Africa & other') for c_ in dest.index]
dest['fy5'] = dfin.reindex(dest.index)[B5].mean(axis=1).fillna(0); dest['fyly'] = dfin.reindex(dest.index)[LY].fillna(0)
dest = dest[(dest.now > 0.5) | (dest.a5 > 50) | (dest.LY > 50)].sort_values('now', ascending=False)

# ---------- weekly pace + milestones ----------
def wk_sales(y, w): return (tc[(y, w)] - tc[(y, w - 1)]) * 1000
last4 = np.mean([wk_sales(CY, w) for w in range(WK - 3, WK + 1)])
last4_5 = np.mean([wk_sales(y, w) for y in B5 for w in range(WK - 3, WK + 1)])
NW = 53
num = {w: 0.0 for w in range(WK + 1, NW + 1)}; den = 0.0
for y in B5:
    den += after[y]
    for w in range(WK + 1, NW + 1):
        num[w] += min(tc.get((y, w), fin[y]) - tc[(y, WK)], after[y]) if w < NW else after[y]
shape = {w: num[w] / den for w in num}
week_date = lambda w: LAST + pd.Timedelta(weeks=w - WK)
MS_W = [w for w in (13, 17, 22, 26, 30, 35) if w > WK]
paths = {'USDA 3,275 (ESR basis)': BAR - now, 'Our pace estimate (5-yr)': BASE - now, '10-yr pace': methods[1][1] - now}
nov_w = [w for w in range(WK + 1, WK + 5)]
pace_need = {k: v * shape[nov_w[-1]] / 4 * 1000 for k, v in paths.items()}

# ================= workbook =================
NAVY = '1F3864'; HDR = PatternFill('solid', fgColor=NAVY); TOT = PatternFill('solid', fgColor='D9E1F2'); KEY = PatternFill('solid', fgColor='FFF2CC')
thin = Side(style='thin', color='BFBFBF'); BOX = Border(left=thin, right=thin, top=thin, bottom=thin)
def font(sz=10, b=False, col='000000', it=False): return Font(name='Calibri', size=sz, bold=b, color=col, italic=it)
def title(ws, t, sub=None):
    ws['A1'] = t; ws['A1'].font = font(14, True, NAVY)
    if sub: ws['A2'] = sub; ws['A2'].font = font(9, False, '595959', True)
def header(ws, r, cols, h=30):
    for j, v in enumerate(cols, 1):
        c_ = ws.cell(r, j, v); c_.font = font(10, True, 'FFFFFF'); c_.fill = HDR; c_.border = BOX
        c_.alignment = Alignment(wrap_text=True, vertical='center', horizontal='center')
    ws.row_dimensions[r].height = h
def put(ws, r, c_, v, fmt=None, f=None, fill=None, al=None):
    cell = ws.cell(r, c_, None if (isinstance(v, float) and np.isnan(v)) else v); cell.border = BOX
    if fmt: cell.number_format = fmt
    if f: cell.font = f
    if fill: cell.fill = fill
    if al: cell.alignment = al
    return cell
def note(ws, r, t, sz=9):
    ws.cell(r, 1, t).font = font(sz, False, '404040', True)
MT, BU, P0, KT, SP = '0.0', '#,##0', '0%', '#,##0', '+0%;-0%;0%'

wb = Workbook()
# ---- Summary
ws = wb.active; ws.title = 'Summary'
title(ws, f'US corn exports {CUR}: is USDA\'s {USDA:,.0f} mln bu achievable on the sales pace?',
      f'USDA FAS weekly export sales through w/e {LAST:%d %b %Y} (week {WK}); 10 seasons {lab(PRIOR[0])}–{lab(PRIOR[-1])}. USDA = Nov 2026 WASDE.')
rows = [
    ('USDA 2026/27 export forecast (mln bu)', USDA, BU), ('  in Mt (1 mln bu = 0.025401 Mt)', BAR_RAW, MT),
    ('  on weekly-sales (ESR) basis, Mt (× ESR/USDA ratio %.3f)' % RATIO, BAR, MT),
    (f'Sold so far, week {WK} (shipped + unshipped), Mt', now, MT), ('Shipped so far, Mt', shipped, MT),
    (f'Same week last season ({lab(LY)}), Mt', tc[(LY, WK)], MT), ('  change vs last season', now / tc[(LY, WK)] - 1, SP),
    (f'5-yr average at week {WK}, Mt', np.mean([tc[(y, WK)] for y in B5]), MT), ('  change vs 5-yr average', now / np.mean([tc[(y, WK)] for y in B5]) - 1, SP),
    ('Still to sell from here for USDA (ESR basis), Mt', BAR - now, MT),
    ('  same, without the basis adjustment, Mt', BAR_RAW - now, MT),
    (f'Sold after week {WK} in the last 10 seasons: lowest / average / highest, Mt', f'{min(after.values()):.1f} / {np.mean(list(after.values())):.1f} / {max(after.values()):.1f}', None),
    ('Share of USDA already sold', now / BAR, P0), (f'Share sold by week {WK} in the last 10 seasons: lowest / average / highest',
                                                    f'{min(tc[(y, WK)] / fin[y] for y in PRIOR):.0%} / {np.mean([tc[(y, WK)] / fin[y] for y in PRIOR]):.0%} / {max(tc[(y, WK)] / fin[y] for y in PRIOR):.0%}', None),
    ('Our pace estimate (remaining sales, 5-yr), mln bu', methods[0][2], BU), ('  80% range, mln bu', f'{methods[0][4]:,.0f} – {methods[0][5]:,.0f}', None),
    ('Chance of finishing at or above USDA (remaining-sales model)', methods[0][6], P0),
    ('Chance on the multiplier model (5-yr)', methods[2][6], P0),
    (f'Season patterns at or above USDA: remaining-sales / multiplier', f'{sum(p[2] >= USDA for p in patterns)} of 10 / {sum(p[4] >= USDA for p in patterns)} of 10', None),
    (f'Total weekly sales, last 4 weeks (kt/week) vs 5-yr avg same weeks', f'{last4:,.0f} vs {last4_5:,.0f}', None),
    (f'Weekly sales needed over the next 4 weeks for USDA (kt/week)', pace_need['USDA 3,275 (ESR basis)'], KT),
]
header(ws, 4, ['Item', 'Value'], h=20)
for i, (k_, v, fm) in enumerate(rows):
    put(ws, 5 + i, 1, k_, f=font(10, k_[0] != ' ')); put(ws, 5 + i, 2, v, fm, f=font(10, True), al=Alignment(horizontal='right'))
r = 6 + len(rows)
ws.cell(r, 1, 'Read-out').font = font(11, True, NAVY)
lines = [
    f'1. Sales are normal, the forecast is not. The book at week {WK} ({now:.1f} Mt) is {now / np.mean([tc[(y, WK)] for y in B5]) - 1:+.0%} vs the 5-yr average '
    f'but {now / tc[(LY, WK)] - 1:+.0%} vs last season. USDA\'s {USDA:,.0f} is a near-record year ({lab(LY)} was the record at {WASDE_FINALS[LY]:,} mln bu).',
    f'2. To reach USDA, buyers must book {BAR - now:.1f} Mt from here (ESR basis). The most ever booked after week {WK} was {max(after.values()):.1f} Mt, in the record {lab(LY)} season.',
    f'3. Every remaining-sales season pattern falls short of USDA (best: {max(p[2] for p in patterns):,.0f} mln bu). On the multiplier view only '
    f'{", ".join(p[0] for p in patterns if p[4] >= USDA) or "no season"} would get there.',
    f'4. The method\'s backtest error at week {WK} is large for corn (±{bt[WK][0]:.1f} Mt, about ±{bt[WK][0] / K:,.0f} mln bu) and recent seasons all finished above it '
    f'(late-season surges). The probability is therefore less certain than it looks; the robust point is that USDA needs a late season bigger than last year\'s record.',
    f'5. Destinations: Mexico, Japan, Colombia and South Korea are at or above their 5-yr level; the shortfall vs last year sits in Unknown, Mexico, South Korea, Taiwan, Spain '
    f'and Central America (see Regions / Destinations).',
    f'6. The case for USDA: sales have accelerated (last 4 weeks {last4:,.0f} kt/week vs a 5-yr average of {last4_5:,.0f}, led by Mexico), and the last three seasons had '
    f'growing late runs ({", ".join(f"{lab(y)} {after[y]:.1f}" for y in PRIOR[-3:])} Mt after week {WK}). Even so, USDA needs ~{pace_need["USDA 3,275 (ESR basis)"]:,.0f} kt/week '
    f'over the next 4 weeks, and a late season ~{(BAR - now) / max(after.values()) - 1:.0%} bigger than last year\'s record.',
]
for i, t_ in enumerate(lines):
    c_ = ws.cell(r + 1 + i, 1, t_); c_.font = font(10); c_.alignment = Alignment(wrap_text=True, vertical='top')
    ws.merge_cells(start_row=r + 1 + i, start_column=1, end_row=r + 1 + i, end_column=2); ws.row_dimensions[r + 1 + i].height = 44
note(ws, r + 8, f'ESR/USDA basis ratio {RATIO:.3f} = average of {", ".join(f"{lab(y)} {v:.3f}" for y, v in ratio_by.items())} (ESR final ÷ WASDE). Two seasons only: indicative. '
                'Without the adjustment the USDA bar is higher and the conclusion stronger.')
ws.column_dimensions['A'].width = 92; ws.column_dimensions['B'].width = 22

# ---- Pace & Probability
wp = wb.create_sheet('Pace & Probability')
title(wp, f'Pace methods vs USDA {USDA:,.0f} mln bu (bar on ESR basis: {BAR:.1f} Mt)',
      'Expected full year = this season\'s sales at week N + what past seasons sold after week N (remaining sales), or × their full-year/week-N ratio (multiplier).')
header(wp, 4, ['Method', 'Expected full year (Mt, ESR)', 'Expected (mln bu, USDA basis)', 'vs USDA (mln bu)', 'Uncertainty ±1 sd (Mt)',
               '80% range low (mln bu)', '80% range high (mln bu)', 'Chance at or above USDA'], h=44)
for i, (nm, m, mb, se, lo, hi, p) in enumerate(methods):
    rr = 5 + i
    put(wp, rr, 1, nm, f=font(10, i == 0)); put(wp, rr, 2, m, MT); put(wp, rr, 3, mb, BU, f=font(10, True)); put(wp, rr, 4, mb - USDA, '+#,##0;-#,##0')
    put(wp, rr, 5, se, MT); put(wp, rr, 6, lo, BU); put(wp, rr, 7, hi, BU); put(wp, rr, 8, p, P0, f=font(11, True, NAVY), al=Alignment(horizontal='center'))
    if i == 0:
        for j in range(1, 9): wp.cell(rr, j).fill = KEY
note(wp, 11, 'Uncertainty = leave-one-out backtest error at this week across the 10 seasons, widened for an out-of-sample year; Student-t with 9 df. '
             'The last-season-only line shows what happens if this year\'s late season repeats the 2025/26 record.')
r0 = 13
wp.cell(r0, 1, f'Season patterns: this season\'s week-{WK} sales + each past season\'s behaviour').font = font(11, True, NAVY)
header(wp, r0 + 1, ['Season pattern', 'Remaining sales: full year (Mt)', 'mln bu', 'Multiplier: full year (Mt)', 'mln bu', 'Remaining sales ≥ USDA?', 'Multiplier ≥ USDA?'], h=36)
for i, (s_, a_, ab, m_, mb) in enumerate(patterns):
    rr = r0 + 2 + i
    for j, (v, f_) in enumerate([(s_, None), (a_, MT), (ab, BU), (m_, MT), (mb, BU), ('yes' if ab >= USDA else 'no', None), ('yes' if mb >= USDA else 'no', None)]):
        put(wp, rr, 1 + j, v, f_)
r1 = r0 + 14
wp.cell(r1, 1, 'How likely is any USDA number? (remaining sales, 5-yr)').font = font(11, True, NAVY)
header(wp, r1 + 1, ['USDA export number (mln bu)', 'Bar on ESR basis (Mt)', 'Still to sell from here (Mt)', 'Chance at or above'], h=30)
for i, x in enumerate(range(2300, 3501, 100)):
    rr = r1 + 2 + i; b_ = x * K * RATIO
    put(wp, rr, 1, x, BU, f=font(10, x == int(USDA // 100 * 100))); put(wp, rr, 2, b_, MT); put(wp, rr, 3, b_ - now, MT); put(wp, rr, 4, pab(BASE, SE_R, b_), P0)
r2 = r1 + 16
wp.cell(r2, 1, 'Backtest: how wrong was each method at each week? (leave-one-out RMSE, Mt, 10 seasons)').font = font(11, True, NAVY)
header(wp, r2 + 1, ['Week', 'Remaining sales', 'Multiplier'], h=22)
for i, w in enumerate(BTW):
    put(wp, r2 + 2 + i, 1, w, '0', f=font(10, w == WK)); put(wp, r2 + 2 + i, 2, bt[w][0], MT); put(wp, r2 + 2 + i, 3, bt[w][1], MT)
r3 = r2 + 3 + len(BTW)
wp.cell(r3, 1, f'Remaining-sales errors by season at week {WK} (predicted − actual, Mt): negative = the season finished stronger than the method said').font = font(10, True, NAVY)
header(wp, r3 + 1, ['Season', 'Error (Mt)'], h=20)
for i, y in enumerate(PRIOR): put(wp, r3 + 2 + i, 1, lab(y)); put(wp, r3 + 2 + i, 2, float(err_wk[i]), '+0.0;-0.0')
wp.column_dimensions['A'].width = 52
for c_ in 'BCDEFGH': wp.column_dimensions[c_].width = 15

# ---- Regions
wr = wb.create_sheet('Regions')
title(wr, f'Corn sales by region at week {WK} vs last season and the 5/10-yr averages, and what USDA needs (Mt)',
      f'"Needed for USDA" = each region\'s usual (5-yr) sales from here (which include Unknown sales later named), all scaled by the same factor ({need_scale:.2f}×) so named destinations '
      f'reach the USDA bar. "Expected" uses the same split scaled to our 5-yr pace ({base_scale:.2f}×). Unknown is shown resolving to zero by the end of the season.')
cols = ['Region', f'Sold, wk {WK}', f'Last season wk {WK}', f'5-yr avg wk {WK}', f'10-yr avg wk {WK}', 'vs last season', 'vs 5-yr',
        f'Full year {lab(LY)}', '5-yr avg full year', '10-yr avg full year', '5-yr usual still to come (Unknown: resolves)', 'Expected full year (5-yr pace)',
        'Needed from here for USDA', 'Full year implied by USDA', 'Best of last 10 seasons', 'USDA-implied rank (1 = highest)']
header(wr, 4, cols, h=58)
for i, row in enumerate(reg_rows):
    rr = 5 + i
    for j, v in enumerate(row):
        fm = None if j == 0 else (SP if j in (5, 6) else ('0' if j == 15 else '0.00'))
        put(wr, rr, 1 + j, v, fm)
rt_ = 5 + len(reg_rows)
tots = [sum((x[j] or 0) for x in reg_rows) for j in range(1, 15)]
put(wr, rt_, 1, 'TOTAL', f=font(10, True), fill=TOT)
for j, v in enumerate(tots, 2):
    if j in (7, 8): v = None
    put(wr, rt_, j, v, '0.00', f=font(10, True), fill=TOT)
put(wr, rt_, 6, tots[0] / tots[1] - 1, SP, f=font(10, True), fill=TOT); put(wr, rt_, 7, tots[0] / tots[2] - 1, SP, f=font(10, True), fill=TOT)
wr.column_dimensions['A'].width = 30
for j in range(2, 17): wr.column_dimensions[L(j)].width = 11.5

# ---- Destinations
wd = wb.create_sheet('Destinations')
title(wd, f'Destinations at week {WK} (kt): who is ahead / behind last season and the 5-yr average',
      'Sorted by current sales. Includes destinations with sales this season or a meaningful 5-yr / last-season book.')
header(wd, 4, ['Destination', 'Region', f'Sold wk {WK}', f'Last season wk {WK}', f'5-yr avg wk {WK}', f'10-yr avg wk {WK}', 'vs last season (kt)',
               'vs 5-yr (kt)', 'vs 5-yr (%)', f'Full year {lab(LY)}', '5-yr avg full year', 'Read'], h=44)
for i, (cn, x) in enumerate(dest.iterrows()):
    rr = 5 + i; v5 = x.now - x.a5
    read = 'ahead' if v5 > max(25, 0.1 * x.a5) else ('behind' if v5 < -max(25, 0.1 * x.a5) else 'in line')
    vals = [cn.title() if cn != 'UNKNOWN' else 'Unknown', x.R, x.now, x.LY, x.a5, x.a10, x.now - x.LY, v5, (x.now / x.a5 - 1) if x.a5 > 1 else None, x.fyly, x.fy5, read]
    for j, v in enumerate(vals):
        fm = None if j in (0, 1, 11) else (SP if j == 8 else '#,##0;-#,##0')
        put(wd, rr, 1 + j, v, fm, f=font(10, j == 11 and read != 'in line'))
wd.column_dimensions['A'].width = 24; wd.column_dimensions['B'].width = 28
for j in range(3, 13): wd.column_dimensions[L(j)].width = 12

# ---- Milestones
wm = wb.create_sheet('Milestones')
title(wm, 'Weekly milestones: cumulative sales (all destinations, Mt) each path needs, and the weekly pace now',
      'Spread with the 5-yr seasonal pattern of sales after this week (share of the remaining season done by each week, all seasons pooled).')
header(wm, 4, ['Week', 'Week ending', 'Share of remaining season done'] + list(paths.keys()), h=40)
put(wm, 5, 1, WK, '0'); put(wm, 5, 2, f'{LAST:%d %b %Y} (today)'); put(wm, 5, 3, 0, P0)
for j in range(len(paths)): put(wm, 5, 4 + j, now, MT, f=font(10, True))
for i, w in enumerate(MS_W):
    rr = 6 + i; put(wm, rr, 1, w, '0'); put(wm, rr, 2, f'{week_date(w):%d %b %Y}'); put(wm, rr, 3, shape[w], P0)
    for j, v in enumerate(paths.values()): put(wm, rr, 4 + j, now + v * shape[w], MT)
rr = 7 + len(MS_W)
wm.cell(rr, 1, 'Weekly sales needed over the next 4 weeks (kt/week)').font = font(11, True, NAVY)
header(wm, rr + 1, ['Path', 'kt/week'], h=20)
for i, (k_, v) in enumerate(pace_need.items()): put(wm, rr + 2 + i, 1, k_); put(wm, rr + 2 + i, 2, v, KT)
put(wm, rr + 5, 1, 'Actual, last 4 weeks', f=font(10, True)); put(wm, rr + 5, 2, last4, KT, f=font(10, True))
put(wm, rr + 6, 1, '5-yr average, same 4 weeks'); put(wm, rr + 6, 2, last4_5, KT)
rw_ = rr + 8
wm.cell(rw_, 1, 'Weekly net sales, last 6 weeks (kt): this season vs 5-yr average').font = font(11, True, NAVY)
header(wm, rw_ + 1, ['Week', 'This season', '5-yr average'], h=20)
for i, w in enumerate(range(WK - 5, WK + 1)):
    put(wm, rw_ + 2 + i, 1, w, '0'); put(wm, rw_ + 2 + i, 2, wk_sales(CY, w), KT); put(wm, rw_ + 2 + i, 3, np.mean([wk_sales(y, w) for y in B5]), KT)
wm.column_dimensions['A'].width = 34; wm.column_dimensions['B'].width = 16
for c_ in 'CDEF': wm.column_dimensions[c_].width = 18

# ---- History
wh = wb.create_sheet('History')
title(wh, f'Season history at week {WK} (Mt, ESR basis)')
header(wh, 4, ['Season', f'Sold by wk {WK}', f'Shipped by wk {WK}', 'Final exports (ESR)', f'Sold after wk {WK}', 'Full year ÷ week sales',
               f'Share sold by wk {WK}', f'Share shipped by wk {WK}', 'WASDE final (mln bu)', 'ESR ÷ WASDE'], h=40)
for i, y in enumerate(PRIOR + [CY]):
    rr = 5 + i; f_ = fin.get(y)
    vals = [lab(y), tc[(y, WK)], ae[(y, WK)], f_, (f_ - tc[(y, WK)]) if f_ else None, (f_ / tc[(y, WK)]) if f_ else None,
            (tc[(y, WK)] / f_) if f_ else None, (ae[(y, WK)] / f_) if f_ else None, WASDE_FINALS.get(y), ratio_by.get(y)]
    for j, v in enumerate(vals):
        put(wh, rr, 1 + j, v, [None, MT, MT, MT, MT, '0.00', P0, P0, BU, '0.000'][j], f=font(10, y == CY))
wh.column_dimensions['A'].width = 12
for j in range(2, 11): wh.column_dimensions[L(j)].width = 14

# ---- Weekly data
wwd = wb.create_sheet('Weekly Data')
title(wwd, 'Cumulative total commitments by week (Mt), all destinations')
header(wwd, 4, ['Week'] + [lab(y) for y in PRIOR + [CY]], h=20)
for w in range(1, NW + 1):
    put(wwd, 4 + w, 1, w, '0')
    for j, y in enumerate(PRIOR + [CY]):
        v = tc.get((y, w)); put(wwd, 4 + w, 2 + j, float(v) if v is not None else None, MT)
wb.save(out)

# ---------- chart (black & grey, Outlook style) ----------
if CHART_DIR:
    import matplotlib; matplotlib.use('Agg'); import matplotlib.pyplot as plt
    BLK, DGR, LGR, GRID = '#000000', '#555555', '#c8c8c8', '#e6e6e6'
    plt.rcParams.update({'font.family': 'Carlito', 'font.size': 11, 'axes.edgecolor': DGR, 'figure.facecolor': 'white', 'axes.facecolor': 'white'})
    vals = [after[y] for y in PRIOR]; order = sorted(range(10), key=lambda i: vals[i])
    fig, a = plt.subplots(figsize=(8, 4.6))
    a.bar(range(10), [vals[i] for i in order], color=LGR, edgecolor=DGR, linewidth=0.6, width=0.62, zorder=3)
    for k_, i in enumerate(order):
        yl = vals[i] - 2.6 if abs(vals[i] - 2.6 - (BASE - now)) > 1.5 else vals[i] - 5.5
        a.text(k_, yl, f'{vals[i]:.1f}', ha='center', fontsize=10, zorder=6)
    a.axhline(BAR - now, color=BLK, lw=2.0, zorder=4, label=f'USDA {USDA:,.0f}: needs {BAR - now:.1f} Mt, above all 10 seasons')
    a.axhline(BASE - now, color=DGR, lw=1.6, ls=(0, (1.5, 2)), zorder=4, label=f'5-yr average: {BASE - now:.1f} Mt (≈{methods[0][2]:,.0f} mln bu)')
    a.legend(frameon=False, loc='upper left', fontsize=10)
    a.set_xticks(range(10), [lab(PRIOR[i])[2:] for i in order]); a.set_ylim(0, 65); a.set_ylabel(f'Mt sold after week {WK}')
    a.spines[['top', 'right']].set_visible(False); a.grid(axis='y', color=GRID, lw=0.8, zorder=0)
    a.set_title(f'US corn: Mt sold after week {WK}, last 10 seasons\nvs what USDA {USDA:,.0f} mln bu needs from here', loc='left', fontsize=12)
    fig.text(0.01, 0.01, f'Source: USDA FAS weekly export sales, through w/e {LAST:%d %b}. USDA bar on weekly-sales basis. Seasons sorted low to high.', fontsize=8.5, color=DGR)
    fig.tight_layout(rect=(0, 0.04, 1, 1)); fig.savefig(os.path.join(CHART_DIR, 'corn_chart1_after_wk_vs_usda.png'), dpi=200); plt.close(fig)

print(f'CUR {CUR} WK {WK} now {now:.2f} LY {tc[(LY, WK)]:.2f} bar {BAR:.2f} (raw {BAR_RAW:.2f}) ratio {RATIO:.3f} need {BAR - now:.2f} max_after {max(after.values()):.2f}')
for m in methods: print(f'  {m[0]:55s} {m[1]:6.1f} Mt {m[2]:6.0f} mbu  P={m[6]:.1%}')
print('  patterns rem', [round(p[2]) for p in patterns]); print('  patterns mul', [round(p[4]) for p in patterns])
print('  RMSE wk', WK, [round(x, 2) for x in bt[WK]], 'last4', round(last4), 'vs5y', round(last4_5), 'need4', {k: round(v) for k, v in pace_need.items()})
for row in reg_rows: print('  ', f'{row[0]:30s} now {row[1]:5.2f} LY {row[2]:5.2f} a5 {row[3]:5.2f} exp5 {row[11]:5.2f} needUSDA {row[12]:5.2f} FYusda {row[13]:5.2f} best {row[14]:5.2f} rank {row[15]}')
