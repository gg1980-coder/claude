"""Soybeans: weekly bookings by region, weeks 11-26, history vs what USDA 1,720 needs (China capped at 25 Mt incl. 60% of Unknown).
Usage: python3 soy_region_w11_26.py <esr_soy.csv> <out.xlsx> [usda_mbu=1720]"""
import sys, numpy as np, pandas as pd
src, out = sys.argv[1:3]; USDA = float(sys.argv[3]) if len(sys.argv) > 3 else 1720
K = 0.0272155; CHINA = 25.0; UNK_CN = 0.6; W0, W1 = 10, 26; NWK = W1 - W0
d = pd.read_csv(src); d['C'] = d.CountryName.str.strip(); d['dt'] = pd.to_datetime(d.WeekendingDate)
d['Y'] = d.ReportingMarketingYearName.str.extract(r'(\d{4})/').astype(int)
d = d[(d.dt >= pd.to_datetime(d.Y.astype(str) + '-08-31')) & (d.dt <= pd.to_datetime((d.Y + 1).astype(str) + '-09-07'))]
CY = int(d.loc[d.dt.idxmax(), 'Y']); WK = int(d[d.Y == CY].WeekNumber.max()); B5 = list(range(CY - 5, CY)); B10 = list(range(CY - 10, CY)); LY = CY - 1
RMAP = {'CHINA, PEOPLES REPUBLIC OF': 'China', 'MEXICO': 'Mexico', 'UNKNOWN': 'Unknown'}
for lst, nm in [(['JAPAN', 'TAIWAN', 'KOREA, REPUBLIC OF', 'HONG KONG'], 'East Asia ex-China'),
                (['INDONESIA', 'VIETNAM', 'MALAYSIA', 'THAILAND', 'PHILIPPINES', 'BURMA', 'CAMBODIA', 'LAOS', 'SINGAPORE'], 'Southeast Asia'),
                (['BANGLADESH', 'PAKISTAN', 'NEPAL', 'SRI LANKA', 'INDIA'], 'South Asia'),
                (['EGYPT', 'ALGERIA', 'TURKEY', 'TUNISIA', 'MOROCCO', 'IRAN', 'IRAQ', 'ISRAEL', 'JORDAN', 'LEBANON', 'LIBYA', 'SAUDI ARABIA', 'QATAR',
                  'UNITED ARAB EMIRATES'], 'Middle East & N. Africa'),
                (['SPAIN', 'NETHERLANDS', 'PORTUGAL', 'ITALY', 'GERMANY', 'UNITED KINGDOM', 'GREECE', 'BELGIUM-LUXEMBOURG', 'FRANCE', 'IRELAND', 'NORWAY',
                  'ROMANIA', 'SWITZERLAND', 'UKRAINE', 'GEORGIA, REPUBLIC OF'], 'Europe')]:
    for c in lst: RMAP[c] = nm
d['R'] = d.C.map(RMAP).fillna('Other (Americas, Africa)')
REGS = ['Mexico', 'East Asia ex-China', 'Southeast Asia', 'South Asia', 'Middle East & N. Africa', 'Europe', 'Other (Americas, Africa)']
cum = d.pivot_table(index=['R', 'WeekNumber'], columns='Y', values='TotalCommitment', aggfunc='sum').fillna(0) / 1e6
fin = d[d.Y < CY].sort_values('dt').groupby(['Y', 'C']).tail(1).groupby(['R', 'Y']).AccumulatedExports.sum().unstack().fillna(0) / 1e6
def c(r, y, w):
    try: return float(cum.loc[(r, w), y])
    except KeyError: return 0.0
now = {r: c(r, CY, WK) for r in REGS}; RW = sum(now.values())
need_total = USDA * K - CHINA - RW                       # rest of world (named) still to book for USDA, incl. the 40% of Unknown that will be named
usual = {r: np.mean([fin.loc[r, y] - c(r, y, W0) for y in B5]) for r in REGS}; U = sum(usual.values()); scale = need_total / U
rows = []
for r in REGS:
    rem = {y: fin.loc[r, y] - c(r, y, W0) for y in B5}
    share26 = sum(min(c(r, y, W1) - c(r, y, W0), rem[y]) for y in B5) / sum(rem.values())
    need = usual[r] * scale * share26 * 1000 / NWK
    hist = {y: (c(r, y, W1) - c(r, y, W0)) * 1000 / NWK for y in B10}
    a5 = np.mean([hist[y] for y in B5]); a10 = np.mean(list(hist.values()))
    last4 = (c(r, CY, WK) - c(r, CY, WK - 4)) * 1000 / 4
    rows.append({'Region': r, 'Booked so far (Mt)': now[r], 'Needed wk11-26 for USDA (kt/wk)': need, '5-yr avg wk11-26 (kt/wk)': a5,
                 '10-yr avg wk11-26 (kt/wk)': a10, f'Last season wk11-26 (kt/wk)': hist[LY], '10-yr low (kt/wk)': min(hist.values()),
                 '10-yr high (kt/wk)': max(hist.values()), 'Last 4 weeks actual (kt/wk)': last4,
                 'Needed vs 5-yr avg': need / a5 - 1 if a5 else np.nan, 'Needed vs last 4 weeks': need / last4 - 1 if last4 > 0 else np.nan})
T = pd.DataFrame(rows)
tot = {k: T[k].sum() for k in T.columns if k not in ('Region', 'Needed vs 5-yr avg', 'Needed vs last 4 weeks', '10-yr low (kt/wk)', '10-yr high (kt/wk)')}
hist_tot = {y: sum((c(r, y, W1) - c(r, y, W0)) * 1000 / NWK for r in REGS) for y in B10}
tot.update({'Region': 'Rest of world', '10-yr low (kt/wk)': min(hist_tot.values()), '10-yr high (kt/wk)': max(hist_tot.values()),
            'Needed vs 5-yr avg': tot['Needed wk11-26 for USDA (kt/wk)'] / tot['5-yr avg wk11-26 (kt/wk)'] - 1,
            'Needed vs last 4 weeks': tot['Needed wk11-26 for USDA (kt/wk)'] / tot['Last 4 weeks actual (kt/wk)'] - 1})
T = pd.concat([T, pd.DataFrame([tot])], ignore_index=True)
# China (+60% Unknown) for reference: 5-yr avg excludes 2025/26 (China absent)
cn_now = c('China', CY, WK) + UNK_CN * c('Unknown', CY, WK)
cn_hist = {y: ((c('China', y, W1) + UNK_CN * c('Unknown', y, W1)) - (c('China', y, W0) + UNK_CN * c('Unknown', y, W0))) * 1000 / NWK for y in B10}
cn_last4 = ((c('China', CY, WK) + UNK_CN * c('Unknown', CY, WK)) - (c('China', CY, WK - 4) + UNK_CN * c('Unknown', CY, WK - 4))) * 1000 / 4
china = {'Region': 'China (+60% Unknown), to reach 25 Mt', 'Booked so far (Mt)': cn_now, 'Needed wk11-26 for USDA (kt/wk)': np.nan,
         '5-yr avg wk11-26 (kt/wk)': np.mean([cn_hist[y] for y in B5 if y != LY]), '10-yr avg wk11-26 (kt/wk)': np.mean(list(cn_hist.values())),
         'Last season wk11-26 (kt/wk)': cn_hist[LY], '10-yr low (kt/wk)': min(cn_hist.values()), '10-yr high (kt/wk)': max(cn_hist.values()),
         'Last 4 weeks actual (kt/wk)': cn_last4}
print(f'USDA {USDA:.0f} = {USDA*K:.2f} Mt; China {CHINA}; RoW named booked {RW:.2f}; RoW still to book {need_total:.2f} Mt; scale vs 5-yr usual {scale:.2f}; China still to book {CHINA - cn_now:.2f} Mt')
pd.set_option('display.width', 250)
print(T.round(2).to_string(index=False)); print(pd.DataFrame([china]).round(2).to_string(index=False))
with pd.ExcelWriter(out) as xw:
    pd.concat([T, pd.DataFrame([china])]).to_excel(xw, sheet_name='Weeks 11-26', index=False, startrow=3)
from openpyxl import load_workbook
from openpyxl.styles import Font, Alignment, Border, Side
wb = load_workbook(out); ws = wb['Weeks 11-26']; thin = Side(style='thin', color='000000')
ws['A1'] = f'US soybeans: weekly bookings by region, weeks 11-26 (mid-Nov to late Feb), vs what USDA {USDA:,.0f} mln bu needs'; ws['A1'].font = Font(name='Calibri', size=13, bold=True)
ws['A2'] = (f'China capped at 25 Mt incl. 60% of Unknown. Rest of world still to book {need_total:.1f} Mt, shared by each region\'s usual (5-yr) bookings from week 10 '
            f'(scale {scale:.2f}) and its 5-yr share booked by week 26. Bookings = weekly change in total commitments (incl. Unknown later named).')
ws['A2'].font = Font(name='Calibri', size=9, italic=True)
for row in ws.iter_rows(min_row=4, max_row=ws.max_row, max_col=ws.max_column):
    for cell in row:
        cell.border = Border(left=thin, right=thin, top=thin, bottom=thin); cell.font = Font(name='Calibri', size=10, bold=cell.row == 4 or cell.row == ws.max_row - 1)
        if cell.row == 4: cell.alignment = Alignment(wrap_text=True, vertical='center')
        elif cell.column == 2: cell.number_format = '0.00'
        elif cell.column in (10, 11): cell.number_format = '+0%;-0%;0%'
        elif cell.column > 2: cell.number_format = '#,##0'
ws.row_dimensions[4].height = 58; ws.column_dimensions['A'].width = 34
for col in 'BCDEFGHIJK': ws.column_dimensions[col].width = 12
wb.save(out)
