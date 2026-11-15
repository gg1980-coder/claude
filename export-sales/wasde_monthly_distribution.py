"""Distribution of USDA (WASDE) US export forecast changes by report month, over the full forecast cycle of each season
(first projection in May of year 1 to the October report of year 2). Seasons 2010/11-2025/26.
Usage: python3 wasde_monthly_distribution.py <wasde_us_exports_by_report.csv> <out.xlsx>"""
import sys, numpy as np, pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
src, out = sys.argv[1:3]
x = pd.read_csv(src, parse_dates=['ReleaseDate'])
rows = []
for (c, my), g in x.groupby(['Commodity', 'MarketYear']):
    y = int(my[:4])
    if not 2010 <= y <= 2025: continue
    g = g.sort_values('ReleaseDate').reset_index(drop=True)
    for i in range(1, len(g)):
        d = g.ReleaseDate[i]; k = (d.year - y) * 12 + d.month - 5          # 0 = May of year 1
        if 1 <= k <= 17:
            gap = (g.ReleaseDate[i].year - g.ReleaseDate[i-1].year) * 12 + g.ReleaseDate[i].month - g.ReleaseDate[i-1].month
            rows.append(dict(commodity=c, season=my, k=k, chg=g.Value[i] - g.Value[i-1], gap=gap))
T = pd.DataFrame(rows)
T = T[T.gap == 1]          # only true month-on-month steps (skips the few missing reports)
lab = {k: pd.Timestamp(2000 + (5 + k - 1) // 12, (5 + k - 1) % 12 + 1, 1).strftime('%b') + (' (yr 1)' if k <= 7 else ' (yr 2)') for k in range(1, 18)}
res = {}
for c in ('Corn', 'Soybeans'):
    t = T[T.commodity == c]; tot = t.chg.abs().sum(); r = []
    for k in range(1, 18):
        v = t[t.k == k].chg
        if v.empty: continue
        ch = v[v.abs() > 0.5]
        r.append({'Report': lab[k], 'Seasons': len(v), 'Changed %': len(ch) / len(v), 'Up %': (v > 0.5).mean(), 'Down %': (v < -0.5).mean(),
                  'Unchanged %': (v.abs() <= 0.5).mean(), 'Avg change': v.mean(), 'Avg size when changed': ch.abs().mean() if len(ch) else 0,
                  'Median size when changed': ch.abs().median() if len(ch) else 0, 'Largest up': max(v.max(), 0), 'Largest down': min(v.min(), 0),
                  'Share of all revisions (abs)': v.abs().sum() / tot})
    res[c] = pd.DataFrame(r)
with pd.ExcelWriter(out) as xw:
    for c, df in res.items(): df.to_excel(xw, sheet_name=c, index=False, startrow=3)
    T.to_excel(xw, sheet_name='Data (monthly steps)', index=False)
wb = load_workbook(out); thin = Side(style='thin', color='000000')
for c in res:
    ws = wb[c]
    ws['A1'] = f'US {c.lower()} exports: how USDA changes its forecast in each WASDE report (mln bu), seasons 2010/11–2025/26'
    ws['A1'].font = Font(name='Calibri', size=13, bold=True)
    ws['A2'] = ('yr 1 = the calendar year the season starts (first projection in May); yr 2 = the following year. Each row compares a report with the previous month. '
                'Share of all revisions = that month\'s share of the total absolute change across the season.')
    ws['A2'].font = Font(name='Calibri', size=9, italic=True)
    for row in ws.iter_rows(min_row=4, max_row=ws.max_row, max_col=12):
        for cell in row:
            cell.border = Border(left=thin, right=thin, top=thin, bottom=thin); cell.font = Font(name='Calibri', size=10, bold=cell.row == 4)
            if cell.row == 4: cell.alignment = Alignment(wrap_text=True, vertical='center'); cell.fill = PatternFill('solid', fgColor='F2F2F2')
            elif cell.column in (3, 4, 5, 6, 12): cell.number_format = '0%'
            elif cell.column > 6: cell.number_format = '+#,##0;-#,##0;0' if cell.column in (7, 10, 11) else '#,##0'
    ws.row_dimensions[4].height = 44; ws.column_dimensions['A'].width = 13
    for col in 'BCDEFGHIJKL': ws.column_dimensions[col].width = 11
wb.save(out)
pd.set_option('display.width', 250)
for c, df in res.items():
    print('==', c); print(df.round(2).to_string(index=False))
