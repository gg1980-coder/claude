import sys, glob, pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter as L
from openpyxl.chart import LineChart, Reference

inp, out = sys.argv[1], sys.argv[2]
d = pd.concat([pd.read_excel(f, header=2) for f in sorted(glob.glob(inp + '/*.xlsx'))])
d = d[d.Time.astype(str).str.contains('-')].copy()
d['Date'] = pd.to_datetime(d.Time)
REG = ['Africa','Asia','Asia - Other','Asia - South','Asia Near East','Australia and Oceania',
       'Central America and Caribbean','Europe','North America','South America','South/Central America']
d = d[~d.Countries.isin(REG)]
d['Countries'] = d.Countries.replace({'All Geographic Regions (World Total)': 'World Total'})
COM = {  # HS code prefix -> (sheet name, title)
    '1518004000': ('UCO', 'US Imports of UCO - HTS 1518.00.4000 Animal & Veg. Fats/Oils, Chemically Modified (mln lbs.)'),
    '1502100040': ('Inedible Tallow', 'US Inedible Tallow Imports - HTS 1502.10.0040 (mln lbs.)'),
    '1502100020': ('Edible Tallow', 'US Edible Tallow Imports - HTS 1502.10.0020 (mln lbs.)'),
}
d['HS'] = d.Commodities.str.split().str[0]
d = d[d.HS.isin(COM)]
d['Kg'] = d['Quantity 1 (Gen)'].fillna(0)
d = d[d.Kg != 0].sort_values(['HS', 'Countries', 'Date'])
last = d.Date.max()

F = 'Arial'
font = lambda **k: Font(name=F, size=9, **k)
green = PatternFill('solid', fgColor='E2EFDA'); grey = PatternFill('solid', fgColor='D9D9D9')
thin = Side(style='thin', color='808080'); box = Border(top=thin, bottom=thin, left=thin, right=thin)
NUM = '#,##0;-#,##0;0'
wb = Workbook()

# ---- Data sheet
ds = wb.active; ds.title = 'Data'
ds['A1'] = 'Conversion (lbs per kg)'; ds['B1'] = 2.20462262185
ds['B1'].font = font(color='0000FF'); ds['A1'].font = font(bold=True)
ds['C1'] = 'Source: U.S. Census Bureau, USA Trade Online, HS district-level imports, General Imports (Quantity 1 Gen, kg), all districts. Accessed Oct 7, 2026.'
ds['C1'].font = font(italic=True)
hdr = ['HS Code', 'Commodity', 'Country', 'Date', 'Year', 'Month', 'Quantity (kg)', 'Quantity (mln lbs)']
for i, h in enumerate(hdr, 1):
    c = ds.cell(3, i, h); c.font = font(bold=True); c.fill = grey
for r, row in enumerate(d.itertuples(), 4):
    ds.cell(r, 1, row.HS); ds.cell(r, 2, COM[row.HS][0]); ds.cell(r, 3, row.Countries)
    ds.cell(r, 4, row.Date.to_pydatetime()).number_format = 'mmm-yy'
    ds.cell(r, 5, row.Date.year); ds.cell(r, 6, row.Date.month)
    ds.cell(r, 7, int(row.Kg)).number_format = '#,##0'
    ds.cell(r, 8, f'=G{r}*$B$1/1000000').number_format = '#,##0.000'
    for c in range(1, 9): ds.cell(r, c).font = font()
N = r
for c, w in zip('ABCDEFGH', [12, 16, 22, 9, 6, 6, 14, 14]): ds.column_dimensions[c].width = w
ds.freeze_panes = 'A4'; ds.auto_filter.ref = f'A3:H{N}'
rng = lambda col: f'Data!${col}$4:${col}${N}'
sumifs = lambda com, ctry, y, m: (f'SUMIFS({rng("H")},{rng("B")},"{com}",{rng("C")},"{ctry}",'
                                 f'{rng("E")},{y},{rng("F")},{m})')

MON = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec']
years = sorted(d.Date.dt.year.unique())
summ = {}
for hs, (name, title) in COM.items():
    ws = wb.create_sheet(name)
    sub = d[d.HS == hs]
    ws.merge_cells('A1:N1'); ws['A1'] = title
    ws['A1'].font = Font(name=F, size=11, bold=True); ws['A1'].alignment = Alignment(horizontal='center')
    r = 3
    for y in years:
        yy = sub[(sub.Date.dt.year == y) & (sub.Countries != 'World Total')]
        ctry = yy.groupby('Countries').Kg.sum().sort_values(ascending=False).index.tolist()
        nm = 12 if y < last.year else last.month
        ws.cell(r, 1).fill = grey; ws.cell(r, 1).border = box
        for m in range(12):
            c = ws.cell(r, m + 2, f'{MON[m]}-{str(y)[2:]}'); c.font = font(bold=True); c.fill = grey
            c.alignment = Alignment(horizontal='center'); c.border = box
        c = ws.cell(r, 14, 'Total'); c.font = font(bold=True); c.fill = grey; c.border = box
        c.alignment = Alignment(horizontal='center')
        wr = r + 1; first, lastr = wr + 1, wr + max(len(ctry), 1)
        ws.cell(wr, 1, 'World Total').font = font(bold=True)
        for m in range(12):
            if m < nm:
                ws.cell(wr, m + 2, f'=SUM({L(m+2)}{first}:{L(m+2)}{lastr})')
        ws.cell(wr, 14, f'=SUM(B{wr}:M{wr})')
        for col in range(1, 15):
            c = ws.cell(wr, col); c.fill = green; c.border = box; c.number_format = NUM
            if col > 1: c.font = font(bold=True)
        summ[(name, y)] = f"'{name}'!N{wr}"
        for i, k in enumerate(ctry):
            rr = first + i
            ws.cell(rr, 1, k)
            for m in range(nm):
                ws.cell(rr, m + 2, f'={sumifs(name, k, y, m + 1)}')
            ws.cell(rr, 14, f'=SUM(B{rr}:M{rr})')
            for col in range(1, 15):
                c = ws.cell(rr, col); c.font = font(bold=(col == 14)); c.number_format = NUM
                c.border = Border(left=thin if col in (1, 14) else None, right=thin if col in (1, 14) else None)
        r = lastr + 2
    ws.column_dimensions['A'].width = 18
    for col in range(2, 15): ws.column_dimensions[L(col)].width = 8.5
    ws.freeze_panes = 'B3'
    ws.sheet_view.showGridLines = False

    # ---- Seasonal (Oct-Sep marketing year) helper + chart
    sc = 17
    ws.cell(3, sc, f'Monthly World Total by Marketing Year (Oct-Sep), mln lbs').font = font(bold=True)
    mys = [y for y in years if y + 1 <= last.year]  # MY y/y+1 starts Oct of y
    order = list(range(10, 13)) + list(range(1, 10))
    ws.cell(4, sc, 'Month').font = font(bold=True)
    for j, y in enumerate(mys):
        c = ws.cell(4, sc + 1 + j, f'{y}/{str(y+1)[2:]}'); c.font = font(bold=True); c.fill = grey
    for i, m in enumerate(order):
        rr = 5 + i
        ws.cell(rr, sc, MON[m - 1]).font = font()
        for j, y in enumerate(mys):
            cy = y if m >= 10 else y + 1
            if pd.Timestamp(cy, m, 1) <= last:
                c = ws.cell(rr, sc + 1 + j, f'={sumifs(name, "World Total", cy, m)}')
                c.number_format = NUM; c.font = font()
    ch = LineChart(); ch.title = f'Monthly US Imports of {name} (mln lbs.)'
    ch.y_axis.title = 'mln lbs'; ch.height = 9; ch.width = 18
    ch.add_data(Reference(ws, min_col=sc + 1, max_col=sc + len(mys), min_row=4, max_row=16), titles_from_data=True)
    ch.set_categories(Reference(ws, min_col=sc, min_row=5, max_row=16))
    ch.series[-1].graphicalProperties.line.width = 38000
    ch.y_axis.delete = False; ch.x_axis.delete = False
    ws.add_chart(ch, f'{L(sc)}18')
    ws.column_dimensions[L(sc)].width = 7

# ---- Summary
ss = wb.create_sheet('Summary', 0)
ss['A1'] = 'US Imports of UCO & Tallow - Annual Totals (mln lbs.)'; ss['A1'].font = Font(name=F, size=11, bold=True)
ss.cell(3, 1, 'Commodity').font = font(bold=True); ss.cell(3, 1).fill = grey
for j, y in enumerate(years):
    lab = f'{y} (Jan-{MON[last.month-1]})' if y == last.year and last.month < 12 else str(y)
    c = ss.cell(3, j + 2, lab); c.font = font(bold=True); c.fill = grey; c.alignment = Alignment(horizontal='center')
for i, (hs, (name, _)) in enumerate(COM.items()):
    ss.cell(4 + i, 1, f'{name} ({hs})').font = font()
    for j, y in enumerate(years):
        c = ss.cell(4 + i, j + 2, f'={summ[(name, y)]}'); c.number_format = NUM; c.font = font()
notes = [
 'Notes',
 '- Source: U.S. Census Bureau, USA Trade Online (two files exported Oct 7, 2026). Data through ' + last.strftime('%b %Y') + '.',
 '- Volumes are General Imports (Quantity 1 Gen) in kg, converted at 2.20462262 lbs/kg (Data!B1) and shown in million lbs.',
 '- UCO = HTS 1518.00.4000 (animal & vegetable fats/oils, chemically modified), the statistical line under which used cooking oil enters the US.',
 '  It can also include other chemically modified fats/oils, so treat it as a close UCO proxy rather than pure UCO.',
 '- HS 150210 (6-digit parent) carried no quantities in the export and is excluded; its 10-digit lines (edible/inedible) are used instead.',
 '- Regional aggregates (Asia, Europe, etc.) are excluded; country rows sum exactly to the Census world total.',
 '- Australia / New Zealand include their territories as reported by Census.',
]
for i, t in enumerate(notes):
    ss.cell(9 + i, 1, t).font = font(bold=(i == 0))
ss.column_dimensions['A'].width = 30
for j in range(len(years)): ss.column_dimensions[L(j + 2)].width = 14
wb.move_sheet('Data', offset=len(wb.sheetnames))
wb.save(out)
