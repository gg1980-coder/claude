"""Build US soybean export sales by-destination workbook (current MY vs last year vs 5-yr avg).
Usage: python build_export_sales.py <ESR csv> <out.xlsx>
Source: USDA FAS Export Sales Reporting (ESR) query export, commodity SOYBEANS (801)."""
import sys, re
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.chart import LineChart, BarChart, Reference
from openpyxl.utils import get_column_letter as L
from openpyxl.formatting.rule import CellIsRule

src, out = sys.argv[1], sys.argv[2]
df = pd.read_csv(src, thousands=',')
df.columns = [c.strip() for c in df.columns]
df['Country'] = df['CountryName'].str.strip()
df['date'] = pd.to_datetime(df['WeekendingDate'])
df['MYstart'] = df['ReportingMarketingYearName'].str.extract(r'Sep (\d{4})').astype(int)
df['MY'] = df['MYstart'].astype(str) + '/' + (df['MYstart'] + 1).astype(str).str[-2:]
# in-season rows only: week-ending date between Aug 31 of start year and Sep 7 of next year
lo = pd.to_datetime(df['MYstart'].astype(str) + '-08-31')
hi = pd.to_datetime((df['MYstart'] + 1).astype(str) + '-09-07')
ins = df[(df.date >= lo) & (df.date <= hi)].copy()

cur_start = int(ins.loc[ins.date.idxmax(), 'MYstart'])
CUR = f"{cur_start}/{str(cur_start+1)[-2:]}"
WK = int(ins[ins.MYstart == cur_start].WeekNumber.max())
CUR_DATE = ins[ins.MYstart == cur_start].date.max()
mys = sorted(m for m in ins.MY.unique() if int(m[:4]) <= cur_start and int(m[:4]) >= cur_start - 6)
PRIOR = [m for m in mys if m != CUR][::-1]          # newest first
LY, FIVE = PRIOR[0], PRIOR[:5]
assert len(FIVE) == 5, FIVE

REGION = {
 'CHINA, PEOPLES REPUBLIC OF':'China','UNKNOWN':'Unknown',
 'JAPAN':'East Asia ex-China','TAIWAN':'East Asia ex-China','KOREA, REPUBLIC OF':'East Asia ex-China','HONG KONG':'East Asia ex-China',
 'INDONESIA':'Southeast Asia','VIETNAM':'Southeast Asia','MALAYSIA':'Southeast Asia','THAILAND':'Southeast Asia','PHILIPPINES':'Southeast Asia',
 'BURMA':'Southeast Asia','CAMBODIA':'Southeast Asia','LAOS':'Southeast Asia','SINGAPORE':'Southeast Asia',
 'BANGLADESH':'South Asia','PAKISTAN':'South Asia','NEPAL':'South Asia','SRI LANKA':'South Asia',
 'EGYPT':'Middle East & N. Africa','ALGERIA':'Middle East & N. Africa','TURKEY':'Middle East & N. Africa','TUNISIA':'Middle East & N. Africa',
 'MOROCCO':'Middle East & N. Africa','IRAN':'Middle East & N. Africa','IRAQ':'Middle East & N. Africa','ISRAEL':'Middle East & N. Africa',
 'JORDAN':'Middle East & N. Africa','LEBANON':'Middle East & N. Africa','LIBYA':'Middle East & N. Africa','SAUDI ARABIA':'Middle East & N. Africa',
 'QATAR':'Middle East & N. Africa',
 'SPAIN':'Europe','NETHERLANDS':'Europe','PORTUGAL':'Europe','ITALY':'Europe','GERMANY':'Europe','UNITED KINGDOM':'Europe','GREECE':'Europe',
 'BELGIUM-LUXEMBOURG':'Europe','FRANCE':'Europe','IRELAND':'Europe','NORWAY':'Europe','ROMANIA':'Europe','SWITZERLAND':'Europe',
 'UKRAINE':'Europe','GEORGIA, REPUBLIC OF':'Europe',
 'MEXICO':'Americas','CANADA':'Americas','COLOMBIA':'Americas','COSTA RICA':'Americas','VENEZUELA':'Americas','PERU':'Americas',
 'ARGENTINA':'Americas','BRAZIL':'Americas','CHILE':'Americas','CUBA':'Americas','DOMINICAN REPUBLIC':'Americas','EL SALVADOR':'Americas',
 'GUATEMALA':'Americas','HONDURAS':'Americas','NICARAGUA':'Americas','PANAMA':'Americas','BARBADOS':'Americas',
 'LEEWARD-WINDWARD ISLANDS':'Americas','VIRGIN ISLANDS OF THE U.S.':'Americas',
 'GHANA':'Sub-Saharan Africa','NIGER':'Sub-Saharan Africa','NIGERIA':'Sub-Saharan Africa','SENEGAL':'Sub-Saharan Africa',
 'SOUTH AFRICA, REPUBLIC OF':'Sub-Saharan Africa','AUSTRALIA':'Oceania',
}
REGIONS = ['China','Unknown','East Asia ex-China','Southeast Asia','South Asia','Middle East & N. Africa','Europe','Americas','Sub-Saharan Africa','Oceania']
miss = set(ins.Country) - set(REGION)
for c in miss: REGION[c] = 'Other'
if miss: REGIONS.append('Other'); print('Unmapped -> Other:', miss)

ins = ins[ins.MY.isin(mys)]
snap = ins[ins.WeekNumber == WK]
final = ins[ins.MY != CUR].sort_values('date').groupby(['MY','Country']).tail(1)  # last in-season report
final = final.set_index(['MY','Country']).AccumulatedExports

# ---------------- styles ----------------
F = 'Arial'
HFILL = PatternFill('solid', fgColor='1F4E78'); SUBFILL = PatternFill('solid', fgColor='D9E1F2')
CURFILL = PatternFill('solid', fgColor='FFF2CC'); TOTFILL = PatternFill('solid', fgColor='EDEDED')
thin = Side(style='thin', color='BFBFBF'); BOX = Border(top=thin, bottom=thin, left=thin, right=thin)
MT = '#,##0;-#,##0;"-"'; PCT = '+0.0%;-0.0%;0.0%'; SH = '0.0%'
BLUE = Font(name=F, size=10, color='0000FF'); GREEN = Font(name=F, size=10, color='008000')
NORM = Font(name=F, size=10); BOLD = Font(name=F, size=10, bold=True)

def title(ws, t, note, width=None):
    ws['A1'] = t; ws['A1'].font = Font(name=F, size=12, bold=True)
    ws['A2'] = note; ws['A2'].font = Font(name=F, size=9, italic=True, color='595959')
def header(ws, row, labels, col=1):
    for i, h in enumerate(labels):
        c = ws.cell(row, col + i, h); c.font = Font(name=F, size=10, bold=True, color='FFFFFF')
        c.fill = HFILL; c.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True); c.border = BOX
    ws.row_dimensions[row].height = 42
def put(ws, r, c, v, fmt=None, font=NORM, fill=None):
    x = ws.cell(r, c, v); x.font = font; x.border = BOX
    if fmt: x.number_format = fmt
    if fill: x.fill = fill
    return x
def dpct(cur, base): return f'=IF({base}=0,IF({cur}=0,0,"new"),{cur}/{base}-1)'

wb = Workbook()
wsS = wb.active; wsS.title = 'Summary'
wsR = wb.create_sheet('Ranking by Destination')
wsM = wb.create_sheet('Top Movers')
wsG = wb.create_sheet('Regions')
wsW = wb.create_sheet('Weekly Commitments')
wsD = wb.create_sheet('Data')

# ---------------- Data (long, hardcoded source) ----------------
title(wsD, f'Source data — USDA FAS Export Sales, SOYBEANS, in-season week {WK} snapshot per marketing year + final MY exports (MT)',
      'Hardcoded from the ESR query export (blue). Rows are the in-season report only (pre-season "week 9" rows of the prior reporting year excluded). '
      'Final MY exports = Accumulated Exports on the last in-season report of that MY (blank for current MY).')
header(wsD, 4, ['Marketing year','Week #','Week ending','Destination','Region','Total commitment','Accumulated exports','Outstanding sales','Final MY exports'])
r = 5
for _, x in snap.sort_values(['MY','Country']).iterrows():
    fe = final.get((x.MY, x.Country), None) if x.MY != CUR else None
    vals = [x.MY, int(x.WeekNumber), x.date.to_pydatetime(), x.Country, None,
            float(x.TotalCommitment), float(x.AccumulatedExports), float(x.OutstandingSales), fe]
    for j, v in enumerate(vals, 1):
        c = put(wsD, r, j, v, font=BLUE)
    wsD.cell(r, 5).value = f"=IFERROR(INDEX('Ranking by Destination'!$C:$C,MATCH(D{r},'Ranking by Destination'!$B:$B,0)),\"Other\")"
    wsD.cell(r, 5).font = GREEN
    wsD.cell(r, 3).number_format = 'dd-mmm-yy'
    for j in (6, 7, 8, 9): wsD.cell(r, j).number_format = MT
    r += 1
# destinations that appear only in final-year data but not in snapshot: add final rows too
extra = [(m, c) for (m, c) in final.index if not ((snap.MY == m) & (snap.Country == c)).any()]
for m, c in extra:
    vals = [m, None, None, c, None, 0, 0, 0, float(final[(m, c)])]
    for j, v in enumerate(vals, 1): put(wsD, r, j, v, font=BLUE)
    wsD.cell(r, 5).value = f"=IFERROR(INDEX('Ranking by Destination'!$C:$C,MATCH(D{r},'Ranking by Destination'!$B:$B,0)),\"Other\")"
    wsD.cell(r, 5).font = GREEN
    for j in (6, 7, 8, 9): wsD.cell(r, j).number_format = MT
    r += 1
DLAST = r - 1
put(wsD, r + 1, 1, f'Note: {len(extra)} rows added with zero week-{WK} values for destinations that shipped later in a MY but had no commitment at week {WK} (needed for full-MY exports).',
    font=Font(name=F, size=9, italic=True)).border = Border()
for col, w in zip('ABCDEFGHI', [14, 8, 12, 30, 22, 16, 16, 16, 16]): wsD.column_dimensions[col].width = w
wsD.freeze_panes = 'A5'; wsD.auto_filter.ref = f'A4:I{DLAST}'
def rng(col): return f"Data!${col}$5:${col}${DLAST}"
def sumifs(valcol, country_ref, my_ref):
    return f"SUMIFS({rng(valcol)},{rng('D')},{country_ref},{rng('A')},{my_ref})"

# ---------------- Ranking by Destination ----------------
countries = sorted(set(ins.Country))
sc = snap.pivot_table(index='Country', columns='MY', values='TotalCommitment', aggfunc='sum').reindex(countries).fillna(0)
avg5 = sc[FIVE].mean(axis=1); dev = sc[CUR] - avg5
order = dev.sort_values(ascending=False).index.tolist()
order = [c for c in order if (sc.loc[c].abs().sum() > 0) or c in final.index.get_level_values(1)]

cols = ['Rank (Δ vs 5-yr)','Destination','Region', CUR] + PRIOR + [
    '5-yr avg', '5-yr min', '5-yr max', f'Δ vs {LY} MT', f'Δ vs {LY} %', 'Δ vs 5-yr avg MT', 'Δ vs 5-yr avg %',
    f'Rank (Δ vs {LY})', 'vs 5-yr range', f'Exports shipped {CUR}', 'Exports shipped 5-yr avg', 'Δ shipped vs 5-yr MT',
    f'Unshipped {CUR}', f'Share of {CUR} total', f'Full-MY exports {LY}', f'Pace: {CUR} commit as % of {LY} full-MY']
title(wsR, f'US Soybean Export Sales by Destination — total commitment at week {WK} ({CUR_DATE:%d %b %Y}), {CUR} vs {LY} vs 5-yr avg ({FIVE[-1]}–{FIVE[0]})',
      f'Total commitment = accumulated exports + outstanding sales for the MY, as of in-season week {WK} of each MY (same point in the season). '
      'Sorted by Δ vs 5-yr average (largest gain first). Ranks are live formulas. Region grouping is an analyst assumption (blue cells, editable). '
      'UNKNOWN = sales to undeclared destinations — historically mostly China.')
header(wsR, 4, cols)
wsR.row_dimensions[4].height = 54
yc = {m: 4 + i for i, m in enumerate([CUR] + PRIOR)}   # column index of each MY
cA = 4 + len(PRIOR) + 1  # 5-yr avg col
names = {k: i + 1 for i, k in enumerate(cols)}
first, last = 5, 4 + len(order)
for i, c in enumerate(order):
    r = first + i
    fill = TOTFILL if c in ('CHINA, PEOPLES REPUBLIC OF', 'UNKNOWN') else None
    put(wsR, r, 2, c, font=BOLD if fill else NORM, fill=fill)
    put(wsR, r, 3, REGION[c], font=BLUE, fill=fill)
    for m, j in yc.items():
        put(wsR, r, j, '=' + sumifs('F', f'$B{r}', f'{L(j)}$4'), MT, fill=CURFILL if m == CUR else fill)
    five = f'{L(yc[FIVE[0]])}{r}:{L(yc[FIVE[-1]])}{r}'
    cur, ly = f'{L(yc[CUR])}{r}', f'{L(yc[LY])}{r}'
    a = f'{L(cA)}{r}'
    put(wsR, r, cA, f'=AVERAGE({five})', MT, fill=fill)
    put(wsR, r, cA + 1, f'=MIN({five})', MT, fill=fill)
    put(wsR, r, cA + 2, f'=MAX({five})', MT, fill=fill)
    put(wsR, r, cA + 3, f'={cur}-{ly}', MT, fill=fill)
    put(wsR, r, cA + 4, dpct(cur, ly), PCT, fill=fill)
    put(wsR, r, cA + 5, f'={cur}-{a}', MT, font=BOLD, fill=fill)
    put(wsR, r, cA + 6, dpct(cur, a), PCT, fill=fill)
    dl = L(cA + 3); d5 = L(cA + 5)
    put(wsR, r, 1, f'=RANK({d5}{r},${d5}${first}:${d5}${last},0)', '0', fill=fill)
    put(wsR, r, cA + 7, f'=RANK({dl}{r},${dl}${first}:${dl}${last},0)', '0', fill=fill)
    put(wsR, r, cA + 8, f'=IF({cur}>{L(cA+2)}{r},"ABOVE 5-yr max",IF({cur}<{L(cA+1)}{r},"below 5-yr min","within range"))', fill=fill)
    j = cA + 9
    put(wsR, r, j, '=' + sumifs('G', f'$B{r}', f'{L(yc[CUR])}$4'), MT, fill=fill)
    put(wsR, r, j + 1, '=(' + '+'.join(sumifs('G', f'$B{r}', f'{L(yc[m])}$4') for m in FIVE) + ')/5', MT, fill=fill)
    put(wsR, r, j + 2, f'={L(j)}{r}-{L(j+1)}{r}', MT, fill=fill)
    put(wsR, r, j + 3, '=' + sumifs('H', f'$B{r}', f'{L(yc[CUR])}$4'), MT, fill=fill)
    put(wsR, r, j + 4, f'=IF({L(yc[CUR])}${last+1}=0,0,{cur}/{L(yc[CUR])}${last+1})', SH, fill=fill)
    put(wsR, r, j + 5, '=' + sumifs('I', f'$B{r}', f'{L(yc[LY])}$4'), MT, fill=fill)
    put(wsR, r, j + 6, f'=IF({L(j+5)}{r}=0,IF({cur}=0,0,"new"),{cur}/{L(j+5)}{r})', '0%', fill=fill)
T = last + 1
put(wsR, T, 2, 'TOTAL (all destinations)', font=BOLD, fill=TOTFILL)
put(wsR, T, 3, '', fill=TOTFILL)
for j in list(yc.values()) + [cA, cA + 3, cA + 5, cA + 9, cA + 10, cA + 11, cA + 12, cA + 14]:
    put(wsR, T, j, f'=SUM({L(j)}{first}:{L(j)}{last})', MT, font=BOLD, fill=TOTFILL)
five = f'{L(yc[FIVE[0]])}{T}:{L(yc[FIVE[-1]])}{T}'
put(wsR, T, cA + 1, f'=MIN({five})', MT, font=BOLD, fill=TOTFILL); put(wsR, T, cA + 2, f'=MAX({five})', MT, font=BOLD, fill=TOTFILL)
put(wsR, T, cA + 4, dpct(f'{L(yc[CUR])}{T}', f'{L(yc[LY])}{T}'), PCT, font=BOLD, fill=TOTFILL)
put(wsR, T, cA + 6, dpct(f'{L(yc[CUR])}{T}', f'{L(cA)}{T}'), PCT, font=BOLD, fill=TOTFILL)
put(wsR, T, cA + 13, f'=SUM({L(cA+13)}{first}:{L(cA+13)}{last})', SH, font=BOLD, fill=TOTFILL)
put(wsR, T, cA + 15, dpct(f'{L(yc[CUR])}{T}', f'{L(cA+14)}{T}').replace('-1)', ')'), '0%', font=BOLD, fill=TOTFILL)
# check row: total must equal ESR "Total" column
tot_chk = {m: float(snap[snap.MY == m].TotalTotalCommitment.iloc[0]) for m in mys}
put(wsR, T + 1, 2, 'ESR reported total (check)', font=Font(name=F, size=9, italic=True))
put(wsR, T + 2, 2, 'Difference (must be 0)', font=Font(name=F, size=9, italic=True))
for m, j in yc.items():
    put(wsR, T + 1, j, tot_chk[m], MT, font=BLUE)
    put(wsR, T + 2, j, f'={L(j)}{T}-{L(j)}{T+1}', MT)
for jj, k in enumerate(['A'], 0): pass
for col in range(1, len(cols) + 1): wsR.column_dimensions[L(col)].width = 13
wsR.column_dimensions['A'].width = 8; wsR.column_dimensions['B'].width = 30; wsR.column_dimensions['C'].width = 22
wsR.column_dimensions[L(cA + 8)].width = 16
wsR.freeze_panes = 'D5'
gr = PatternFill('solid', fgColor='C6EFCE'); rd = PatternFill('solid', fgColor='FFC7CE')
for j in (cA + 3, cA + 5):
    rr = f'{L(j)}{first}:{L(j)}{last}'
    wsR.conditional_formatting.add(rr, CellIsRule(operator='greaterThan', formula=['0'], fill=gr))
    wsR.conditional_formatting.add(rr, CellIsRule(operator='lessThan', formula=['0'], fill=rd))
RC = {k: L(v) for k, v in names.items()}

# ---------------- Summary ----------------
title(wsS, f'US Soybean Export Sales — {CUR} vs {LY} vs 5-yr avg, at in-season week {WK} (w/e {CUR_DATE:%d %b %Y})',
      'All figures link to "Ranking by Destination" (SUMIF by region). RoW = all destinations excluding China AND Unknown. '
      'Unknown is shown separately because it is a booking placeholder that is usually re-declared to China later in the season.')
def block(r0, label, srcol_cur, srcols_prior, src_avg_fn, fmt_note):
    put(wsS, r0, 1, label, font=Font(name=F, size=11, bold=True)).border = Border()
    header(wsS, r0 + 1, ['Group', CUR, LY, '5-yr avg', f'Δ vs {LY} MT', f'Δ vs {LY} %', 'Δ vs 5-yr MT', 'Δ vs 5-yr %', f'Share {CUR}', 'Share 5-yr avg', f'4-yr avg excl. {LY}', 'Δ vs 4-yr MT', 'Δ vs 4-yr %'])
    groups = [('China', 'China'), ('Unknown', 'Unknown'), ('RoW excl. China & Unknown', None), ('Total', 'ALL')]
    rows = {}
    for i, (g, key) in enumerate(groups):
        r = r0 + 2 + i; rows[g] = r
        fill = TOTFILL if g == 'Total' else None
        put(wsS, r, 1, g, font=BOLD if fill else NORM, fill=fill)
        for j, colL in ((2, srcol_cur), (3, srcols_prior), (4, src_avg_fn)):
            if key == 'ALL': f = f"=SUM('Ranking by Destination'!{colL}{first}:{colL}{last})"
            elif key: f = f"=SUMIF('Ranking by Destination'!$C${first}:$C${last},\"{key}\",'Ranking by Destination'!{colL}{first}:{colL}{last})"
            else: f = f'={L(j)}{r0+5}-{L(j)}{r0+2}-{L(j)}{r0+3}'
            put(wsS, r, j, f, MT, font=GREEN if key else NORM, fill=fill)
        put(wsS, r, 5, f'=B{r}-C{r}', MT, fill=fill); put(wsS, r, 6, dpct(f'B{r}', f'C{r}'), PCT, fill=fill)
        put(wsS, r, 7, f'=B{r}-D{r}', MT, font=BOLD, fill=fill); put(wsS, r, 8, dpct(f'B{r}', f'D{r}'), PCT, fill=fill)
        put(wsS, r, 9, f'=IF(B${r0+5}=0,0,B{r}/B${r0+5})', SH, fill=fill); put(wsS, r, 10, f'=IF(D${r0+5}=0,0,D{r}/D${r0+5})', SH, fill=fill)
        put(wsS, r, 11, f'=(D{r}*5-C{r})/4', MT, fill=fill)
        put(wsS, r, 12, f'=B{r}-K{r}', MT, font=BOLD, fill=fill); put(wsS, r, 13, dpct(f'B{r}', f'K{r}'), PCT, fill=fill)
    return rows
b1 = block(4, f'1) Total commitment (shipped + unshipped) at week {WK}', RC[CUR], RC[LY], RC['5-yr avg'], '')
b2 = block(11, f'2) Accumulated exports (already shipped) at week {WK}', RC[f'Exports shipped {CUR}'],
           None, RC['Exports shipped 5-yr avg'], '')
# exports LY column: need LY exports by destination; compute from Data directly
for g, r in b2.items():
    key = {'China': 'CHINA, PEOPLES REPUBLIC OF', 'Unknown': 'UNKNOWN'}.get(g)
    if g == 'Total': f = f"=SUMIFS({rng('G')},{rng('A')},\"{LY}\",{rng('B')},{WK})"
    elif key: f = f"=SUMIFS({rng('G')},{rng('A')},\"{LY}\",{rng('D')},\"{key}\")"
    else: f = f'=C{b2["Total"]}-C{b2["China"]}-C{b2["Unknown"]}'
    wsS.cell(r, 3).value = f
# regional block
r0 = 18
put(wsS, r0, 1, f'3) Total commitment by region at week {WK} (ranked by Δ vs 5-yr avg on the Regions tab)', font=Font(name=F, size=11, bold=True)).border = Border()
put(wsS, r0 + 1, 1, '→ see "Regions" and "Top Movers" tabs', font=Font(name=F, size=10, italic=True)).border = Border()
# pace block
r0 = 21
put(wsS, r0, 1, f'4) Seasonal pace — week-{WK} commitment as % of the final MY exports (how much of the year is usually booked by now)', font=Font(name=F, size=11, bold=True)).border = Border()
header(wsS, r0 + 1, ['Marketing year', f'Commitment wk {WK}', 'Final MY exports', 'Booked by wk %'])
pr = r0 + 2
for i, m in enumerate(FIVE[::-1]):
    r = pr + i
    put(wsS, r, 1, m)
    put(wsS, r, 2, f"=SUMIFS({rng('F')},{rng('A')},A{r})", MT, font=GREEN)
    put(wsS, r, 3, f"=SUMIFS({rng('I')},{rng('A')},A{r})", MT, font=GREEN)
    put(wsS, r, 4, f'=IF(C{r}=0,0,B{r}/C{r})', SH)
re_ = pr + 5
put(wsS, re_, 1, '5-yr average pace', font=BOLD, fill=TOTFILL); put(wsS, re_, 2, '', fill=TOTFILL); put(wsS, re_, 3, '', fill=TOTFILL)
put(wsS, re_, 4, f'=AVERAGE(D{pr}:D{pr+4})', SH, font=BOLD, fill=TOTFILL)
put(wsS, re_ + 1, 1, f'{CUR} commitment', font=NORM); put(wsS, re_ + 1, 2, f'=B{b1["Total"]}', MT)
put(wsS, re_ + 2, 1, f'Implied {CUR} exports at avg pace', font=BOLD); put(wsS, re_ + 2, 2, f'=IF(D{re_}=0,0,B{re_+1}/D{re_})', MT, font=BOLD)
put(wsS, re_ + 3, 1, 'Range (fastest / slowest pace yr)'); put(wsS, re_ + 3, 2, f'=B{re_+1}/MAX(D{pr}:D{pr+4})', MT); put(wsS, re_ + 3, 3, f'=B{re_+1}/MIN(D{pr}:D{pr+4})', MT)
put(wsS, re_ + 4, 1, 'Caveat: a mechanical extrapolation, not a forecast. Booking pace depends on the price spread vs Brazil and on Chinese buying timing; '
     f'{LY} pace was distorted by China\'s absence from US sales.', font=Font(name=F, size=9, italic=True)).border = Border()
notes = re_ + 6
for k, t in enumerate([
    'Reading notes:',
    f'• "This year" = MY {CUR} (Sep–Aug). Every MY is compared at the same in-season week number ({WK}); report dates differ by up to ±4 days across years.',
    f'• 5-yr avg = simple average of {FIVE[-1]} … {FIVE[0]}. It includes {LY}, when China bought almost nothing from the US early in the season, which pulls the China average down.',
    f'• Columns K–M: 4-yr avg excluding {LY} (the anomalous China-absent year) — the fairer "normal year" benchmark for China and RoW.',
    '• Total commitment captures purchase intent (bookings); accumulated exports capture physical flow. Early in the MY, commitment is the better demand signal.',
    '• Outstanding sales can be cancelled or rolled; UNKNOWN is typically re-allocated to named destinations (mostly China) when shipped.',
    f'• Source: USDA FAS Export Sales Reporting (ESR) query, SOYBEANS (801), data through w/e {CUR_DATE:%d %b %Y}.']):
    put(wsS, notes + k, 1, t, font=Font(name=F, size=9, bold=(k == 0), italic=(k > 0))).border = Border()
wsS.column_dimensions['A'].width = 34
for col in 'BCDEFGHIJKLM': wsS.column_dimensions[col].width = 14

# ---------------- Regions ----------------
title(wsG, f'Total commitment by region at week {WK} — {CUR} vs {LY} vs 5-yr avg (MT)',
      'SUMIF on "Ranking by Destination" region column. Sorted by Δ vs 5-yr avg at build time; numbers are live.')
header(wsG, 4, ['Region', '# destinations', CUR, LY, '5-yr avg', f'Δ vs {LY} MT', f'Δ vs {LY} %', 'Δ vs 5-yr MT', 'Δ vs 5-yr %', f'Share {CUR}', 'Share 5-yr avg'])
reg_dev = {}
for g in REGIONS:
    cs = [c for c in order if REGION[c] == g]
    reg_dev[g] = sum(dev.get(c, 0) for c in cs)
regs = sorted([g for g in REGIONS if any(REGION[c] == g for c in order)], key=lambda g: -reg_dev[g])
RR = f"'Ranking by Destination'!$C${first}:$C${last}"
for i, g in enumerate(regs):
    r = 5 + i
    put(wsG, r, 1, g, font=BOLD)
    put(wsG, r, 2, f'=COUNTIF({RR},A{r})', '0')
    for j, k in ((3, CUR), (4, LY), (5, '5-yr avg')):
        put(wsG, r, j, f"=SUMIF({RR},$A{r},'Ranking by Destination'!{RC[k]}${first}:{RC[k]}${last})", MT, font=GREEN)
    put(wsG, r, 6, f'=C{r}-D{r}', MT); put(wsG, r, 7, dpct(f'C{r}', f'D{r}'), PCT)
    put(wsG, r, 8, f'=C{r}-E{r}', MT, font=BOLD); put(wsG, r, 9, dpct(f'C{r}', f'E{r}'), PCT)
    GT = 5 + len(regs)
    put(wsG, r, 10, f'=IF(C${GT}=0,0,C{r}/C${GT})', SH); put(wsG, r, 11, f'=IF(E${GT}=0,0,E{r}/E${GT})', SH)
GT = 5 + len(regs)
put(wsG, GT, 1, 'TOTAL', font=BOLD, fill=TOTFILL)
for j in range(2, 7): put(wsG, GT, j, f'=SUM({L(j)}5:{L(j)}{GT-1})', MT if j > 2 else '0', font=BOLD, fill=TOTFILL)
put(wsG, GT, 7, dpct(f'C{GT}', f'D{GT}'), PCT, font=BOLD, fill=TOTFILL)
put(wsG, GT, 8, f'=SUM(H5:H{GT-1})', MT, font=BOLD, fill=TOTFILL); put(wsG, GT, 9, dpct(f'C{GT}', f'E{GT}'), PCT, font=BOLD, fill=TOTFILL)
put(wsG, GT, 10, f'=SUM(J5:J{GT-1})', SH, font=BOLD, fill=TOTFILL); put(wsG, GT, 11, f'=SUM(K5:K{GT-1})', SH, font=BOLD, fill=TOTFILL)
wsG.column_dimensions['A'].width = 26
for col in 'BCDEFGHIJK': wsG.column_dimensions[col].width = 14
for j in (6, 8):
    rr = f'{L(j)}5:{L(j)}{GT-1}'
    wsG.conditional_formatting.add(rr, CellIsRule(operator='greaterThan', formula=['0'], fill=gr))
    wsG.conditional_formatting.add(rr, CellIsRule(operator='lessThan', formula=['0'], fill=rd))
ch = BarChart(); ch.type = 'bar'; ch.title = f'Δ commitment vs 5-yr avg by region (MT), wk {WK}'
ch.add_data(Reference(wsG, min_col=8, min_row=4, max_row=GT - 1), titles_from_data=True)
ch.set_categories(Reference(wsG, min_col=1, min_row=5, max_row=GT - 1)); ch.legend = None
ch.y_axis.numFmt = '#,##0'; ch.width = 18; ch.height = 8; ch.x_axis.scaling.orientation = 'maxMin'
ch.x_axis.delete = False; ch.y_axis.delete = False; ch.x_axis.tickLblPos = 'low'
wsG.add_chart(ch, f'A{GT+3}')

# ---------------- Top Movers ----------------
N = 15
ex = [c for c in order if c not in ('CHINA, PEOPLES REPUBLIC OF', 'UNKNOWN')]
dly = (sc[CUR] - sc[LY])
title(wsM, f'Top {N} gainers and losers by destination (excl. China & Unknown) — total commitment at week {WK}, MT',
      'Lists sorted at build time; numbers are INDEX/MATCH links to "Ranking by Destination" and update live. China & Unknown are on the Summary tab. '
      'Caution: many buyers take 1–3 cargoes in a season, so single-country swings of ~60–70k t can be one vessel booked earlier or later.')
def mover_table(r0, c0, label, names_, dcol, pcol):
    put(wsM, r0, c0, label, font=Font(name=F, size=11, bold=True)).border = Border()
    header(wsM, r0 + 1, ['#', 'Destination', 'Region', CUR, LY, '5-yr avg', 'Δ MT', 'Δ %', 'vs 5-yr range'], c0)
    for i, c in enumerate(names_):
        r = r0 + 2 + i
        put(wsM, r, c0, i + 1, '0'); put(wsM, r, c0 + 1, c)
        nm = f'{L(c0+1)}{r}'
        for k, (hdr, fmt) in enumerate([('Region', None), (CUR, MT), (LY, MT), ('5-yr avg', MT), (dcol, MT), (pcol, PCT), ('vs 5-yr range', None)]):
            colL = RC[hdr]
            put(wsM, r, c0 + 2 + k, f"=INDEX('Ranking by Destination'!${colL}:${colL},MATCH({nm},'Ranking by Destination'!$B:$B,0))", fmt, font=GREEN)
g5 = ex[:N]; l5 = [c for c in ex[::-1] if dev[c] < 0][:N]
gl = [c for c in dly.drop(['CHINA, PEOPLES REPUBLIC OF', 'UNKNOWN']).sort_values(ascending=False).index if dly[c] > 0][:N]
ll = [c for c in dly.drop(['CHINA, PEOPLES REPUBLIC OF', 'UNKNOWN']).sort_values().index if dly[c] < 0][:N]
mover_table(4, 1, f'A) Taking MORE vs 5-yr avg', g5, 'Δ vs 5-yr avg MT', 'Δ vs 5-yr avg %')
mover_table(4, 11, f'B) Taking LESS vs 5-yr avg', l5, 'Δ vs 5-yr avg MT', 'Δ vs 5-yr avg %')
mover_table(8 + N, 1, f'C) Taking MORE vs {LY}', gl, f'Δ vs {LY} MT', f'Δ vs {LY} %')
mover_table(8 + N, 11, f'D) Taking LESS vs {LY}', ll, f'Δ vs {LY} MT', f'Δ vs {LY} %')
for c0 in (1, 11):
    for k, w in enumerate([5, 28, 20, 12, 12, 12, 12, 10, 15]): wsM.column_dimensions[L(c0 + k)].width = w
wsM.column_dimensions['J'].width = 3
for c0 in (1, 11):
    for r0 in (6, 10 + N):
        rr = f'{L(c0+6)}{r0}:{L(c0+6)}{r0+N-1}'
        wsM.conditional_formatting.add(rr, CellIsRule(operator='greaterThan', formula=['0'], fill=gr))
        wsM.conditional_formatting.add(rr, CellIsRule(operator='lessThan', formula=['0'], fill=rd))

# ---------------- RoW vs Avg & Median ----------------
wsV = wb.create_sheet('RoW vs Avg & Median', index=3)
med5 = sc[FIVE].median(axis=1)
dA, dM = sc[CUR] - avg5, sc[CUR] - med5
def status(c):
    a, m = dA[c], dM[c]
    if a > 0 and m > 0: return 'AHEAD of avg & median'
    if a < 0 and m < 0: return 'BEHIND avg & median'
    if a > 0: return 'Ahead of avg only'
    if m > 0: return 'Ahead of median only'
    if a < 0: return 'Behind avg only'
    if m < 0: return 'Behind median only'
    return 'In line'
GROUPS = ['AHEAD of avg & median', 'Ahead of avg only', 'Ahead of median only', 'In line', 'Behind median only', 'Behind avg only', 'BEHIND avg & median']
fin5 = {c for (m, c), v in final.items() if m in FIVE and v > 0}
row_c = [c for c in order if c not in ('CHINA, PEOPLES REPUBLIC OF', 'UNKNOWN') and (sc.loc[c, [CUR] + FIVE].abs().sum() > 0 or c in fin5)]
title(wsV, f'RoW destinations (excl. China & Unknown) — total commitment at week {WK}: {CUR} vs 5-yr avg and 5-yr median ({FIVE[-1]}–{FIVE[0]}, incl. {LY})',
      'Year values link to "Ranking by Destination"; avg, median, Δ and Status are live formulas. Grouped by status at build time, then by Δ vs avg. '
      'Median keeps all 5 years but is not pulled by one extreme year. "Behind avg only" with a 0 median = occasional buyer with no booking yet. '
      '"≥ 1 cargo move?" flags |Δ| ≥ the cargo size in cell Q3 (editable; 65 kt ≈ one Panamax) — smaller moves are timing noise.')
wsV['P3'] = 'Cargo size (MT):'; wsV['P3'].font = BOLD; wsV['P3'].alignment = Alignment(horizontal='right')
wsV['Q3'] = 65000; wsV['Q3'].font = BLUE; wsV['Q3'].number_format = MT; wsV['Q3'].fill = PatternFill('solid', fgColor='FFFF00')
from openpyxl.workbook.defined_name import DefinedName
wb.defined_names['Cargo_MT'] = DefinedName('Cargo_MT', attr_text="'RoW vs Avg & Median'!$Q$3")
yrs = [CUR] + FIVE
hdrV = ['Status', 'Destination', 'Region'] + yrs + ['5-yr avg', '5-yr median', 'Δ vs avg MT', 'Δ vs avg %', 'Δ vs median MT', 'Δ vs median %', '≥ 1 cargo move?', '']
header(wsV, 4, hdrV)
cY = {m: 4 + i for i, m in enumerate(yrs)}; cAv = 4 + len(yrs); cMd = cAv + 1
GFILL = {'AHEAD of avg & median': 'C6EFCE', 'Ahead of avg only': 'EBF1DE', 'Ahead of median only': 'EBF1DE', 'In line': 'F2F2F2', 'Behind avg only': 'FDE9D9', 'Behind median only': 'FDE9D9', 'BEHIND avg & median': 'FFC7CE'}
cF = cMd + 7   # full-MY block (one spacer column after cargo flag)
FULLH = [f'Full-MY exports {m}' for m in FIVE] + ['Full-MY exports 5-yr avg', f'{LY}: wk {WK} commit as % of full year',
         f'5-yr avg: wk {WK} commit as % of full year', f'{CUR} wk {WK} commit as % of {LY} full year',
         f'{CUR} wk {WK} commit as % of 5-yr avg full year', f'Implied {CUR} full year at 5-yr avg pace']
def pct_or(n, d, alt='"n/a"'): return f'=IF({d}=0,IF({n}=0,0,{alt}),{n}/{d})'
def fullcols(r, src, fill=None):
    """src(m, col) -> formula for full-MY exports of MY m in row r."""
    for k, m in enumerate(FIVE):
        put(wsV, r, cF + k, src(m, cF + k), MT, font=GREEN if fill is None else BOLD, fill=fill)
    fa = f'{L(cF+5)}{r}'
    put(wsV, r, cF + 5, f'=AVERAGE({L(cF)}{r}:{L(cF+4)}{r})', MT, fill=fill)
    put(wsV, r, cF + 6, pct_or(f'{L(cY[LY])}{r}', f'{L(cF)}{r}'), '0%', fill=fill)
    put(wsV, r, cF + 7, pct_or(f'{L(cAv)}{r}', fa), '0%', fill=fill)
    put(wsV, r, cF + 8, pct_or(f'{L(cY[CUR])}{r}', f'{L(cF)}{r}', '"new"'), '0%', font=BOLD, fill=fill)
    put(wsV, r, cF + 9, pct_or(f'{L(cY[CUR])}{r}', fa, '"new"'), '0%', font=BOLD, fill=fill)
    put(wsV, r, cF + 10, f'=IF(N({L(cF+7)}{r})=0,"n/a",{L(cY[CUR])}{r}/{L(cF+7)}{r})', MT, fill=fill)
header(wsV, 4, FULLH, cF); wsV.row_dimensions[4].height = 66
wsV.cell(4, cMd + 6).fill = PatternFill(None); wsV.cell(4, cMd + 6).border = Border()
r = 5; vfirst = 5
for g in GROUPS:
    cs = sorted([c for c in row_c if status(c) == g], key=lambda c: -dA[c])
    if not cs: continue
    put(wsV, r, 1, f'{g} — {len(cs)} destinations', font=Font(name=F, size=10, bold=True), fill=PatternFill('solid', fgColor=GFILL[g]))
    for j in list(range(2, len(hdrV))) + list(range(cF, cF + len(FULLH))): put(wsV, r, j, None, fill=PatternFill('solid', fgColor=GFILL[g]))
    r += 1
    for c in cs:
        rr = first + order.index(c)
        put(wsV, r, 2, c); put(wsV, r, 3, f"='Ranking by Destination'!C{rr}", font=GREEN)
        for m in yrs:
            put(wsV, r, cY[m], f"='Ranking by Destination'!{L(yc[m])}{rr}", MT, font=GREEN, fill=CURFILL if m == CUR else None)
        five = f'{L(cY[FIVE[0]])}{r}:{L(cY[FIVE[-1]])}{r}'; cur = f'{L(cY[CUR])}{r}'
        put(wsV, r, cAv, f'=AVERAGE({five})', MT); put(wsV, r, cMd, f'=MEDIAN({five})', MT)
        put(wsV, r, cMd + 1, f'={cur}-{L(cAv)}{r}', MT, font=BOLD); put(wsV, r, cMd + 2, dpct(cur, f'{L(cAv)}{r}'), PCT)
        put(wsV, r, cMd + 3, f'={cur}-{L(cMd)}{r}', MT, font=BOLD); put(wsV, r, cMd + 4, dpct(cur, f'{L(cMd)}{r}'), PCT)
        a, m_ = f'{L(cMd+1)}{r}', f'{L(cMd+3)}{r}'
        put(wsV, r, 1, f'=IF(AND({a}>0,{m_}>0),"AHEAD of avg & median",IF(AND({a}<0,{m_}<0),"BEHIND avg & median",IF({a}>0,"Ahead of avg only",IF({m_}>0,"Ahead of median only",IF({a}<0,"Behind avg only",IF({m_}<0,"Behind median only","In line"))))))')
        put(wsV, r, cMd + 5, f'=IF(MAX(ABS({a}),ABS({m_}))>=Cargo_MT,"yes","")')
        fullcols(r, lambda m, col, r=r: '=' + sumifs('I', f'$B{r}', f'"{m}"'))
        r += 1
vlast = r - 1
# totals by status
r += 1
put(wsV, r, 1, 'Totals by status', font=Font(name=F, size=11, bold=True)).border = Border(); r += 1
header(wsV, r, ['Status', '# destinations', ''] + yrs + ['5-yr avg', '5-yr median*', 'Δ vs avg MT', 'Δ vs avg %', 'Δ vs median MT', 'Δ vs median %']); header(wsV, r, FULLH, cF); wsV.row_dimensions[r].height = 66; r += 1
SR = f'$A${vfirst}:$A${vlast}'
tfirst = r
for g in GROUPS + ['RoW TOTAL']:
    fill = TOTFILL if g == 'RoW TOTAL' else None
    put(wsV, r, 1, g, font=BOLD, fill=fill)
    put(wsV, r, 2, f'=COUNTIF({SR},A{r})' if fill is None else f'=SUM(B{tfirst}:B{r-1})', '0', fill=fill); put(wsV, r, 3, None, fill=fill)
    for m in yrs:
        cl = L(cY[m])
        f = f'=SUMIF({SR},$A{r},{cl}${vfirst}:{cl}${vlast})' if fill is None else f'=SUM({cl}{tfirst}:{cl}{r-1})'
        put(wsV, r, cY[m], f, MT, fill=fill)
    five = f'{L(cY[FIVE[0]])}{r}:{L(cY[FIVE[-1]])}{r}'; cur = f'{L(cY[CUR])}{r}'
    put(wsV, r, cAv, f'=AVERAGE({five})', MT, fill=fill); put(wsV, r, cMd, f'=MEDIAN({five})', MT, fill=fill)
    put(wsV, r, cMd + 1, f'={cur}-{L(cAv)}{r}', MT, font=BOLD, fill=fill); put(wsV, r, cMd + 2, dpct(cur, f'{L(cAv)}{r}'), PCT, fill=fill)
    put(wsV, r, cMd + 3, f'={cur}-{L(cMd)}{r}', MT, font=BOLD, fill=fill); put(wsV, r, cMd + 4, dpct(cur, f'{L(cMd)}{r}'), PCT, fill=fill)
    if fill is None:
        fullcols(r, lambda m, col, r=r: f'=SUMIF({SR},$A{r},{L(col)}${vfirst}:{L(col)}${vlast})')
    else:
        fullcols(r, lambda m, col, r=r: f'=SUM({L(col)}{tfirst}:{L(col)}{r-1})', fill)
    r += 1
put(wsV, r, 1, '* Group median = median of the group\'s yearly totals, not the sum of country medians. RoW TOTAL ties to Summary (RoW excl. China & Unknown).',
    font=Font(name=F, size=9, italic=True)).border = Border()
put(wsV, r + 1, 1, f'Check: RoW TOTAL {CUR} minus Summary RoW (must be 0)', font=Font(name=F, size=9, italic=True)).border = Border()
put(wsV, r + 1, cY[CUR], f"={L(cY[CUR])}{r-1}-Summary!B8", MT)
put(wsV, r + 1, cF, f"={L(cF)}{r-1}-(SUMIFS({rng('I')},{rng('A')},\"{LY}\")-SUMIFS({rng('I')},{rng('A')},\"{LY}\",{rng('D')},\"CHINA, PEOPLES REPUBLIC OF\")-SUMIFS({rng('I')},{rng('A')},\"{LY}\",{rng('D')},\"UNKNOWN\"))", MT)
put(wsV, r + 1, cF - 1, 'Full-year check →', font=Font(name=F, size=9, italic=True)).border = Border()
for k, t in enumerate([
    f'Full-MY exports = accumulated exports on the last in-season report of each MY (≈ week 53). Booked-by-wk-{WK} % above 100% means part of the early book was cancelled, rolled to the next MY, or re-declared elsewhere.',
    f'5-yr avg pace = 5-yr avg wk-{WK} commitment ÷ 5-yr avg full-year exports (ratio of averages, robust to years with zero). "Implied full year" = {CUR} commitment ÷ that pace — mechanical, noisy for 1–3 cargo buyers.']):
    put(wsV, r + 3 + k, 1, t, font=Font(name=F, size=9, italic=True)).border = Border()
wsV.column_dimensions['A'].width = 30; wsV.column_dimensions['B'].width = 28; wsV.column_dimensions['C'].width = 22
for j in range(4, cF + len(FULLH)): wsV.column_dimensions[L(j)].width = 12
wsV.column_dimensions[L(cMd + 6)].width = 3
wsV.freeze_panes = 'D5'
for j in (cMd + 1, cMd + 3):
    rg = f'{L(j)}{vfirst}:{L(j)}{vlast}'
    wsV.conditional_formatting.add(rg, CellIsRule(operator='greaterThan', formula=['0'], fill=gr))
    wsV.conditional_formatting.add(rg, CellIsRule(operator='lessThan', formula=['0'], fill=rd))

# ---------------- Weekly Commitments ----------------
title(wsW, f'Cumulative total commitment by in-season week — {CUR} vs prior MYs (MT)',
      'Aggregated from the ESR file (hardcoded, blue). RoW excludes China and Unknown. 5-yr avg columns are formulas.')
wk = ins.copy(); wk['grp'] = wk.Country.map(lambda c: 'China' if c.startswith('CHINA') else ('Unknown' if c == 'UNKNOWN' else 'RoW'))
pv = wk.pivot_table(index='WeekNumber', columns=['grp', 'MY'], values='TotalCommitment', aggfunc='sum')
pv = pv.loc[1:52].fillna(0)  # destination absent that week = 0 commitment
hdr = ['Week']
blocks = []
for g in ['Total', 'China', 'RoW']:
    for m in [CUR, LY]: hdr.append(f'{g} {m}')
    hdr.append(f'{g} 5-yr avg')
header(wsW, 4, hdr)
for w in pv.index:
    r = 4 + w
    put(wsW, r, 1, int(w), '0')
    col = 2
    for g in ['Total', 'China', 'RoW']:
        for m in [CUR, LY]:
            if g == 'Total':
                v = sum(pv.get((gg, m), pd.Series(dtype=float)).get(w, float('nan')) for gg in ['China', 'Unknown', 'RoW'] if (gg, m) in pv.columns)
            else:
                v = pv[(g, m)].get(w) if (g, m) in pv.columns else None
            v = None if v is None or pd.isna(v) else float(v)
            put(wsW, r, col, v, MT, font=BLUE); col += 1
        if g == 'Total':
            vals = [pv[[(gg, m) for gg in ['China', 'Unknown', 'RoW'] if (gg, m) in pv.columns]].loc[w].sum() for m in FIVE]
        else:
            vals = [pv[(g, m)].get(w, 0) if (g, m) in pv.columns else 0 for m in FIVE]
        # store per-year values on hidden helper area to keep avg as formula
        hc = 20 + ['Total', 'China', 'RoW'].index(g) * 5
        for k, v in enumerate(vals):
            put(wsW, r, hc + k, float(0 if pd.isna(v) else v), MT, font=BLUE)
        put(wsW, r, col, f'=AVERAGE({L(hc)}{r}:{L(hc+4)}{r})', MT); col += 1
for gi, g in enumerate(['Total', 'China', 'RoW']):
    for k, m in enumerate(FIVE):
        c = wsW.cell(4, 20 + gi * 5 + k, f'{g} {m}'); c.font = Font(name=F, size=9, bold=True)
wsW.cell(3, 20, 'Helper: per-year values feeding the 5-yr avg columns').font = Font(name=F, size=9, italic=True)
# blank current-MY cells beyond current week
for r in range(5 + WK, 57):
    for col in (2, 5, 8): wsW.cell(r, col).value = None
for col in range(1, 35): wsW.column_dimensions[L(col)].width = 13
wsW.freeze_panes = 'B5'
for i, (g, c1) in enumerate([('Total', 2), ('China', 5), ('RoW', 8)]):
    lc = LineChart(); lc.title = f'{g} cumulative commitment (MT)'; lc.y_axis.numFmt = '#,##0'
    lc.add_data(Reference(wsW, min_col=c1, max_col=c1 + 2, min_row=4, max_row=56), titles_from_data=True)
    lc.set_categories(Reference(wsW, min_col=1, min_row=5, max_row=56)); lc.width = 16; lc.height = 7.5
    lc.x_axis.title = 'MY week'; lc.x_axis.delete = False; lc.y_axis.delete = False
    wsW.add_chart(lc, f'K{5 + i * 16}')

# ================= Demand & seasonal pace =================
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.datavalidation import DataValidation
YEL = PatternFill('solid', fgColor='FFFF00')
wsP = wb.create_sheet('Demand & Pace', index=1)
wsC = wb.create_sheet('Pace Chart', index=2)
wsK = wb.create_sheet('Weekly by Destination', index=len(wb.sheetnames) - 1)

# ---- weekly cumulative commitment by destination x MY (source) ----
NW = 52
wkd = ins[ins.MY.isin(FIVE + [CUR])]
wp = wkd.pivot_table(index=['MY', 'Country'], columns='WeekNumber', values='TotalCommitment', aggfunc='sum')
wp = wp.reindex(columns=range(1, NW + 1)).fillna(0)
title(wsK, f'Cumulative total commitment by destination and in-season week (MT) — {FIVE[-1]} … {CUR}',
      f'Hardcoded from the ESR file (blue). Full-MY exports = accumulated exports on the last in-season report. {CUR} filled through week {WK} only.')
header(wsK, 4, ['Marketing year', 'Destination', 'Full-MY exports'] + [f'W{w}' for w in range(1, NW + 1)])
r = 5
for (m, c), row in wp.sort_index().iterrows():
    put(wsK, r, 1, m, font=BLUE); put(wsK, r, 2, c, font=BLUE)
    fe = None if m == CUR else float(final.get((m, c), 0) or 0)
    put(wsK, r, 3, fe, MT, font=BLUE)
    for w in range(1, NW + 1):
        if m == CUR and w > WK: continue
        wsK.cell(r, 3 + w, float(row[w])).number_format = MT
        wsK.cell(r, 3 + w).font = BLUE
    r += 1
KL = r - 1
wsK.column_dimensions['A'].width = 12; wsK.column_dimensions['B'].width = 28; wsK.column_dimensions['C'].width = 14
for w in range(1, NW + 1): wsK.column_dimensions[L(3 + w)].width = 11
wsK.freeze_panes = 'D5'
KM, KC, KF = f"'Weekly by Destination'!$A$5:$A${KL}", f"'Weekly by Destination'!$B$5:$B${KL}", f"'Weekly by Destination'!$C$5:$C${KL}"
KW = f"'Weekly by Destination'!$D$5:${L(3+NW)}${KL}"

# ---- inputs ----
title(wsP, f'Annual demand by destination and {CUR} seasonal pace — are buyers ahead of or behind the pace needed to reach demand?',
      'Yellow = inputs. Demand = 5-yr avg full-MY exports per destination, scaled so RoW sums to USDA total minus the China assumption (override any country in col F). '
      'Pace = average across base years of (commitment at week N ÷ that year\'s full-MY exports); years with zero exports are skipped. Expected = demand × pace.')
inp = [
    ('USDA total US soybean exports (mln bu)', 1695, '#,##0', True, 'User input (USDA forecast).'),
    ('Conversion (MT per bushel)', 0.0272155, '0.0000000', False, '1 bu soybeans = 60 lb = 27.2155 kg.'),
    ('USDA total exports (MT)', '=B4*1000000*B5', MT, None, ''),
    ('China assumption (MT)', 25000000, MT, True, 'User assumption.'),
    ('RoW demand = USDA − China (MT)', '=B6-B7', MT, None, 'Includes RoW\'s share of sales now booked as Unknown.'),
    ('Pace week N (in-season week)', WK, '0', True, f'Latest data = week {WK}. Actuals exist only up to week {WK}.'),
    ('RoW 5-yr avg full-MY exports, sum of destinations (MT)', None, MT, None, 'ESR basis.'),
    ('USDA-implied RoW vs 5-yr avg', None, PCT, None, ''),
    ('RoW scaling factor applied to each destination', None, '0.000', None, ''),
]
for i, (lab, v, fmt, is_in, note) in enumerate(inp):
    r = 4 + i
    put(wsP, r, 1, lab, font=BOLD)
    c = put(wsP, r, 2, v, fmt, font=BLUE if is_in else NORM, fill=YEL if is_in else None)
    wsP.cell(r, 3, note).font = Font(name=F, size=9, italic=True, color='595959')
put(wsP, 13, 1, 'Base marketing years for demand & pace', font=BOLD)
for k, m in enumerate(FIVE): put(wsP, 13, 2 + k, m, font=BLUE, fill=YEL)
wsP.cell(13, 7, 'Edit to change the base (e.g. drop 2025/26 for China-sensitive reads).').font = Font(name=F, size=9, italic=True, color='595959')
for nm, ref in [('Pace_Week', '$B$9'), ('RoW_Target', '$B$8'), ('China_Target', '$B$7'), ('Base_MY', '$B$13:$F$13'), ('RoW_Scale', '$B$12')]:
    wb.defined_names[nm] = DefinedName(nm, attr_text=f"'Demand & Pace'!{ref}")

# ---- per-country table ----
HR = 15
ph = ['Status (±1 cargo = on pace)', 'Destination', 'Region', 'Demand: own 5-yr avg full-MY exports', 'Demand: scaled to USDA',
      'Demand override (optional)', 'Demand used', 'Avg pace at wk N (% of year committed)', 'Pace min (base yrs)', 'Pace max (base yrs)',
      'Expected commitment at wk N', f'Actual {CUR} at wk N', 'Δ vs expected MT', 'Δ vs expected %', '% of demand committed',
      'Remaining to sell (MT)', 'Weekly sales needed rest of MY', 'Typical weekly sales rest of MY', 'Needed ÷ typical',
      'Expected at wk N using own 5-yr demand', 'Δ vs own-history expected MT']
header(wsP, HR, ph); wsP.row_dimensions[HR].height = 70
HC = len(ph) + 2   # helper per-year pace columns start (one spacer)
for k in range(5):
    c = wsP.cell(HR, HC + k, f'Pace yr {k+1} (=Base_MY col {k+1})'); c.font = Font(name=F, size=8, bold=True, color='595959'); c.alignment = Alignment(wrap_text=True)

# python mirror for ordering (USDA-scaled demand, avg-of-ratios pace)
def pace_py(c, w):
    vals = []
    for m in FIVE:
        fe = final.get((m, c), 0) or 0
        if fe > 0:
            cm = wp.loc[(m, c), w] if (m, c) in wp.index else 0
            vals.append(cm / fe)
    return sum(vals) / len(vals) if vals else 0
dem5 = {c: sum((final.get((m, c), 0) or 0) for m in FIVE) / 5 for c in set(ins.Country)}
row_sum = sum(v for c, v in dem5.items() if c not in ('CHINA, PEOPLES REPUBLIC OF', 'UNKNOWN'))
row_target = 1695e6 * 0.0272155 - 25e6
curc = {c: (wp.loc[(CUR, c), WK] if (CUR, c) in wp.index else 0) for c in set(ins.Country)}
pace_list = [c for c in set(ins.Country) if c not in ('CHINA, PEOPLES REPUBLIC OF', 'UNKNOWN') and (dem5[c] > 0 or curc[c] > 0)]
def dpy(c): return curc[c] - dem5[c] * row_target / row_sum * pace_py(c, WK)
pace_list.sort(key=lambda c: -dpy(c))

def wcol(): return f'INDEX({KW},0,Pace_Week)'
def prow(r, c, is_china):
    B = f'$B{r}'
    put(wsP, r, 2, c, font=BOLD if is_china else NORM, fill=TOTFILL if is_china else None)
    put(wsP, r, 3, REGION[c], font=GREEN)
    put(wsP, r, 4, f'=SUMPRODUCT(({KC}={B})*ISNUMBER(MATCH({KM},Base_MY,0))*{KF})/COUNTA(Base_MY)', MT)
    put(wsP, r, 5, '=China_Target' if is_china else f'=D{r}*RoW_Scale', MT)
    put(wsP, r, 6, None, MT, font=BLUE, fill=YEL)
    put(wsP, r, 7, f'=IF(F{r}<>"",F{r},E{r})', MT, font=BOLD)
    for k in range(5):
        y = f"INDEX(Base_MY,{k+1})"
        put(wsP, r, HC + k, f'=IFERROR(SUMIFS({wcol()},{KC},{B},{KM},{y})/SUMIFS({KF},{KC},{B},{KM},{y}),"")', '0%',
            font=Font(name=F, size=8, color='595959'))
    hp = f'{L(HC)}{r}:{L(HC+4)}{r}'
    put(wsP, r, 8, f'=IFERROR(AVERAGE({hp}),0)', '0%')
    put(wsP, r, 9, f'=IF(COUNT({hp})=0,"",MIN({hp}))', '0%'); put(wsP, r, 10, f'=IF(COUNT({hp})=0,"",MAX({hp}))', '0%')
    put(wsP, r, 11, f'=G{r}*H{r}', MT)
    put(wsP, r, 12, f'=IF(Pace_Week>{WK},"",SUMIFS({wcol()},{KC},{B},{KM},"{CUR}"))', MT, fill=CURFILL)
    put(wsP, r, 13, f'=IF(L{r}="","",L{r}-K{r})', MT, font=BOLD)
    put(wsP, r, 14, f'=IF(L{r}="","",IF(K{r}=0,IF(L{r}=0,0,"new"),L{r}/K{r}-1))', PCT)
    put(wsP, r, 1, f'=IF(M{r}="","",IF(ABS(M{r})<Cargo_MT,"On pace",IF(M{r}>0,"ABOVE pace","BELOW pace")))')
    put(wsP, r, 15, f'=IF(G{r}=0,"",L{r}/G{r})', '0%')
    put(wsP, r, 16, f'=G{r}-L{r}', MT)
    put(wsP, r, 17, f'=MAX(0,P{r})/(52-Pace_Week)', MT)
    put(wsP, r, 18, f'=G{r}*MAX(0,1-H{r})/(52-Pace_Week)', MT)
    put(wsP, r, 19, f'=IF(R{r}=0,"",Q{r}/R{r})', '0.00"x"')
    put(wsP, r, 20, f'=D{r}*H{r}', MT); put(wsP, r, 21, f'=IF(L{r}="","",L{r}-T{r})', MT)

r = HR + 1
prow(r, 'CHINA, PEOPLES REPUBLIC OF', True); CH = r
r += 1
put(wsP, r, 1, 'Rest of World — sorted by Δ vs expected (largest surplus first)', font=Font(name=F, size=10, bold=True)).border = Border()
r += 1; P1 = r
for c in pace_list:
    prow(r, c, False); r += 1
P2 = r - 1
TR = r
put(wsP, TR, 2, 'RoW TOTAL (sum of destinations)', font=BOLD, fill=TOTFILL); put(wsP, TR, 3, '', fill=TOTFILL)
for j in (4, 5, 7, 11, 12, 13, 16, 17, 18, 20, 21):
    put(wsP, TR, j, f'=SUM({L(j)}{P1}:{L(j)}{P2})', MT, font=BOLD, fill=TOTFILL)
put(wsP, TR, 8, f'=IF(G{TR}=0,0,K{TR}/G{TR})', '0%', font=BOLD, fill=TOTFILL)
put(wsP, TR, 14, f'=IF(K{TR}=0,"",L{TR}/K{TR}-1)', PCT, font=BOLD, fill=TOTFILL)
put(wsP, TR, 15, f'=IF(G{TR}=0,"",L{TR}/G{TR})', '0%', font=BOLD, fill=TOTFILL)
put(wsP, TR, 19, f'=IF(R{TR}=0,"",Q{TR}/R{TR})', '0.00"x"', font=BOLD, fill=TOTFILL)
GT2 = TR + 1
put(wsP, GT2, 2, 'TOTAL China + RoW', font=BOLD, fill=TOTFILL); put(wsP, GT2, 3, '', fill=TOTFILL)
for j in (4, 5, 7, 11, 12, 13, 16, 17, 18, 20, 21):
    put(wsP, GT2, j, f'={L(j)}{CH}+{L(j)}{TR}', MT, font=BOLD, fill=TOTFILL)
put(wsP, GT2, 14, f'=IF(K{GT2}=0,"",L{GT2}/K{GT2}-1)', PCT, font=BOLD, fill=TOTFILL)
put(wsP, GT2, 15, f'=IF(G{GT2}=0,"",L{GT2}/G{GT2})', '0%', font=BOLD, fill=TOTFILL)
put(wsP, GT2, 19, f'=IF(R{GT2}=0,"",Q{GT2}/R{GT2})', '0.00"x"', font=BOLD, fill=TOTFILL)
# inputs depending on table
wsP['B10'] = f'=D{TR}'; wsP['B11'] = '=IF(B10=0,0,RoW_Target/B10-1)'; wsP['B12'] = '=IF(B10=0,0,RoW_Target/B10)'
for a in ('B10', 'B11', 'B12'): wsP[a].font = NORM
# memo: Unknown
m0 = GT2 + 2
put(wsP, m0, 1, 'Memo — Unknown destinations', font=Font(name=F, size=11, bold=True)).border = Border()
put(wsP, m0 + 1, 1, f'Unknown commitment at wk N ({CUR})'); put(wsP, m0 + 1, 12, f'=IF(Pace_Week>{WK},"",SUMIFS({wcol()},{KC},"UNKNOWN",{KM},"{CUR}"))', MT)
put(wsP, m0 + 2, 1, 'China + Unknown at wk N vs China expected'); put(wsP, m0 + 2, 11, f'=K{CH}', MT); put(wsP, m0 + 2, 12, f'=L{CH}+L{m0+1}', MT)
put(wsP, m0 + 2, 13, f'=L{m0+2}-K{m0+2}', MT, font=BOLD)
put(wsP, m0 + 3, 1, 'Check: China + RoW + Unknown actual minus ESR total at wk N (must be 0)')
put(wsP, m0 + 3, 12, f'=IF(Pace_Week>{WK},"",L{GT2}+L{m0+1}-SUMIFS({wcol()},{KM},"{CUR}"))', MT)
notes = [
    'How to read: Expected = Demand used × average share of the year that the destination had committed by week N in the base years. ABOVE/BELOW pace only when |Δ| ≥ one cargo (Cargo_MT on the RoW tab).',
    f'Unknown ({CUR}) is not allocated to any destination. Historically most of it is re-declared to China at shipment, so China\'s own line understates its true pace; compare the memo line.',
    'Pace uses ESR full-MY exports; USDA\'s total is Census basis, which usually runs above ESR. Scaling RoW to USDA therefore mixes bases slightly — treat Δ within a few % as noise.',
    'China pace includes 2025/26 when China barely booked US beans early; that lowers China\'s average pace. Edit Base_MY (row 13) to test excluding it.',
    'Small buyers (1–3 cargoes/season) have erratic pace; rely on the larger destinations and on the totals.']
for k, t in enumerate(notes):
    put(wsP, m0 + 5 + k, 1, t, font=Font(name=F, size=9, italic=True)).border = Border()
wsP.column_dimensions['A'].width = 44; wsP.column_dimensions['B'].width = 28; wsP.column_dimensions['C'].width = 20
for j in range(4, len(ph) + 1): wsP.column_dimensions[L(j)].width = 13
wsP.column_dimensions[L(len(ph) + 1)].width = 3
for k in range(5): wsP.column_dimensions[L(HC + k)].width = 9
wsP.freeze_panes = f'D{HR+1}'
for j in (13, 21):
    rg = f'{L(j)}{CH}:{L(j)}{P2}'
    wsP.conditional_formatting.add(rg, CellIsRule(operator='greaterThan', formula=['0'], fill=gr))
    wsP.conditional_formatting.add(rg, CellIsRule(operator='lessThan', formula=['0'], fill=rd))

# ---- Pace Chart (selectable destination) ----
title(wsC, 'Seasonal pace chart — pick a destination in B4',
      'Expected path = Demand used (from "Demand & Pace") × average base-year pace for each week. Actual = cumulative total commitment.')
put(wsC, 4, 1, 'Destination', font=BOLD); put(wsC, 4, 2, 'EGYPT', font=BLUE, fill=YEL)
dv = DataValidation(type='list', formula1=f"='Demand & Pace'!$B${CH}:$B${P2}", allow_blank=False); wsC.add_data_validation(dv); dv.add('B4')
put(wsC, 5, 1, 'Demand used (MT)', font=BOLD)
put(wsC, 5, 2, f"=INDEX('Demand & Pace'!$G:$G,MATCH($B$4,'Demand & Pace'!$B:$B,0))", MT, font=GREEN)
put(wsC, 6, 1, 'Status at wk N', font=BOLD)
put(wsC, 6, 2, f"=INDEX('Demand & Pace'!$A:$A,MATCH($B$4,'Demand & Pace'!$B:$B,0))", font=GREEN)
ch_h = ['Week'] + [f'Pace yr {k+1}' for k in range(5)] + ['Avg pace', 'Expected path (MT)', f'Actual {CUR}', f'Actual {LY}']
header(wsC, 8, ch_h)
for w in range(1, NW + 1):
    r = 8 + w
    put(wsC, r, 1, w, '0')
    col = f"INDEX({KW},0,$A{r})"
    for k in range(5):
        y = f'INDEX(Base_MY,{k+1})'
        put(wsC, r, 2 + k, f'=IFERROR(SUMIFS({col},{KC},$B$4,{KM},{y})/SUMIFS({KF},{KC},$B$4,{KM},{y}),"")', '0%', font=Font(name=F, size=9, color='595959'))
    put(wsC, r, 7, f'=IFERROR(AVERAGE(B{r}:F{r}),0)', '0%')
    put(wsC, r, 8, f'=$B$5*G{r}', MT)
    if w <= WK: put(wsC, r, 9, f'=SUMIFS({col},{KC},$B$4,{KM},"{CUR}")', MT, fill=CURFILL)
    put(wsC, r, 10, f'=SUMIFS({col},{KC},$B$4,{KM},"{LY}")', MT)
for j, w_ in zip('ABCDEFGHIJ', [8, 9, 9, 9, 9, 9, 10, 16, 14, 14]): wsC.column_dimensions[j].width = w_
wsC.column_dimensions['A'].width = 18
lc = LineChart(); lc.title = 'Cumulative commitment vs expected pace (selected destination)'; lc.y_axis.numFmt = '#,##0'
lc.add_data(Reference(wsC, min_col=8, max_col=10, min_row=8, max_row=8 + NW), titles_from_data=True)
lc.set_categories(Reference(wsC, min_col=1, min_row=9, max_row=8 + NW)); lc.width = 20; lc.height = 10
lc.x_axis.title = 'MY week'; lc.x_axis.delete = False; lc.y_axis.delete = False
wsC.add_chart(lc, 'L4')

for ws in wb:
    ws.sheet_view.showGridLines = False
wb.save(out)
print('CUR', CUR, 'WK', WK, 'LY', LY, 'FIVE', FIVE, 'rows', len(order))
