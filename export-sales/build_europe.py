"""Build a stand-alone Europe workbook for US soybean export sales.
Usage: python build_europe.py <ESR csv> <out.xlsx>
Source: USDA FAS Export Sales Reporting (ESR) query export, commodity SOYBEANS (801).
Self-contained: weekly data for European destinations, full-year exports for every RoW destination
(needed for the floor+swing allocation), inputs, live formulas throughout."""
import sys
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.chart import LineChart, BarChart, Reference
from openpyxl.utils import get_column_letter as L
from openpyxl.formatting.rule import CellIsRule
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.datavalidation import DataValidation

src, out = sys.argv[1], sys.argv[2]
df = pd.read_csv(src, thousands=',')
df.columns = [c.strip() for c in df.columns]
df['Country'] = df['CountryName'].str.strip()
df['date'] = pd.to_datetime(df['WeekendingDate'])
df['MYstart'] = df['ReportingMarketingYearName'].str.extract(r'Sep (\d{4})').astype(int)
df['MY'] = df['MYstart'].astype(str) + '/' + (df['MYstart'] + 1).astype(str).str[-2:]
lo = pd.to_datetime(df['MYstart'].astype(str) + '-08-31')
hi = pd.to_datetime((df['MYstart'] + 1).astype(str) + '-09-07')
ins = df[(df.date >= lo) & (df.date <= hi)].copy()
cs = int(ins.loc[ins.date.idxmax(), 'MYstart'])
CUR = f"{cs}/{str(cs+1)[-2:]}"
WK = int(ins[ins.MYstart == cs].WeekNumber.max())
CUR_DATE = ins[ins.MYstart == cs].date.max()
MYS = [f"{y}/{str(y+1)[-2:]}" for y in range(cs - 6, cs + 1)]          # 6 prior + current
PRIOR = MYS[:-1]; LY = PRIOR[-1]; FIVE = PRIOR[::-1][:5]                  # FIVE newest first
ins = ins[ins.MY.isin(MYS)]
NW = 52

EU27 = ['SPAIN', 'NETHERLANDS', 'GERMANY', 'PORTUGAL', 'ITALY', 'GREECE', 'BELGIUM-LUXEMBOURG', 'FRANCE', 'IRELAND', 'ROMANIA',
        'AUSTRIA', 'BULGARIA', 'CROATIA', 'CYPRUS', 'CZECH REPUBLIC', 'DENMARK', 'ESTONIA', 'FINLAND', 'HUNGARY', 'LATVIA',
        'LITHUANIA', 'MALTA', 'POLAND', 'SLOVAKIA', 'SLOVENIA', 'SWEDEN']
NONEU = ['UNITED KINGDOM', 'NORWAY', 'SWITZERLAND', 'UKRAINE', 'GEORGIA, REPUBLIC OF', 'ICELAND', 'SERBIA', 'ALBANIA']
present = set(ins.Country)
fin = ins[ins.MY != CUR].sort_values('date').groupby(['MY', 'Country']).tail(1).set_index(['MY', 'Country']).AccumulatedExports
def fy(m, c): return float(fin.get((m, c), 0) or 0)
EUR = [c for c in EU27 + NONEU if c in present]
avg5 = {c: sum(fy(m, c) for m in FIVE) / 5 for c in present}
EUR.sort(key=lambda c: (c not in EU27, -avg5[c]))
GRP = {c: 'EU' if c in EU27 else 'non-EU' for c in EUR}
ROW = sorted([c for c in present if c not in ('CHINA, PEOPLES REPUBLIC OF', 'UNKNOWN')], key=lambda c: -avg5[c])

# ---------------- styles ----------------
F = 'Arial'
HFILL = PatternFill('solid', fgColor='1F4E78'); TOTFILL = PatternFill('solid', fgColor='EDEDED')
CURFILL = PatternFill('solid', fgColor='FFF2CC'); YEL = PatternFill('solid', fgColor='FFFF00')
GR = PatternFill('solid', fgColor='C6EFCE'); RD = PatternFill('solid', fgColor='FFC7CE')
thin = Side(style='thin', color='BFBFBF'); BOX = Border(top=thin, bottom=thin, left=thin, right=thin)
MT = '#,##0;-#,##0;"-"'; PCT = '+0.0%;-0.0%;0.0%'; BU = '#,##0.0;-#,##0.0;"-"'
BLUE = Font(name=F, size=10, color='0000FF'); GREEN = Font(name=F, size=10, color='008000')
NORM = Font(name=F, size=10); BOLD = Font(name=F, size=10, bold=True)
SMALL = Font(name=F, size=8, color='595959'); NOTE = Font(name=F, size=9, italic=True, color='404040')

def title(ws, t, note):
    ws['A1'] = t; ws['A1'].font = Font(name=F, size=13, bold=True, color='1F4E78')
    ws['A2'] = note; ws['A2'].font = Font(name=F, size=9, italic=True, color='595959')
def header(ws, row, labels, col=1, h=40):
    for i, x in enumerate(labels):
        c = ws.cell(row, col + i, x); c.font = Font(name=F, size=9, bold=True, color='FFFFFF')
        c.fill = HFILL; c.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True); c.border = BOX
    ws.row_dimensions[row].height = h
def put(ws, r, c, v, fmt=None, font=NORM, fill=None, border=True):
    x = ws.cell(r, c, v); x.font = font
    if border: x.border = BOX
    if fmt: x.number_format = fmt
    if fill: x.fill = fill
    return x
def text(ws, r, c, t, font=NOTE): x = ws.cell(r, c, t); x.font = font; return x
def section(ws, r, t): text(ws, r, 1, t, Font(name=F, size=11, bold=True, color='1F4E78'))
def name(wb, n, sheet, ref): wb.defined_names[n] = DefinedName(n, attr_text=f"'{sheet}'!{ref}")
def dpct(a, b): return f'IF({b}=0,IF({a}=0,0,"new"),{a}/{b}-1)'
def posneg(ws, rg):
    ws.conditional_formatting.add(rg, CellIsRule(operator='greaterThan', formula=['0'], fill=GR))
    ws.conditional_formatting.add(rg, CellIsRule(operator='lessThan', formula=['0'], fill=RD))

wb = Workbook()
wsH = wb.active; wsH.title = 'Dashboard'
wsDst = wb.create_sheet('Destinations')
wsWk = wb.create_sheet('Weekly Europe')
wsWin = wb.create_sheet('Nov-Feb Window')
wsCC = wb.create_sheet('Country Chart')
wsHis = wb.create_sheet('History')
wsN = wb.create_sheet('Method & Notes')
wsFY = wb.create_sheet('Data RoW Full-Year')
wsDC = wb.create_sheet('Data Commit')
wsDX = wb.create_sheet('Data Exports')

# ================= Inputs (Dashboard top) =================
title(wsH, f'Europe — US soybean export sales monitor, MY {CUR} (data through w/e {CUR_DATE:%d %b %Y}, week {WK})',
      'Yellow cells are inputs. Everything else is formula-driven from the Data tabs (USDA FAS Export Sales, SOYBEANS). '
      'Europe = EU-27 + non-EU Europe (UK, Norway, Switzerland, …).')
inputs = [
    ('USDA total US soybean exports (mln bu)', 1695, '#,##0', True, 'User input (USDA forecast).'),
    ('Conversion (MT per bushel)', 0.0272155, '0.0000000', True, '60 lb bushel = 27.2155 kg.'),
    ('USDA total exports (MT)', '=B4*1000000*B5', MT, False, ''),
    ('China assumption (MT)', 25000000, MT, True, 'User assumption.'),
    ('RoW demand = USDA − China (MT)', '=B6-B7', MT, False, 'Shared among all non-China destinations (incl. Europe).'),
    ('Pace week N', WK, '0', True, f'Latest available week = {WK}.'),
    ('Demand allocation method', 'Floor+swing (refined)', None, True, 'One factor / Floor+swing / Floor+swing (refined) — see "Method & Notes".'),
    ('Cargo size for "on pace" band (MT)', 65000, MT, True, '≈ one Panamax: smaller gaps = timing noise.'),
    ('Current MY', CUR, None, False, ''),
    ('Last MY', LY, None, False, ''),
]
for i, (lab, v, fmt, is_in, nt) in enumerate(inputs):
    r = 4 + i
    put(wsH, r, 1, lab, font=BOLD)
    put(wsH, r, 2, v, fmt, font=BLUE if is_in else NORM, fill=YEL if is_in else None)
    text(wsH, r, 3, nt)
wsH['B10'].alignment = Alignment(horizontal='center')
put(wsH, 14, 1, 'Base marketing years (newest first)', font=BOLD)
for k, m in enumerate(FIVE): put(wsH, 14, 2 + k, m, font=BLUE, fill=YEL)
put(wsH, 15, 1, 'Method index', font=BOLD)
put(wsH, 15, 2, '=MATCH(B10,{"One factor","Floor+swing","Floor+swing (refined)"},0)', '0')
dv = DataValidation(type='list', formula1='"One factor,Floor+swing,Floor+swing (refined)"'); wsH.add_data_validation(dv); dv.add('B10')
dvb = DataValidation(type='list', formula1='"' + ','.join(PRIOR[::-1]) + '"'); wsH.add_data_validation(dvb)
for k in range(5): dvb.add(wsH.cell(14, 2 + k))
for n, ref in [('Conv', '$B$5'), ('RoW_Target', '$B$8'), ('China_Target', '$B$7'), ('Pace_Week', '$B$9'), ('Cargo_MT', '$B$11'),
               ('CUR_MY', '$B$12'), ('LY_MY', '$B$13'), ('Base_MY', '$B$14:$F$14'), ('Method_Idx', '$B$15')]:
    name(wb, n, 'Dashboard', ref)

# ================= Data Commit / Data Exports =================
def data_sheet(ws, field, label, pref):
    title(ws, f'{label} — cumulative by in-season week, European destinations (MT)',
          f'Hardcoded from the ESR file (blue). {CUR} filled through week {WK}. Group: EU = EU-27 member, non-EU = rest of Europe.')
    header(ws, 4, ['Marketing year', 'Destination', 'Group'] + [f'W{w}' for w in range(1, NW + 1)])
    pv = ins[ins.Country.isin(EUR)].pivot_table(index=['MY', 'Country'], columns='WeekNumber', values=field, aggfunc='sum')
    pv = pv.reindex(columns=range(1, NW + 1)).fillna(0)
    r = 5
    for m in MYS:
        for c in EUR:
            put(ws, r, 1, m, font=BLUE, border=False); put(ws, r, 2, c, font=BLUE, border=False); put(ws, r, 3, GRP[c], font=BLUE, border=False)
            for w in range(1, NW + 1):
                if m == CUR and w > WK: continue
                v = float(pv.loc[(m, c), w]) if (m, c) in pv.index else 0.0
                x = ws.cell(r, 3 + w, v); x.number_format = MT; x.font = BLUE
            r += 1
    last = r - 1
    ws.column_dimensions['A'].width = 12; ws.column_dimensions['B'].width = 26; ws.column_dimensions['C'].width = 8
    for w in range(1, NW + 1): ws.column_dimensions[L(3 + w)].width = 10
    ws.freeze_panes = 'D5'
    sh = ws.title
    name(wb, f'{pref}_MY', sh, f'$A$5:$A${last}'); name(wb, f'{pref}_D', sh, f'$B$5:$B${last}')
    name(wb, f'{pref}_G', sh, f'$C$5:$C${last}'); name(wb, f'{pref}_W', sh, f'$D$5:${L(3+NW)}${last}')
data_sheet(wsDC, 'TotalCommitment', 'Total commitment (shipped + unshipped)', 'DC')
data_sheet(wsDX, 'AccumulatedExports', 'Accumulated exports (shipped)', 'DX')
def DCW(col): return f'INDEX(DC_W,0,{col})'
def DXW(col): return f'INDEX(DX_W,0,{col})'

# ================= Data RoW Full-Year (+ allocation engine) =================
title(wsFY, 'Full-MY exports by destination, all RoW destinations (excl. China & Unknown), and demand allocation engine (MT)',
      'Full-MY exports (blue) = accumulated exports on the last in-season report of each MY. Grey columns compute the allocation for the base years '
      'chosen on the Dashboard. The whole RoW is needed because floor+swing shares one RoW pool across all destinations.')
FYH = ['Destination', 'Europe group'] + PRIOR + [f'Base yr {k+1}' for k in range(5)] + [
    '5-yr avg', 'Floor (lowest)', 'Floor refined (2nd-lowest)', '5-yr high', 'Swing', 'Swing refined',
    'Demand: one factor', 'Demand: floor+swing', 'Demand: refined', 'Demand: selected method']
header(wsFY, 8, FYH, h=44)
cP0 = 3; cB0 = cP0 + len(PRIOR); cAvg = cB0 + 5
cMin, cMin2, cMax, cSw, cSw2, cU, cFS, cFR, cSel = range(cAvg + 1, cAvg + 10)
F1 = 9
for i, c in enumerate(ROW):
    r = F1 + i
    put(wsFY, r, 1, c, font=BLUE, border=False); put(wsFY, r, 2, GRP.get(c), font=BLUE, border=False)
    for k, m in enumerate(PRIOR): put(wsFY, r, cP0 + k, fy(m, c), MT, font=BLUE, border=False)
    for k in range(5):
        put(wsFY, r, cB0 + k, f'=INDEX(${L(cP0)}{r}:${L(cP0+len(PRIOR)-1)}{r},MATCH(INDEX(Base_MY,{k+1}),${L(cP0)}$8:${L(cP0+len(PRIOR)-1)}$8,0))', MT, font=SMALL, border=False)
    b = f'{L(cB0)}{r}:{L(cB0+4)}{r}'
    for col, f in [(cAvg, f'=AVERAGE({b})'), (cMin, f'=MIN({b})'), (cMin2, f'=SMALL({b},2)'), (cMax, f'=MAX({b})'),
                   (cSw, f'={L(cAvg)}{r}-{L(cMin)}{r}'), (cSw2, f'=MAX(0,{L(cAvg)}{r}-{L(cMin2)}{r})'),
                   (cU, f'={L(cAvg)}{r}*Scale_1F'), (cFS, f'={L(cMin)}{r}+{L(cSw)}{r}*F_Basic'),
                   (cFR, f'=MIN({L(cMax)}{r},{L(cMin2)}{r}+{L(cSw2)}{r}*F_Ref)'),
                   (cSel, f'=CHOOSE(Method_Idx,{L(cU)}{r},{L(cFS)}{r},{L(cFR)}{r})')]:
        put(wsFY, r, col, f, MT, font=BOLD if col == cSel else NORM, border=False)
F2 = F1 + len(ROW) - 1
FT = F2 + 1
put(wsFY, FT, 1, 'RoW TOTAL', font=BOLD, fill=TOTFILL)
for col in list(range(cP0, cB0 + 5)) + [cAvg, cMin, cMin2, cMax, cSw, cSw2, cU, cFS, cFR, cSel]:
    put(wsFY, FT, col, f'=SUM({L(col)}{F1}:{L(col)}{F2})', MT, font=BOLD, fill=TOTFILL)
eng = [('One-factor scale (RoW target ÷ Σ 5-yr avg)', f'=RoW_Target/{L(cAvg)}{FT}', '0.000'),
       ('Swing factor — floor+swing', f'=(RoW_Target-{L(cMin)}{FT})/{L(cSw)}{FT}', '0.000'),
       ('Swing factor — refined', f'=(RoW_Target-{L(cMin2)}{FT})/{L(cSw2)}{FT}', '0.000'),
       ('Check: Σ demand (selected) − RoW target', f'={L(cSel)}{FT}-RoW_Target', MT)]
for i, (lab, f, fmt) in enumerate(eng):
    put(wsFY, 3 + i, 1, lab, font=BOLD); put(wsFY, 3 + i, 2, f, fmt, font=BOLD)
text(wsFY, 7, 1, 'A negative swing factor means the RoW target is below the sum of floors: use "One factor" instead.')
name(wb, 'Scale_1F', wsFY.title, '$B$3'); name(wb, 'F_Basic', wsFY.title, '$B$4'); name(wb, 'F_Ref', wsFY.title, '$B$5')
wsFY.column_dimensions['A'].width = 40; wsFY.column_dimensions['B'].width = 12
for col in range(3, cSel + 1): wsFY.column_dimensions[L(col)].width = 12
wsFY.freeze_panes = f'C{F1}'
FYR = f"'Data RoW Full-Year'"
def fylook(col, rc): return f"IFERROR(INDEX({FYR}!${L(col)}:${L(col)},MATCH({rc},{FYR}!$A:$A,0)),0)"

# ================= Destinations =================
title(wsDst, f'European destinations — week-N commitment vs history, demand by method, seasonal pace (MT)',
      'Demand used = override if entered (col R), otherwise the method chosen on the Dashboard. Pace = average over base years of '
      '(commitment at week N ÷ full-MY exports); years with zero exports skipped. Status uses the cargo band.')
DH = ['Destination', 'Group', f'{CUR} wk N', f'{LY} wk N', '5-yr avg wk N', '5-yr median wk N', 'Δ vs 5-yr avg', 'Δ vs median',
      f'Shipped {CUR}', f'Unshipped {CUR}', f'Full-MY {LY}', 'Full-MY 5-yr avg', 'Floor refined (2nd-lowest)', '5-yr high',
      'Demand: one factor', 'Demand: floor+swing', 'Demand: refined', 'Demand override', 'Demand used',
      'Avg pace wk N', 'Pace min', 'Pace max', 'Expected wk N', 'Δ vs expected', 'Δ vs expected %', 'Status',
      '% of demand committed', 'Remaining to sell', 'Weekly sales needed rest of MY', 'Implied full year (MT)',
      'Implied full year (mln bu)', 'Demand used (mln bu)', 'Δ vs expected (mln bu)']
header(wsDst, 4, DH, h=52)
HB = len(DH) + 2   # helpers: commit base yr k (5), FY base yr k (5), pace yr k (5)
for k, h in enumerate([f'Commit wk N base yr {k+1}' for k in range(5)] + [f'Full-MY base yr {k+1}' for k in range(5)] + [f'Pace base yr {k+1}' for k in range(5)]):
    x = wsDst.cell(4, HB + k, h); x.font = Font(name=F, size=8, bold=True, color='595959'); x.alignment = Alignment(wrap_text=True)
D1 = 5
for i, c in enumerate(EUR):
    r = D1 + i; A = f'$A{r}'
    put(wsDst, r, 1, c); put(wsDst, r, 2, GRP[c])
    for k in range(5):
        put(wsDst, r, HB + k, f'=SUMIFS({DCW("Pace_Week")},DC_D,{A},DC_MY,INDEX(Base_MY,{k+1}))', MT, font=SMALL, border=False)
        put(wsDst, r, HB + 5 + k, '=' + fylook(cB0 + k, A), MT, font=SMALL, border=False)
        put(wsDst, r, HB + 10 + k, f'=IF({L(HB+5+k)}{r}=0,"",{L(HB+k)}{r}/{L(HB+5+k)}{r})', '0%', font=SMALL, border=False)
    cm = f'{L(HB)}{r}:{L(HB+4)}{r}'; pc = f'{L(HB+10)}{r}:{L(HB+14)}{r}'
    put(wsDst, r, 3, f'=IF(Pace_Week>{WK},"",SUMIFS({DCW("Pace_Week")},DC_D,{A},DC_MY,CUR_MY))', MT, fill=CURFILL)
    put(wsDst, r, 4, f'=SUMIFS({DCW("Pace_Week")},DC_D,{A},DC_MY,LY_MY)', MT)
    put(wsDst, r, 5, f'=AVERAGE({cm})', MT); put(wsDst, r, 6, f'=MEDIAN({cm})', MT)
    put(wsDst, r, 7, f'=IF(C{r}="","",C{r}-E{r})', MT, font=BOLD); put(wsDst, r, 8, f'=IF(C{r}="","",C{r}-F{r})', MT, font=BOLD)
    put(wsDst, r, 9, f'=IF(Pace_Week>{WK},"",SUMIFS({DXW("Pace_Week")},DC_D,{A},DC_MY,CUR_MY))', MT)
    put(wsDst, r, 10, f'=IF(C{r}="","",C{r}-I{r})', MT)
    put(wsDst, r, 11, f"=SUMIFS({FYR}!${L(cP0+len(PRIOR)-1)}:${L(cP0+len(PRIOR)-1)},{FYR}!$A:$A,{A})", MT)
    for col, src in ((12, cAvg), (13, cMin2), (14, cMax), (15, cU), (16, cFS), (17, cFR)):
        put(wsDst, r, col, '=' + fylook(src, A), MT, font=GREEN)
    put(wsDst, r, 18, None, MT, font=BLUE, fill=YEL)
    put(wsDst, r, 19, f'=IF(R{r}<>"",R{r},CHOOSE(Method_Idx,O{r},P{r},Q{r}))', MT, font=BOLD)
    put(wsDst, r, 20, f'=IFERROR(AVERAGE({pc}),0)', '0%')
    put(wsDst, r, 21, f'=IF(COUNT({pc})=0,"",MIN({pc}))', '0%'); put(wsDst, r, 22, f'=IF(COUNT({pc})=0,"",MAX({pc}))', '0%')
    put(wsDst, r, 23, f'=S{r}*T{r}', MT)
    put(wsDst, r, 24, f'=IF(C{r}="","",C{r}-W{r})', MT, font=BOLD)
    put(wsDst, r, 25, f'=IF(C{r}="","",{dpct(f"C{r}", f"W{r}")})', PCT)
    put(wsDst, r, 26, f'=IF(X{r}="","",IF(ABS(X{r})<Cargo_MT,"On pace",IF(X{r}>0,"ABOVE pace","BELOW pace")))')
    put(wsDst, r, 27, f'=IF(OR(S{r}=0,C{r}=""),"",C{r}/S{r})', '0%')
    put(wsDst, r, 28, f'=IF(C{r}="","",S{r}-C{r})', MT)
    put(wsDst, r, 29, f'=IF(AB{r}="","",MAX(0,AB{r})/(52-Pace_Week))', MT)
    put(wsDst, r, 30, f'=IF(OR(N(T{r})=0,C{r}=""),"",C{r}/T{r})', MT)
    put(wsDst, r, 31, f'=IF(AD{r}="","",AD{r}/Conv/1000000)', BU)
    put(wsDst, r, 32, f'=S{r}/Conv/1000000', BU)
    put(wsDst, r, 33, f'=IF(X{r}="","",X{r}/Conv/1000000)', BU, font=BOLD)
D2 = D1 + len(EUR) - 1
SUBS = [('EU-27', 'EU'), ('Non-EU Europe', 'non-EU'), ('EUROPE TOTAL', None)]
SUBROW = {}
for k, (lab, key) in enumerate(SUBS):
    r = D2 + 1 + k; SUBROW[lab] = r
    put(wsDst, r, 1, lab, font=BOLD, fill=TOTFILL); put(wsDst, r, 2, '', fill=TOTFILL)
    for col in [3, 4, 5, 7, 9, 10, 11, 12, 15, 16, 17, 19, 23, 24, 28]:
        cl = L(col)
        f = f'=SUMIF($B${D1}:$B${D2},"{key}",{cl}{D1}:{cl}{D2})' if key else f'=SUM({cl}{D1}:{cl}{D2})'
        put(wsDst, r, col, f, MT, font=BOLD, fill=TOTFILL)
    put(wsDst, r, 20, f'=IF(S{r}=0,0,W{r}/S{r})', '0%', font=BOLD, fill=TOTFILL)
    put(wsDst, r, 25, f'=IF(W{r}=0,"",C{r}/W{r}-1)', PCT, font=BOLD, fill=TOTFILL)
    put(wsDst, r, 26, f'=IF(ABS(X{r})<Cargo_MT,"On pace",IF(X{r}>0,"ABOVE pace","BELOW pace"))', font=BOLD, fill=TOTFILL)
    put(wsDst, r, 27, f'=IF(S{r}=0,"",C{r}/S{r})', '0%', font=BOLD, fill=TOTFILL)
    put(wsDst, r, 29, f'=MAX(0,AB{r})/(52-Pace_Week)', MT, font=BOLD, fill=TOTFILL)
    put(wsDst, r, 30, f'=IF(N(T{r})=0,"",C{r}/T{r})', MT, font=BOLD, fill=TOTFILL)
    put(wsDst, r, 31, f'=IF(AD{r}="","",AD{r}/Conv/1000000)', BU, font=BOLD, fill=TOTFILL)
    put(wsDst, r, 32, f'=S{r}/Conv/1000000', BU, font=BOLD, fill=TOTFILL)
    put(wsDst, r, 33, f'=X{r}/Conv/1000000', BU, font=BOLD, fill=TOTFILL)
    for col in (6, 8, 13, 14, 18, 21, 22): put(wsDst, r, col, None, fill=TOTFILL)
text(wsDst, D2 + 5, 1, 'Subtotal pace = demand-weighted (expected ÷ demand); subtotal implied full year = actual ÷ that pace. Medians are not summed (see "Weekly Europe" for the Europe-level median).')
wsDst.column_dimensions['A'].width = 26; wsDst.column_dimensions['B'].width = 8
for col in range(3, len(DH) + 1): wsDst.column_dimensions[L(col)].width = 12
wsDst.column_dimensions['Z'].width = 12
for k in range(15): wsDst.column_dimensions[L(HB + k)].width = 9
wsDst.freeze_panes = 'C5'
for col in (7, 8, 24): posneg(wsDst, f'{L(col)}{D1}:{L(col)}{D2 + 3}')

# ================= History =================
title(wsHis, 'Full-MY US soybean exports to Europe by destination and marketing year (MT)',
      'From "Data RoW Full-Year" (ESR basis). US total includes China and Unknown re-declared at shipment.')
header(wsHis, 4, ['Destination', 'Group'] + PRIOR + ['5-yr avg (base yrs)', 'Floor refined', '5-yr high'])
for i, c in enumerate(EUR):
    r = 5 + i
    put(wsHis, r, 1, c); put(wsHis, r, 2, GRP[c])
    for k, m in enumerate(PRIOR): put(wsHis, r, 3 + k, '=' + fylook(cP0 + k, f'$A{r}'), MT, font=GREEN)
    for j, col in enumerate((cAvg, cMin2, cMax)): put(wsHis, r, 3 + len(PRIOR) + j, '=' + fylook(col, f'$A{r}'), MT, font=GREEN)
H2 = 4 + len(EUR)
hr = {}
for k, (lab, key) in enumerate(SUBS):
    r = H2 + 1 + k; hr[lab] = r
    put(wsHis, r, 1, lab, font=BOLD, fill=TOTFILL); put(wsHis, r, 2, '', fill=TOTFILL)
    for col in range(3, 3 + len(PRIOR) + 1):
        cl = L(col)
        f = f'=SUMIF($B$5:$B${H2},"{key}",{cl}5:{cl}{H2})' if key else f'=SUM({cl}5:{cl}{H2})'
        put(wsHis, r, col, f, MT, font=BOLD, fill=TOTFILL)
usr = H2 + 5
put(wsHis, usr, 1, 'US total full-MY exports (all destinations)', font=BOLD)
for k, m in enumerate(PRIOR):
    tot = float(ins[(ins.MY == m)].sort_values('date').groupby('Country').tail(1).AccumulatedExports.sum())
    put(wsHis, usr, 3 + k, tot, MT, font=BLUE)
put(wsHis, usr + 1, 1, 'Europe share of US exports', font=BOLD)
put(wsHis, usr + 2, 1, 'EU-27 share of US exports', font=BOLD)
for k in range(len(PRIOR)):
    cl = L(3 + k)
    put(wsHis, usr + 1, 3 + k, f'=IF({cl}{usr}=0,0,{cl}{hr["EUROPE TOTAL"]}/{cl}{usr})', '0.0%')
    put(wsHis, usr + 2, 3 + k, f'=IF({cl}{usr}=0,0,{cl}{hr["EU-27"]}/{cl}{usr})', '0.0%')
wsHis.column_dimensions['A'].width = 40; wsHis.column_dimensions['B'].width = 8
for col in range(3, 3 + len(PRIOR) + 3): wsHis.column_dimensions[L(col)].width = 13
wsHis.freeze_panes = 'C5'
bc = BarChart(); bc.type = 'col'; bc.grouping = 'stacked'; bc.overlap = 100
bc.title = 'Full-MY US soybean exports to Europe (MT)'; bc.y_axis.numFmt = '#,##0'
from openpyxl.chart import Series
for lab_ in ('EU-27', 'Non-EU Europe'):
    bc.series.append(Series(Reference(wsHis, min_col=3, max_col=2 + len(PRIOR), min_row=hr[lab_]), title=lab_))
bc.set_categories(Reference(wsHis, min_col=3, max_col=2 + len(PRIOR), min_row=4))
bc.width = 16; bc.height = 8; bc.x_axis.delete = False; bc.y_axis.delete = False
wsHis.add_chart(bc, f'A{usr + 5}')

# Europe full-year by base year (for aggregate pace)
def euro_fy(k, key=None):
    g = f',{FYR}!$B:$B,"{key}"' if key else f',{FYR}!$B:$B,"<>"'
    return f'SUMIFS({FYR}!${L(cB0+k)}:${L(cB0+k)}{g})'

# ================= Weekly Europe =================
title(wsWk, f'Europe weekly — cumulative commitment and weekly net sales, {CUR} vs {LY} vs 5-yr (MT)',
      'Net sales = week-on-week change in cumulative commitment (net of cancellations). Expected path = Europe demand used × average base-year pace '
      '(cumulative ÷ that year\'s Europe full-MY exports). Week 1 = opening book incl. pre-season sales.')
WH = ['Week', f'w/e {CUR}', f'Europe cum {CUR}', f'Europe cum {LY}', 'Europe cum 5-yr avg', 'Europe cum 5-yr median', 'Europe cum 5-yr min',
      'Europe cum 5-yr max', 'Expected path (demand × pace)', f'Europe net sales {CUR}', 'Europe net sales 5-yr avg', f'Europe net sales {LY}',
      f'EU-27 cum {CUR}', 'EU-27 cum 5-yr avg', f'EU-27 net sales {CUR}', 'EU-27 net sales 5-yr avg', 'Europe avg pace']
header(wsWk, 4, WH, h=52)
WB = len(WH) + 2
for k, h in enumerate([f'Europe cum base yr {k+1}' for k in range(5)] + [f'EU cum base yr {k+1}' for k in range(5)] + [f'Europe pace yr {k+1}' for k in range(5)]):
    x = wsWk.cell(4, WB + k, h); x.font = Font(name=F, size=8, bold=True, color='595959'); x.alignment = Alignment(wrap_text=True)
# Europe full-year per base year (row 3 helper)
text(wsWk, 3, WB - 1, 'Europe full-MY →', Font(name=F, size=8, italic=True))
for k in range(5): put(wsWk, 3, WB + 10 + k, '=' + euro_fy(k), MT, font=SMALL, border=False)
W1 = 5
d0 = ins[(ins.MY == CUR) & (ins.WeekNumber == 1)].date.min()
for w in range(1, NW + 1):
    r = W1 + w - 1
    put(wsWk, r, 1, w, '0'); put(wsWk, r, 2, (d0 + pd.Timedelta(days=7 * (w - 1))).to_pydatetime(), 'dd-mmm')
    cw = DCW(f'$A{r}')
    for k in range(5):
        put(wsWk, r, WB + k, f'=SUMIFS({cw},DC_MY,INDEX(Base_MY,{k+1}))', MT, font=SMALL, border=False)
        put(wsWk, r, WB + 5 + k, f'=SUMIFS({cw},DC_MY,INDEX(Base_MY,{k+1}),DC_G,"EU")', MT, font=SMALL, border=False)
        put(wsWk, r, WB + 10 + k, f'=IF({L(WB+10+k)}$3=0,"",{L(WB+k)}{r}/{L(WB+10+k)}$3)', '0%', font=SMALL, border=False)
    eb = f'{L(WB)}{r}:{L(WB+4)}{r}'; ub = f'{L(WB+5)}{r}:{L(WB+9)}{r}'
    if w <= WK:
        put(wsWk, r, 3, f'=SUMIFS({cw},DC_MY,CUR_MY)', MT, fill=CURFILL)
        put(wsWk, r, 10, f'=C{r}' if w == 1 else f'=C{r}-C{r-1}', MT, font=BOLD, fill=CURFILL)
        put(wsWk, r, 13, f'=SUMIFS({cw},DC_MY,CUR_MY,DC_G,"EU")', MT, fill=CURFILL)
        put(wsWk, r, 15, f'=M{r}' if w == 1 else f'=M{r}-M{r-1}', MT, font=BOLD, fill=CURFILL)
    put(wsWk, r, 4, f'=SUMIFS({cw},DC_MY,LY_MY)', MT)
    put(wsWk, r, 5, f'=AVERAGE({eb})', MT); put(wsWk, r, 6, f'=MEDIAN({eb})', MT)
    put(wsWk, r, 7, f'=MIN({eb})', MT); put(wsWk, r, 8, f'=MAX({eb})', MT)
    put(wsWk, r, 17, f'=IFERROR(AVERAGE({L(WB+10)}{r}:{L(WB+14)}{r}),0)', '0%')
    put(wsWk, r, 9, f"=Destinations!$S${SUBROW['EUROPE TOTAL']}*Q{r}", MT)
    put(wsWk, r, 11, f'=E{r}' if w == 1 else f'=E{r}-E{r-1}', MT)
    put(wsWk, r, 12, f'=D{r}' if w == 1 else f'=D{r}-D{r-1}', MT)
    put(wsWk, r, 14, f'=AVERAGE({ub})', MT)
    put(wsWk, r, 16, f'=N{r}' if w == 1 else f'=N{r}-N{r-1}', MT)
W2 = W1 + NW - 1
wsWk.column_dimensions['A'].width = 6; wsWk.column_dimensions['B'].width = 8
for col in range(3, len(WH) + 1): wsWk.column_dimensions[L(col)].width = 12
for k in range(15): wsWk.column_dimensions[L(WB + k)].width = 9
wsWk.freeze_panes = 'C5'
lc = LineChart(); lc.title = 'Europe cumulative commitment (MT)'; lc.y_axis.numFmt = '#,##0'
lc.add_data(Reference(wsWk, min_col=3, max_col=5, min_row=4, max_row=W2), titles_from_data=True)
lc.add_data(Reference(wsWk, min_col=9, max_col=9, min_row=4, max_row=W2), titles_from_data=True)
lc.set_categories(Reference(wsWk, min_col=1, min_row=W1, max_row=W2)); lc.width = 20; lc.height = 9
lc.x_axis.title = 'MY week'; lc.x_axis.delete = False; lc.y_axis.delete = False
wsWk.add_chart(lc, f'{L(WB+16)}4')
bc2 = BarChart(); bc2.title = 'Europe weekly net sales (MT)'; bc2.y_axis.numFmt = '#,##0'
bc2.add_data(Reference(wsWk, min_col=10, max_col=11, min_row=4, max_row=W2), titles_from_data=True)
bc2.set_categories(Reference(wsWk, min_col=1, min_row=W1, max_row=W2)); bc2.width = 20; bc2.height = 9
bc2.x_axis.delete = False; bc2.y_axis.delete = False
wsWk.add_chart(bc2, f'{L(WB+16)}24')

# ================= Nov-Feb Window =================
title(wsWin, 'Booking window — net sales and shipments to Europe between two MY weeks (default weeks 10–26 ≈ Nov–Feb)',
      'Net sales in window = cumulative commitment at end week − at (start week − 1). Shipments = same on accumulated exports. Edit the yellow weeks to move the window.')
put(wsWin, 4, 1, 'Window start week', font=BOLD); put(wsWin, 4, 2, 10, '0', font=BLUE, fill=YEL)
put(wsWin, 5, 1, 'Window end week', font=BOLD); put(wsWin, 5, 2, 26, '0', font=BLUE, fill=YEL)
put(wsWin, 6, 1, 'Weeks in window', font=BOLD); put(wsWin, 6, 2, '=B5-B4+1', '0')
name(wb, 'Win_S', wsWin.title, '$B$4'); name(wb, 'Win_E', wsWin.title, '$B$5')
WNH = ['Marketing year', 'EU-27 net sales', 'EU-27 shipments', 'EU-27 full-MY', 'EU-27 share of year booked in window',
       'Europe net sales', 'Europe shipments', 'Europe full-MY', 'Europe share of year booked in window', 'Europe avg weekly net sales in window']
header(wsWin, 8, WNH, h=52)
def win(field_w, my, grp=None):
    g = f',DC_G,"{grp}"' if grp else ''
    return f'(SUMIFS(INDEX({field_w},0,Win_E),DC_MY,{my}{g})-IF(Win_S<=1,0,SUMIFS(INDEX({field_w},0,Win_S-1),DC_MY,{my}{g})))'
for i, m in enumerate(PRIOR):
    r = 9 + i; my = f'$A{r}'
    put(wsWin, r, 1, m)
    put(wsWin, r, 2, '=' + win('DC_W', my, 'EU'), MT); put(wsWin, r, 3, '=' + win('DX_W', my, 'EU'), MT)
    put(wsWin, r, 4, f"=SUMIFS(INDEX({FYR}!${L(cP0)}:${L(cP0+len(PRIOR)-1)},0,MATCH({my},{FYR}!${L(cP0)}$8:${L(cP0+len(PRIOR)-1)}$8,0)),{FYR}!$B:$B,\"EU\")", MT)
    put(wsWin, r, 5, f'=IF(D{r}=0,"",B{r}/D{r})', '0%')
    put(wsWin, r, 6, '=' + win('DC_W', my), MT); put(wsWin, r, 7, '=' + win('DX_W', my), MT)
    put(wsWin, r, 8, f"=SUMIFS(INDEX({FYR}!${L(cP0)}:${L(cP0+len(PRIOR)-1)},0,MATCH({my},{FYR}!${L(cP0)}$8:${L(cP0+len(PRIOR)-1)}$8,0)),{FYR}!$B:$B,\"<>\")", MT)
    put(wsWin, r, 9, f'=IF(H{r}=0,"",F{r}/H{r})', '0%')
    put(wsWin, r, 10, f'=F{r}/$B$6', MT)
PL = 9 + len(PRIOR) - 1
rc = PL + 1
put(wsWin, rc, 1, f'{CUR} so far', font=BOLD, fill=CURFILL)
put(wsWin, rc, 2, f'=IF({WK}<Win_S,0,SUMIFS(INDEX(DC_W,0,MIN(Win_E,{WK})),DC_MY,CUR_MY,DC_G,"EU")-SUMIFS(INDEX(DC_W,0,Win_S-1),DC_MY,CUR_MY,DC_G,"EU"))', MT, fill=CURFILL)
put(wsWin, rc, 6, f'=IF({WK}<Win_S,0,SUMIFS(INDEX(DC_W,0,MIN(Win_E,{WK})),DC_MY,CUR_MY)-SUMIFS(INDEX(DC_W,0,Win_S-1),DC_MY,CUR_MY))', MT, fill=CURFILL)
for col in (3, 4, 5, 7, 8, 9, 10): put(wsWin, rc, col, None, fill=CURFILL)
put(wsWin, rc, 3, f'=IF({WK}<Win_S,"window not started (data to wk {WK})","")', fill=CURFILL)
bm = rc + 2
section(wsWin, bm, 'Benchmarks for the current window')
header(wsWin, bm + 1, ['Benchmark', 'EU-27 total in window', 'EU-27 per week', '', '', 'Europe total in window', 'Europe per week'], h=30)
bl = [('5-yr avg (base years)', 'AVERAGEIF'), ('5-yr avg excl. newest base yr', 'EXLY'), ('5-yr median', 'MEDIAN')]
def base_rows(col):
    return ','.join(f'INDEX({col}9:{col}{PL},MATCH(INDEX(Base_MY,{k+1}),$A$9:$A${PL},0))' for k in range(5))
for i, (lab, kind) in enumerate(bl):
    r = bm + 2 + i
    put(wsWin, r, 1, lab, font=BOLD)
    for col, cl in ((2, 'B'), (6, 'F')):
        args = base_rows(f'${cl}$')
        if kind == 'AVERAGEIF': f = f'=AVERAGE({args})'
        elif kind == 'MEDIAN': f = f'=MEDIAN({args})'
        else: f = '=AVERAGE(' + ','.join(f'INDEX(${cl}$9:${cl}${PL},MATCH(INDEX(Base_MY,{k+1}),$A$9:$A${PL},0))' for k in range(1, 5)) + ')'
        put(wsWin, r, col, f, MT); put(wsWin, r, col + 1, f'={L(col)}{r}/$B$6', MT, font=BOLD)
r = bm + 5
put(wsWin, r, 1, 'Needed in window for demand used (expected path)', font=BOLD, fill=CURFILL)
for col, sub in ((2, 'EU-27'), (6, 'EUROPE TOTAL')):
    if sub == 'EU-27':
        f = f"=Destinations!$S${SUBROW['EU-27']}*(INDEX('Weekly Europe'!$Q$5:$Q$56,Win_E)-IF(Win_S<=1,0,INDEX('Weekly Europe'!$Q$5:$Q$56,Win_S-1)))"
    else:
        f = f"=INDEX('Weekly Europe'!$I$5:$I$56,Win_E)-IF(Win_S<=1,0,INDEX('Weekly Europe'!$I$5:$I$56,Win_S-1))"
    put(wsWin, r, col, f, MT, fill=CURFILL); put(wsWin, r, col + 1, f'={L(col)}{r}/$B$6', MT, font=BOLD, fill=CURFILL)
text(wsWin, r + 1, 1, 'EU-27 "needed" applies the Europe-wide pace curve to EU-27 demand (EU-27 is ~95% of Europe, so the curves are near-identical).')
wsWin.column_dimensions['A'].width = 44
for col in range(2, 11): wsWin.column_dimensions[L(col)].width = 14
bw = BarChart(); bw.title = 'EU-27: net sales and shipments in window (MT)'; bw.y_axis.numFmt = '#,##0'
bw.add_data(Reference(wsWin, min_col=2, max_col=3, min_row=8, max_row=PL), titles_from_data=True)
bw.set_categories(Reference(wsWin, min_col=1, min_row=9, max_row=PL)); bw.width = 18; bw.height = 8
bw.x_axis.delete = False; bw.y_axis.delete = False
wsWin.add_chart(bw, 'L4')
lw = LineChart(); lw.title = 'Share of the year booked in window'; lw.y_axis.numFmt = '0%'
lw.add_data(Reference(wsWin, min_col=5, max_col=5, min_row=8, max_row=PL), titles_from_data=True)
lw.add_data(Reference(wsWin, min_col=9, max_col=9, min_row=8, max_row=PL), titles_from_data=True)
lw.set_categories(Reference(wsWin, min_col=1, min_row=9, max_row=PL)); lw.width = 18; lw.height = 7
lw.x_axis.delete = False; lw.y_axis.delete = False
wsWin.add_chart(lw, 'L22')

# ================= Country Chart =================
title(wsCC, 'Destination pace chart — pick a European destination in B4',
      'Expected path = demand used (Destinations tab) × average base-year pace by week. Actual = cumulative total commitment.')
put(wsCC, 4, 1, 'Destination', font=BOLD); put(wsCC, 4, 2, 'GERMANY', font=BLUE, fill=YEL)
dvc = DataValidation(type='list', formula1=f'=Destinations!$A${D1}:$A${D2}'); wsCC.add_data_validation(dvc); dvc.add('B4')
put(wsCC, 5, 1, 'Demand used (MT)', font=BOLD); put(wsCC, 5, 2, f'=INDEX(Destinations!$S:$S,MATCH($B$4,Destinations!$A:$A,0))', MT, font=GREEN)
put(wsCC, 6, 1, 'Status at week N', font=BOLD); put(wsCC, 6, 2, f'=INDEX(Destinations!$Z:$Z,MATCH($B$4,Destinations!$A:$A,0))', font=GREEN)
put(wsCC, 7, 1, f'Unshipped {CUR} (MT)', font=BOLD); put(wsCC, 7, 2, f'=INDEX(Destinations!$J:$J,MATCH($B$4,Destinations!$A:$A,0))', MT, font=GREEN)
CH = ['Week', 'Expected path', f'Actual {CUR}', f'{LY}', '5-yr avg', 'Avg pace'] + [f'Cum base yr {k+1}' for k in range(5)] + [f'FY base yr {k+1}' for k in range(5)] + [f'Pace yr {k+1}' for k in range(5)]
header(wsCC, 9, CH, h=40)
for k in range(5):
    put(wsCC, 8, 12 + k, f'=INDEX(Destinations!${L(HB+5+k)}:${L(HB+5+k)},MATCH($B$4,Destinations!$A:$A,0))', MT, font=SMALL, border=False)
for w in range(1, NW + 1):
    r = 9 + w; cw = DCW(f'$A{r}')
    put(wsCC, r, 1, w, '0')
    for k in range(5):
        put(wsCC, r, 7 + k, f'=SUMIFS({cw},DC_D,$B$4,DC_MY,INDEX(Base_MY,{k+1}))', MT, font=SMALL, border=False)
        put(wsCC, r, 12 + k, f'={L(12+k)}$8', MT, font=SMALL, border=False)
        put(wsCC, r, 17 + k, f'=IF({L(12+k)}{r}=0,"",{L(7+k)}{r}/{L(12+k)}{r})', '0%', font=SMALL, border=False)
    put(wsCC, r, 6, f'=IFERROR(AVERAGE(Q{r}:U{r}),0)', '0%')
    put(wsCC, r, 2, f'=$B$5*F{r}', MT)
    if w <= WK: put(wsCC, r, 3, f'=SUMIFS({cw},DC_D,$B$4,DC_MY,CUR_MY)', MT, fill=CURFILL)
    put(wsCC, r, 4, f'=SUMIFS({cw},DC_D,$B$4,DC_MY,LY_MY)', MT)
    put(wsCC, r, 5, f'=AVERAGE(G{r}:K{r})', MT)
wsCC.column_dimensions['A'].width = 22
for col in range(2, 22): wsCC.column_dimensions[L(col)].width = 11
lcc = LineChart(); lcc.title = 'Selected destination: cumulative commitment vs expected (MT)'; lcc.y_axis.numFmt = '#,##0'
lcc.add_data(Reference(wsCC, min_col=2, max_col=5, min_row=9, max_row=9 + NW), titles_from_data=True)
lcc.set_categories(Reference(wsCC, min_col=1, min_row=10, max_row=9 + NW)); lcc.width = 20; lcc.height = 10
lcc.x_axis.title = 'MY week'; lcc.x_axis.delete = False; lcc.y_axis.delete = False
wsCC.add_chart(lcc, 'W4')

# ================= Market View =================
MKT_MONTHS = ['Sep', 'Oct', 'Nov', 'Dec', 'Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug']
MKT = {  # kt; user-provided market table: imports of US soybeans by month into EU-27 + UK (market's 'EU' = old EU-28); last row = market estimate
    '2016/17': [122, 796, 591, 470, 689, 322, 384, 457, 7, 191, 493, 219],
    '2017/18': [219, 606, 604, 370, 561, 482, 275, 277, 511, 831, 658, 1085],
    '2018/19': [723, 1406, 1499, 1337, 1159, 610, 295, 415, 152, 218, 283, 296],
    '2019/20': [819, 960, 493, 608, 878, 356, 244, 114, 128, 7, 297, 851],
    '2020/21': [528, 856, 688, 740, 1091, 717, 31, 26, 18, 21, 108, 169],
    '2021/22': [393, 812, 847, 1008, 809, 292, 81, 160, 250, 310, 354, 557],
    '2022/23': [258, 662, 875, 642, 897, 583, 352, 455, 151, 186, 727, 608],
    '2023/24': [485, 672, 1171, 1177, 680, 557, 299, 33, 62, 333, 302, 293],
    '2024/25': [604, 1046, 1368, 866, 775, 352, 288, 390, 6, 243, 364, 336],
    '2025/26': [530, 1033, 1027, 470, 697, 284, 73, 12, 156, 219, 195, 340],
    CUR:       [480, 780, 1180, 820, 660, 510, 210, 40, 30, 50, 20, 115],
}
wsM = wb.create_sheet('Market View', index=1)
title(wsM, f'Market view — EU imports of US soybeans by month (kt) and how the market\'s {CUR} number compares with our model',
      'Source: market table provided by the user — monthly imports of US soybeans (kt) into the market\'s "EU" = EU-27 + UK (old EU-28); the last row is the market\'s estimate for the current season. '
      'This is a different basis from USDA export sales (arrivals vs US sales/shipments), so it is bridged with a conversion ratio in section C.')
section(wsM, 4, 'A) Monthly table (blue = input; edit the current-season row to update)')
header(wsM, 5, ['Season'] + MKT_MONTHS + ['Total', 'Sep–Feb', 'Mar–Aug', 'Mar–Aug share'], h=30)
MS = list(MKT.keys())
for i, sea in enumerate(MS):
    r = 6 + i
    put(wsM, r, 1, sea, font=BOLD, fill=CURFILL if sea == CUR else None)
    for j, v in enumerate(MKT[sea]): put(wsM, r, 2 + j, v, MT, font=BLUE, fill=YEL if sea == CUR else None)
    put(wsM, r, 14, f'=SUM(B{r}:M{r})', MT, font=BOLD); put(wsM, r, 15, f'=SUM(B{r}:G{r})', MT); put(wsM, r, 16, f'=SUM(H{r}:M{r})', MT)
    put(wsM, r, 17, f'=IF(N{r}=0,0,P{r}/N{r})', '0%')
ML = 6 + len(MS) - 1; MCR = ML
MR = f'$A$6:$A${ML}'
def mk(col, sea_ref): return f'INDEX(${col}$6:${col}${ML},MATCH({sea_ref},{MR},0))'
ra_ = ML + 1
put(wsM, ra_, 1, '5-yr avg (base yrs)', font=BOLD, fill=TOTFILL)
for j in range(2, 17):
    cl = L(j)
    put(wsM, ra_, j, '=(' + '+'.join(mk(cl, f'INDEX(Base_MY,{k+1})') for k in range(5)) + ')/5', MT, font=BOLD, fill=TOTFILL)
put(wsM, ra_, 17, f'=P{ra_}/N{ra_}', '0%', font=BOLD, fill=TOTFILL)
rs_ = ra_ + 1
put(wsM, rs_, 1, f'{CUR} vs 5-yr avg', font=BOLD)
for j in range(2, 17): put(wsM, rs_, j, f'=IF({L(j)}{ra_}=0,"",{L(j)}{MCR}/{L(j)}{ra_}-1)', PCT, font=BOLD)
posneg(wsM, f'B{rs_}:P{rs_}')
rsh = rs_ + 1
put(wsM, rsh, 1, f'Monthly share {CUR}', font=NORM)
put(wsM, rsh + 1, 1, 'Monthly share 5-yr avg', font=NORM)
for j in range(2, 14):
    put(wsM, rsh, j, f'=IF($N${MCR}=0,0,{L(j)}{MCR}/$N${MCR})', '0%')
    put(wsM, rsh + 1, j, f'=IF($N${ra_}=0,0,{L(j)}{ra_}/$N${ra_})', '0%')

b0 = rsh + 3
section(wsM, b0, f'B) What the market\'s {CUR} number assumes')
kb = [('Market estimate for the season (kt)', f'=N{MCR}', MT),
      ('vs last season', f'=N{MCR}/{mk("N", "LY_MY")}-1', PCT),
      ('vs 5-yr avg', f'=N{MCR}/N{ra_}-1', PCT),
      ('Sep–Feb vs 5-yr avg', f'=O{MCR}/O{ra_}-1', PCT),
      ('Mar–Aug vs 5-yr avg', f'=P{MCR}/P{ra_}-1', PCT),
      ('Mar–Aug share of the year: market vs 5-yr avg', f'=TEXT(Q{MCR},"0%")&" vs "&TEXT(Q{ra_},"0%")', None)]
for i, (lab, f, fmt) in enumerate(kb):
    put(wsM, b0 + 1 + i, 1, lab, font=BOLD); put(wsM, b0 + 1 + i, 2, f, fmt, font=Font(name=F, size=11, bold=True, color='1F4E78'))
text(wsM, b0 + 7, 1, 'Reading: the market keeps Sep–Feb close to a normal year and cuts almost all of Mar–Aug, i.e. it assumes the US is priced out of the EU '
                     'once Brazil\'s crop arrives. EUDR (if applied from 30 Dec 2026, US low-risk) is the main risk to that assumption.')

c0 = b0 + 9
UKR = D1 + EUR.index('UNITED KINGDOM')
section(wsM, c0, 'C) Bridge to our model (USDA export-sales basis, EU-27 + UK = same countries as the market table)')
header(wsM, c0 + 1, ['Base season', 'Market table (kt)', 'Export sales, EU-27 + UK full-MY (kt)', 'Ratio (export sales ÷ market)'], h=40)
for k in range(5):
    r = c0 + 2 + k
    put(wsM, r, 1, f'=INDEX(Base_MY,{k+1})', font=GREEN)
    put(wsM, r, 2, '=' + mk('N', f'$A{r}'), MT)
    put(wsM, r, 3, f"=SUMIFS(INDEX({FYR}!${L(cP0)}:${L(cP0+len(PRIOR)-1)},0,MATCH($A{r},{FYR}!${L(cP0)}$8:${L(cP0+len(PRIOR)-1)}$8,0)),{FYR}!$B:$B,\"EU\")/1000+SUMIFS(INDEX({FYR}!${L(cP0)}:${L(cP0+len(PRIOR)-1)},0,MATCH($A{r},{FYR}!${L(cP0)}$8:${L(cP0+len(PRIOR)-1)}$8,0)),{FYR}!$A:$A,\"UNITED KINGDOM\")/1000", MT)
    put(wsM, r, 4, f'=IF(B{r}=0,"",C{r}/B{r})', '0.00')
rq = c0 + 7
put(wsM, rq, 1, 'Average ratio', font=BOLD, fill=TOTFILL); put(wsM, rq, 4, f'=AVERAGE(D{c0+2}:D{c0+6})', '0.00', font=BOLD, fill=TOTFILL)
put(wsM, rq, 2, None, fill=TOTFILL); put(wsM, rq, 3, None, fill=TOTFILL)
br = [('Market estimate on export-sales basis (kt)', f'=N{MCR}*D{rq}', MT),
      ('Our model: EU-27 + UK demand used (kt)', f"=(Destinations!S{SUBROW['EU-27']}+Destinations!S{UKR})/1000", MT),
      ('Difference: market − our model (kt)', f'=B{rq+1}-B{rq+2}', '+#,##0;-#,##0;0'),
      ('Difference (mln bu)', f'=B{rq+3}*1000/Conv/1000000', '+#,##0.0;-#,##0.0;0.0'),
      ('Market vs its own 5-yr avg', f'=N{MCR}/N{ra_}-1', PCT),
      ('Our model vs export-sales 5-yr avg (EU-27 + UK)', f"=(Destinations!S{SUBROW['EU-27']}+Destinations!S{UKR})/(Destinations!L{SUBROW['EU-27']}+Destinations!L{UKR})-1", PCT)]
for i, (lab, f, fmt) in enumerate(br):
    put(wsM, rq + 1 + i, 1, lab, font=BOLD); put(wsM, rq + 1 + i, 2, f, fmt, font=Font(name=F, size=11, bold=True, color='1F4E78'))
p0 = rq + 8
section(wsM, p0, 'D) Are EU-27 + UK bookings on pace for the market\'s number?')
pk = [('EU-27 + UK booked so far, export-sales basis (kt)', f"=(Destinations!C{SUBROW['EU-27']}+Destinations!C{UKR})/1000", MT),
      ('Share usually booked by now (EU-27 + UK, demand-weighted)', f"=(Destinations!W{SUBROW['EU-27']}+Destinations!W{UKR})/(Destinations!S{SUBROW['EU-27']}+Destinations!S{UKR})", '0%'),
      ('Should have booked for the market number (kt)', f'=B{rq+1}*B{p0+2}', MT),
      ('Gap (kt)', f'=B{p0+1}-B{p0+3}', '+#,##0;-#,##0;0'),
      ('Status', f'=IF(ABS(B{p0+4})*1000<Cargo_MT,"On pace",IF(B{p0+4}>0,"AHEAD of the market number","BEHIND the market number"))', None),
      ('Same, for our model\'s demand: gap (kt)', f"=(Destinations!X{SUBROW['EU-27']}+Destinations!X{UKR})/1000", '+#,##0;-#,##0;0')]
for i, (lab, f, fmt) in enumerate(pk):
    put(wsM, p0 + 1 + i, 1, lab, font=BOLD); put(wsM, p0 + 1 + i, 2, f, fmt, font=Font(name=F, size=11, bold=True, color='1F4E78'))
text(wsM, p0 + 8, 1, 'Caveat: the ratio bridges two different datasets (arrivals in the EU vs US sales/shipments, with timing and transshipment differences); '
                     'it varies by season (see col D above), so treat the bridged number as ± a few hundred kt.')
wsM.column_dimensions['A'].width = 50
for j in range(2, 18): wsM.column_dimensions[L(j)].width = 10
wsM.column_dimensions['B'].width = 14; wsM.column_dimensions['C'].width = 14
mc = BarChart(); mc.type = 'col'; mc.title = f'Monthly share of the year: market {CUR} vs 5-yr avg'; mc.y_axis.numFmt = '0%'
from openpyxl.chart import Series
mc.series.append(Series(Reference(wsM, min_col=2, max_col=13, min_row=rsh), title=f'Market {CUR}'))
mc.series.append(Series(Reference(wsM, min_col=2, max_col=13, min_row=rsh + 1), title='5-yr avg'))
mc.set_categories(Reference(wsM, min_col=2, max_col=13, min_row=5)); mc.width = 18; mc.height = 8
mc.x_axis.delete = False; mc.y_axis.delete = False
wsM.add_chart(mc, f'F{b0}')

# ================= Dashboard body =================
r0 = 18
section(wsH, r0, f'A) Europe at week N — {CUR} vs history and vs demand')
DBH = ['Group', f'{CUR} wk N', f'{LY} wk N', '5-yr avg wk N', 'Δ vs LY %', 'Δ vs 5-yr %', 'Demand used (MT)', 'Demand used (mln bu)',
       'Expected wk N', 'Δ vs expected', 'Status', '% of demand committed', 'Remaining to sell', 'Weekly needed rest of MY',
       'Implied full year (MT)', 'Implied full year (mln bu)', 'Implied vs demand (mln bu)']
header(wsH, r0 + 1, DBH, h=52)
for i, (lab, _) in enumerate(SUBS):
    r = r0 + 2 + i; s = SUBROW[lab]
    fill = TOTFILL if lab == 'EUROPE TOTAL' else None
    put(wsH, r, 1, lab, font=BOLD, fill=fill)
    for col, src_ in ((2, 'C'), (3, 'D'), (4, 'E'), (7, 'S'), (9, 'W'), (10, 'X'), (13, 'AB'), (14, 'AC'), (15, 'AD')):
        put(wsH, r, col, f'=Destinations!{src_}{s}', MT, font=BOLD if col == 10 else GREEN, fill=fill)
    put(wsH, r, 5, '=' + dpct(f'B{r}', f'C{r}'), PCT, fill=fill); put(wsH, r, 6, '=' + dpct(f'B{r}', f'D{r}'), PCT, fill=fill)
    put(wsH, r, 8, f'=G{r}/Conv/1000000', BU, fill=fill)
    put(wsH, r, 11, f'=Destinations!Z{s}', font=BOLD, fill=fill)
    put(wsH, r, 12, f'=Destinations!AA{s}', '0%', fill=fill)
    put(wsH, r, 16, f'=IF(O{r}="","",O{r}/Conv/1000000)', BU, fill=fill)
    put(wsH, r, 17, f'=IF(P{r}="","",P{r}-H{r})', '+#,##0.0;-#,##0.0;0.0', font=BOLD, fill=fill)
posneg(wsH, f'J{r0+2}:J{r0+4}'); posneg(wsH, f'Q{r0+2}:Q{r0+4}')

r1 = r0 + 7
section(wsH, r1, 'B) Weekly net-sales monitor (Europe total)')
header(wsH, r1 + 1, ['Measure', f'{CUR}', '5-yr avg same weeks', 'Δ', f'{LY} same weeks'], h=30)
WKR = "'Weekly Europe'"
mon = [('Latest week net sales (week N)', f'INDEX({WKR}!$J$5:$J$56,Pace_Week)', f'INDEX({WKR}!$K$5:$K$56,Pace_Week)', f'INDEX({WKR}!$L$5:$L$56,Pace_Week)'),
       ('Average of last 4 weeks', f'AVERAGE(OFFSET({WKR}!$J$5,Pace_Week-4,0,4,1))',
        f'AVERAGE(OFFSET({WKR}!$K$5,Pace_Week-4,0,4,1))',
        f'AVERAGE(OFFSET({WKR}!$L$5,Pace_Week-4,0,4,1))'),
       ('Cumulative commitment at week N', f'INDEX({WKR}!$C$5:$C$56,Pace_Week)', f'INDEX({WKR}!$E$5:$E$56,Pace_Week)', f'INDEX({WKR}!$D$5:$D$56,Pace_Week)')]
for i, (lab, a, b_, c_) in enumerate(mon):
    r = r1 + 2 + i
    put(wsH, r, 1, lab, font=BOLD)
    put(wsH, r, 2, f'=IF(Pace_Week>{WK},"",{a})', MT, fill=CURFILL); put(wsH, r, 3, f'={b_}', MT)
    put(wsH, r, 4, f'=IF(B{r}="","",B{r}-C{r})', MT, font=BOLD); put(wsH, r, 5, f'={c_}', MT)
posneg(wsH, f'D{r1+2}:D{r1+4}')
r = r1 + 5
put(wsH, r, 1, 'Window benchmarks (weekly net sales, Europe) →', font=BOLD)
WN = "'Nov-Feb Window'"
put(wsH, r + 1, 1, '  Normal (5-yr avg excl. newest base yr)'); put(wsH, r + 1, 2, f'={WN}!G{bm+3}', MT, font=GREEN)
put(wsH, r + 2, 1, '  5-yr avg (all base yrs)'); put(wsH, r + 2, 2, f'={WN}!G{bm+2}', MT, font=GREEN)
put(wsH, r + 3, 1, '  Needed for demand used'); put(wsH, r + 3, 2, f'={WN}!G{bm+5}', MT, font=GREEN)
put(wsH, r + 4, 1, '  Weak-year marker (lowest base yr in window)')
put(wsH, r + 4, 2, '=MIN(' + ','.join(f"INDEX({WN}!$J$9:$J${PL},MATCH(INDEX(Base_MY,{k+1}),{WN}!$A$9:$A${PL},0))" for k in range(5)) + ')', MT, font=GREEN)
text(wsH, r + 1, 3, 'Source: "Nov-Feb Window" tab (window weeks set there; default 10–26).')

r2 = r + 7
section(wsH, r2, 'C) Key destinations')
KH = ['Destination', f'{CUR} wk N', 'Expected wk N', 'Δ vs expected', 'Status', f'Unshipped {CUR}', 'Demand used', 'Pace by wk N', 'Weekly needed rest of MY']
header(wsH, r2 + 1, KH, h=40)
keys = [c for c in EUR if avg5[c] >= 100000][:8]
for i, c in enumerate(keys):
    r = r2 + 2 + i; dr = D1 + EUR.index(c)
    put(wsH, r, 1, c, font=BOLD)
    for col, src_, fmt in ((2, 'C', MT), (3, 'W', MT), (4, 'X', MT), (5, 'Z', None), (6, 'J', MT), (7, 'S', MT), (8, 'T', '0%'), (9, 'AC', MT)):
        put(wsH, r, col, f'=Destinations!{src_}{dr}', fmt, font=BOLD if col in (4, 5) else GREEN)
posneg(wsH, f'D{r2+2}:D{r2+1+len(keys)}')

r3 = r2 + 3 + len(keys)
section(wsH, r3, 'D) What to watch (analyst notes — not data)')
for i, t in enumerate([
    'Booking window: Europe usually books roughly half of its US year in weeks 10–26 (Nov–Feb), before Brazil\'s harvest. Compare weekly net sales with the benchmarks in B).',
    'Crush hubs: the Netherlands and Germany. A build-up of unshipped sales there is the first sign of demand returning; zero unshipped = hand-to-mouth buying.',
    'Feed markets: Spain, Portugal and Italy hold most of Europe\'s unshipped book.',
    'EUDR (EU deforestation regulation): our understanding as of mid-2026 is that obligations for large operators apply from 30 Dec 2026, with the US benchmarked low-risk '
    'and Brazil standard-risk. If so, EU buyers may shift some Jan–Aug 2027 purchases to US origin, which would break the usual seasonal pace (the pace model would then '
    'understate Europe). Verify the current EUDR status before relying on this.',
    'Counter-force: China filling its US purchase commitment frees Brazilian supply, which can make Brazil cheaper into Europe in H1 2027.',
    'Basis: ESR data; USDA\'s total is Census basis (usually higher). Small destinations (1–3 cargoes/yr) have erratic pace — rely on the large ones and the totals.']):
    text(wsH, r3 + 1 + i, 1, t)
wsH.column_dimensions['A'].width = 46
for col in range(2, 18): wsH.column_dimensions[L(col)].width = 13
wsH.column_dimensions['C'].width = 15

# ================= Method & Notes =================
title(wsN, 'Method and notes', 'How each number in this workbook is built.')
notes = [
    ('Data', 'USDA FAS Export Sales Reporting (ESR), commodity SOYBEANS. Only "in-season" rows are used (week-ending dates within the marketing year); '
             'the same week numbers also appear for next-MY sales in the prior reporting year, and those are excluded.'),
    ('Total commitment', 'Accumulated exports (shipped) + outstanding sales (unshipped) for the marketing year, as of the report week.'),
    ('Full-MY exports', 'Accumulated exports on the last in-season report of each MY (≈ week 53). ESR basis; USDA balance-sheet exports are Census basis.'),
    ('Comparisons', 'Every MY is compared at the same in-season week number (report dates differ by up to ±4 days).'),
    ('RoW demand', 'USDA total (mln bu × 0.0272155) minus the China assumption. Europe receives a share of this, set by the allocation method.'),
    ('One factor', 'Every RoW destination = its 5-yr average × (RoW target ÷ Σ RoW 5-yr averages). Same % cut for all.'),
    ('Floor+swing', 'Floor = lowest base-yr volume; swing = 5-yr average − floor. Pool = RoW target − Σ floors. Each destination = floor + swing × (pool ÷ Σ swings). '
                    'Steady buyers are cut little, volatile buyers absorb most of a lower total.'),
    ('Refined (default)', 'As floor+swing, but the floor is the 2nd-lowest base year (one outlier year cannot set it) and demand is capped at the 5-yr high.'),
    ('Override', 'A value in "Destinations" col R replaces the method for that destination (the RoW total then no longer ties exactly to the target).'),
    ('Seasonal pace', 'For each base year: commitment at week N ÷ that year\'s full-MY exports; averaged across base years (years with zero exports skipped).'),
    ('Expected / Δ', 'Expected = demand used × pace. Status = ABOVE/BELOW when |Δ| ≥ cargo size, otherwise "On pace".'),
    ('Implied full year', 'Actual commitment ÷ pace. Mechanical; noisy for small buyers. Group totals use demand-weighted pace.'),
    ('Booking window', 'Net sales between two weeks = change in cumulative commitment; shipments = change in accumulated exports.'),
    ('Groups', 'EU = EU-27 members present in the data; non-EU Europe = UK, Norway, Switzerland, etc.'),
    ('Caveats', 'Five base years is a short history. 2025/26 (China absent from US purchases) distorts several averages; use the base-year inputs to test without it. '
                'Policy events (EUDR, trade measures) can break seasonal patterns.'),
]
for i, (k, v) in enumerate(notes):
    put(wsN, 4 + i, 1, k, font=BOLD); x = put(wsN, 4 + i, 2, v); x.alignment = Alignment(wrap_text=True, vertical='top')
    wsN.row_dimensions[4 + i].height = 42
wsN.column_dimensions['A'].width = 20; wsN.column_dimensions['B'].width = 130

for ws in wb:
    ws.sheet_view.showGridLines = False
    for ch_ in ws._charts:
        if isinstance(ch_, LineChart):
            for s_ in ch_.series: s_.smooth = False
wb.save(out)
print('CUR', CUR, 'WK', WK, 'Europe destinations', len(EUR), 'RoW destinations', len(ROW))
