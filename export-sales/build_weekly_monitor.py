"""Weekly monitor: week-by-week sales (by region) and shipments needed for USDA's export number and for the desk number.
Usage: python build_weekly_monitor.py <ESR csv (converted)> <out.xlsx>
Method: China fixed (input, default 25 Mt). Rest of world must book (target on ESR basis − China − RoW booked now); that requirement is
shared across regions in proportion to each region's usual (5-yr) bookings from this week to season end, and spread over the weeks with
each region's usual 5-yr seasonal shape. Shipments: all destinations, accumulated exports must reach the target (ESR basis)."""
import sys
import numpy as np
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.chart import LineChart, Reference
from openpyxl.utils import get_column_letter as L
from openpyxl.formatting.rule import CellIsRule, FormulaRule
from openpyxl.workbook.defined_name import DefinedName

src, out = sys.argv[1], sys.argv[2]
d = pd.read_csv(src, thousands=','); d.columns = [c.strip() for c in d.columns]
d['C'] = d.CountryName.str.strip(); d['dt'] = pd.to_datetime(d.WeekendingDate)
d['Y'] = d.ReportingMarketingYearName.str.extract(r'Sep (\d{4})').astype(int)
d = d[(d.dt >= pd.to_datetime(d.Y.astype(str) + '-08-31')) & (d.dt <= pd.to_datetime((d.Y + 1).astype(str) + '-09-07'))]
cy = int(d.loc[d.dt.idxmax(), 'Y']); WK = int(d[d.Y == cy].WeekNumber.max()); LAST = d[d.Y == cy].dt.max()
CUR = f'{cy}/{str(cy+1)[-2:]}'; B5 = list(range(cy - 5, cy)); NW = 53
RMAP = {'CHINA, PEOPLES REPUBLIC OF': 'China', 'MEXICO': 'Mexico', 'UNKNOWN': 'Unknown'}
for lst, nm in [(['JAPAN', 'TAIWAN', 'KOREA, REPUBLIC OF', 'HONG KONG'], 'East Asia ex-China'),
                (['INDONESIA', 'VIETNAM', 'MALAYSIA', 'THAILAND', 'PHILIPPINES', 'BURMA', 'CAMBODIA', 'LAOS', 'SINGAPORE'], 'Southeast Asia'),
                (['BANGLADESH', 'PAKISTAN', 'NEPAL', 'SRI LANKA', 'INDIA'], 'South Asia'),
                (['EGYPT', 'ALGERIA', 'TURKEY', 'TUNISIA', 'MOROCCO', 'IRAN', 'IRAQ', 'ISRAEL', 'JORDAN', 'LEBANON', 'LIBYA', 'SAUDI ARABIA', 'QATAR',
                  'UNITED ARAB EMIRATES'], 'Middle East & N. Africa'),
                (['SPAIN', 'NETHERLANDS', 'PORTUGAL', 'ITALY', 'GERMANY', 'UNITED KINGDOM', 'GREECE', 'BELGIUM-LUXEMBOURG', 'FRANCE', 'IRELAND', 'NORWAY',
                  'ROMANIA', 'SWITZERLAND', 'UKRAINE', 'GEORGIA, REPUBLIC OF'], 'Europe')]:
    for c in lst: RMAP[c] = nm
d['R'] = d.C.map(RMAP).fillna('Other Americas, Africa & rest')
REGS = ['Mexico', 'East Asia ex-China', 'Southeast Asia', 'South Asia', 'Middle East & N. Africa', 'Europe', 'Other Americas, Africa & rest']
cum = d.pivot_table(index=['R', 'WeekNumber'], columns='Y', values='TotalCommitment', aggfunc='sum').fillna(0) / 1e6
fin = d[d.Y < cy].sort_values('dt').groupby(['Y', 'C']).tail(1).groupby(['R', 'Y']).AccumulatedExports.sum().unstack().fillna(0) / 1e6
def ser(r, y): return cum.loc[r][y] if r in cum.index.get_level_values(0) else pd.Series(0, index=range(1, NW + 1))
shape, usual, now_r = {}, {}, {}
for r in REGS + ['China']:
    yrs = B5 if r != 'China' else [y for y in B5 if y != cy - 1]   # China: drop 2025/26 (China absent)
    rem = {y: fin.loc[r, y] - ser(r, y).get(WK, 0) for y in yrs}
    tot_rem = sum(rem.values())   # ratio of totals across seasons: stable when one season has almost nothing left to book
    shape[r] = {w: sum(ser(r, y).get(w, ser(r, y).max()) - ser(r, y).get(WK, 0) for y in yrs) / tot_rem for w in range(WK, NW + 1)}
    usual[r] = np.mean([rem[y] for y in yrs]); now_r[r] = float(ser(r, cy).get(WK, 0))
unk_now = float(ser('Unknown', cy).get(WK, 0))
avg_wk = {r: {w: np.mean([ser(r, y).get(w, 0) - ser(r, y).get(w - 1, 0) for y in B5]) for w in range(WK + 1, NW + 1)} for r in REGS + ['China']}
exp = d.groupby(['Y', 'WeekNumber']).AccumulatedExports.sum().unstack(0) / 1e6
finT = d[d.Y < cy].sort_values('dt').groupby(['Y', 'C']).tail(1).groupby('Y').AccumulatedExports.sum() / 1e6
ship_now = float(exp.loc[WK, cy])
sshape = {w: np.mean([(exp.loc[w, y] - exp.loc[WK, y]) / (finT[y] - exp.loc[WK, y]) for y in B5]) for w in range(WK, NW + 1)}
ship_avg = {w: np.mean([exp.loc[w, y] - exp.loc[w - 1, y] for y in B5]) for w in range(WK + 1, NW + 1)}

# ---------------- styles ----------------
F = 'Arial'; NAVY = '1F4E78'
HF = PatternFill('solid', fgColor=NAVY); TOT = PatternFill('solid', fgColor='EDEDED'); YEL = PatternFill('solid', fgColor='FFFF00')
CURF = PatternFill('solid', fgColor='FFF2CC'); GRN = PatternFill('solid', fgColor='C6EFCE'); RED = PatternFill('solid', fgColor='FFC7CE')
AMB = PatternFill('solid', fgColor='FFEB9C')
thin = Side(style='thin', color='D9D9D9'); BOX = Border(bottom=thin)
KT = '#,##0'; MT = '0.00'
def font(sz=10, b=False, col='000000', i=False): return Font(name=F, size=sz, bold=b, color=col, italic=i)
BLUE = font(10, False, '0000FF')
def put(ws, r, c, v, fmt=None, f=None, fill=None, al=None):
    x = ws.cell(r, c, v); x.font = f or font(); x.border = BOX
    if fmt: x.number_format = fmt
    if fill: x.fill = fill
    if al: x.alignment = al
    return x
def header(ws, r, labels, c0=1, h=48):
    for i, t in enumerate(labels):
        x = ws.cell(r, c0 + i, t); x.font = font(9, True, 'FFFFFF'); x.fill = HF
        x.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    ws.row_dimensions[r].height = h
def title(ws, t, sub):
    ws['A1'] = t; ws['A1'].font = font(14, True, NAVY); ws['A2'] = sub; ws['A2'].font = font(9, False, '595959', True)
def name(wb, n, sh, ref): wb.defined_names[n] = DefinedName(n, attr_text=f"'{sh}'!{ref}")

wb = Workbook()
wsI = wb.active; wsI.title = 'Inputs'
wsA = wb.create_sheet('Weekly targets USDA')
wsB = wb.create_sheet('Weekly targets desk')
wsT = wb.create_sheet('Tracker')
wsS = wb.create_sheet('Shapes (data)')

# ---------------- Inputs ----------------
title(wsI, f'Weekly monitor inputs — US soybean exports {CUR}', f'Yellow = inputs. Starting point = ESR week {WK} (w/e {LAST:%d %b %Y}). Volumes Mt unless stated.')
inp = [('USDA export forecast (mln bu)', 1720, '#,##0', True), ('Desk number (mln bu)', 1800, '#,##0', True),
       ('China full-year purchases (Mt, ESR basis)', 25, MT, True), ('Bushels → tonnes (Mt per mln bu)', 0.0272155, '0.0000000', False),
       ('ESR ÷ USDA (Census) ratio', 0.97729741421324, '0.000', False),
       ('USDA forecast on ESR basis (Mt)', '=B4*B7*B8', MT, False), ('Desk number on ESR basis (Mt)', '=B5*B7*B8', MT, False),
       (f'Rest of world booked at week {WK} (Mt, excl. China & Unknown)', sum(now_r[r] for r in REGS), MT, False),
       (f'China booked at week {WK} (Mt, named only)', now_r['China'], MT, False), (f'Unknown at week {WK} (Mt)', unk_now, MT, False),
       (f'Total shipped at week {WK} (Mt, all destinations)', ship_now, MT, False),
       ('Rest of world still to book — USDA path (Mt)', '=B9-B6-B11', MT, False), ('Rest of world still to book — desk path (Mt)', '=B10-B6-B11', MT, False),
       ('Rest of world usual still to book, 5-yr (Mt)', sum(usual[r] for r in REGS), MT, False),
       ('Scaling vs usual — USDA path', '=B15/B17', '0%', False), ('Scaling vs usual — desk path', '=B16/B17', '0%', False),
       ('Still to ship — USDA path (Mt)', '=B9-B14', MT, False), ('Still to ship — desk path (Mt)', '=B10-B14', MT, False)]
for i, (lb, v, fmt, isin) in enumerate(inp):
    r = 4 + i
    put(wsI, r, 1, lb, f=font(10, True)); put(wsI, r, 2, v, fmt, f=BLUE if isin or not isinstance(v, str) else font(), fill=YEL if isin else None)
for n_, ref in [('kUSDA', '$B$18'), ('kDESK', '$B$19'), ('ShipUSDA', '$B$20'), ('ShipDESK', '$B$21'), ('ShipNow', '$B$14'), ('ChinaFY', '$B$6')]:
    name(wb, n_, 'Inputs', ref)
r = 24
put(wsI, r, 1, 'By region', f=font(11, True, NAVY))
header(wsI, r + 1, ['Region', f'Booked at wk {WK} (Mt)', 'Usual still to book from here (5-yr, Mt)', 'USDA path: still to book (Mt)', 'Desk path: still to book (Mt)',
                    'USDA path full year (Mt)', 'Desk path full year (Mt)'], h=44)
REGROW = {}
for i, g in enumerate(REGS + ['China']):
    rr = r + 2 + i; REGROW[g] = rr
    put(wsI, rr, 1, g, f=font(10, True)); put(wsI, rr, 2, now_r[g], MT, f=BLUE); put(wsI, rr, 3, usual[g], MT, f=BLUE)
    if g == 'China':
        put(wsI, rr, 4, f'=MAX(0,ChinaFY-B{rr})', MT)
        put(wsI, rr, 5, f'=D{rr}', MT)
    else:
        put(wsI, rr, 4, f'=C{rr}*kUSDA', MT); put(wsI, rr, 5, f'=C{rr}*kDESK', MT)
    put(wsI, rr, 6, f'=B{rr}+D{rr}', MT); put(wsI, rr, 7, f'=B{rr}+E{rr}', MT)
note_r = r + 3 + len(REGS) + 1
wsI.cell(note_r, 1, 'China: named China bookings needed to reach the full-year China number (Unknown sales re-declared to China show up here as they are declared). '
         'Rest of world: the requirement is shared by each region\'s usual remaining bookings (same scaling for all regions).').font = font(9, False, '404040', True)
wsI.column_dimensions['A'].width = 52
for c in 'BCDEFG': wsI.column_dimensions[c].width = 16

# ---------------- Shapes (data) ----------------
title(wsS, 'Seasonal shapes (5-yr average share of the remaining bookings / shipments done by each week) and 5-yr average weekly volumes',
      f'Computed from the ESR weekly data ({B5[0]}/{str(B5[0]+1)[-2:]}–{B5[-1]}/{str(B5[-1]+1)[-2:]}; China excludes 2025/26). Hardcoded (blue).')
cols = REGS + ['China']
header(wsS, 4, ['Week'] + [f'Shape {g}' for g in cols] + ['Shape shipments'] + [f'5-yr avg weekly {g} (Mt)' for g in cols] + ['5-yr avg weekly shipments (Mt)'], h=60)
for w in range(WK, NW + 1):
    rr = 5 + w - WK
    put(wsS, rr, 1, w, '0')
    for j, g in enumerate(cols): put(wsS, rr, 2 + j, float(shape[g][w]), '0.000', f=BLUE)
    put(wsS, rr, 2 + len(cols), float(sshape[w]), '0.000', f=BLUE)
    for j, g in enumerate(cols): put(wsS, rr, 3 + len(cols) + j, float(avg_wk[g].get(w, 0)) if w > WK else None, '0.000', f=BLUE)
    put(wsS, rr, 3 + 2 * len(cols), float(ship_avg.get(w, 0)) if w > WK else None, '0.000', f=BLUE)
SL = 5 + NW - WK
for c in range(1, 4 + 2 * len(cols)): wsS.column_dimensions[L(c)].width = 11
def shp(g_idx, row): return f"'Shapes (data)'!${L(2+g_idx)}${row}"

# ---------------- Weekly targets (two sheets) ----------------
def targets(ws, which):
    k = 'kUSDA' if which == 'USDA' else 'kDESK'; shipT = 'ShipUSDA' if which == 'USDA' else 'ShipDESK'
    lab = 'USDA ' + "1,720" if which == 'USDA' else 'desk 1,800'
    title(ws, f'Week-by-week levels needed for the {"USDA forecast" if which == "USDA" else "desk number"} (values follow the Inputs tab)',
          'Weekly = net new sales (or shipments) needed in that week, kt. Cumulative = total booked (or shipped) needed by that week, Mt. '
          'Compare each Thursday\'s export sales report with these lines; the 5-yr average columns show what a normal week looks like.')
    H = ['Week', 'Week ending'] + [f'{g} — weekly sales (kt)' for g in REGS] + ['REST OF WORLD — weekly sales (kt)', 'Rest of world 5-yr avg weekly (kt)',
         'China (named) — weekly sales (kt)', 'TOTAL SHIPMENTS — weekly (kt)', 'Shipments 5-yr avg weekly (kt)'] + \
        [f'{g} — cumulative (Mt)' for g in REGS] + ['Rest of world — cumulative (Mt)', 'China named — cumulative (Mt)', 'Shipments — cumulative (Mt)']
    header(ws, 4, H, h=66)
    nR = len(REGS)
    cRW, cAVG, cCN, cSH, cSHA = 3 + nR, 4 + nR, 5 + nR, 6 + nR, 7 + nR
    cC0 = 8 + nR; cCRW, cCCN, cCSH = cC0 + nR, cC0 + nR + 1, cC0 + nR + 2
    r0 = 5
    for w in range(WK, NW + 1):
        rr = r0 + w - WK; srow = 5 + w - WK
        put(ws, rr, 1, w, '0'); put(ws, rr, 2, (LAST + pd.Timedelta(weeks=w - WK)).to_pydatetime(), 'dd-mmm')
        for j, g in enumerate(REGS):
            ir = REGROW[g]
            put(ws, rr, cC0 + j, f"=Inputs!$B${ir}+Inputs!$C${ir}*{k}*{shp(j, srow)}", MT)
        put(ws, rr, cCRW, f'=SUM({L(cC0)}{rr}:{L(cC0+nR-1)}{rr})', MT, f=font(10, True))
        irc = REGROW['China']
        put(ws, rr, cCCN, f"=Inputs!$B${irc}+Inputs!$D${irc}*{shp(nR, srow)}", MT)
        put(ws, rr, cCSH, f"=ShipNow+{shipT}*{shp(nR+1, srow)}", MT, f=font(10, True))
        if w == WK:
            for c in list(range(3, cSHA + 1)): put(ws, rr, c, None, fill=CURF)
            ws.cell(rr, 3, 'starting point (actual)').font = font(9, False, '595959', True)
            continue
        for j in range(nR): put(ws, rr, 3 + j, f'=({L(cC0+j)}{rr}-{L(cC0+j)}{rr-1})*1000', KT)
        put(ws, rr, cRW, f'=({L(cCRW)}{rr}-{L(cCRW)}{rr-1})*1000', KT, f=font(10, True))
        put(ws, rr, cAVG, '=(' + '+'.join(f"'Shapes (data)'!{L(3+len(cols)+j)}{srow}" for j in range(nR)) + ')*1000', KT, f=font(9, False, '595959'))
        put(ws, rr, cCN, f'=({L(cCCN)}{rr}-{L(cCCN)}{rr-1})*1000', KT)
        put(ws, rr, cSH, f'=({L(cCSH)}{rr}-{L(cCSH)}{rr-1})*1000', KT, f=font(10, True))
        put(ws, rr, cSHA, f"='Shapes (data)'!{L(3+2*len(cols))}{srow}*1000", KT, f=font(9, False, '595959'))
    RL = r0 + NW - WK
    ws.column_dimensions['A'].width = 6; ws.column_dimensions['B'].width = 9
    for c in range(3, cCSH + 1): ws.column_dimensions[L(c)].width = 12
    ws.freeze_panes = 'C5'
    return dict(r0=r0, RL=RL, cCRW=cCRW, cCSH=cCSH, cCCN=cCCN, cRW=cRW, cSH=cSH, cC0=cC0)
TA = targets(wsA, 'USDA'); TB = targets(wsB, 'DESK')

# ---------------- Tracker ----------------
title(wsT, 'Tracker — paste each Thursday\'s cumulative numbers (Mt) into the yellow columns',
      'Pace score = how far the actual sits between the USDA path (0) and the desk path (1). Below 0 = slower than needed for USDA; above 1 = ahead of our 1,800.')
TH = ['Week', 'Week ending', 'Rest of world booked — ACTUAL (Mt)', 'Needed for USDA', 'Needed for desk', 'Pace score (RoW)', 'Status (RoW)',
      'China + 60% Unknown — ACTUAL (Mt)', 'China full-year assumption', 'Total shipped — ACTUAL (Mt)', 'Needed for USDA', 'Needed for desk', 'Pace score (shipments)']
header(wsT, 4, TH, h=56)
for w in range(WK, NW + 1):
    rr = 5 + w - WK; tr = TA['r0'] + w - WK
    put(wsT, rr, 1, w, '0'); put(wsT, rr, 2, f"='Weekly targets USDA'!B{tr}", 'dd-mmm')
    put(wsT, rr, 3, round(sum(now_r[r] for r in REGS), 3) if w == WK else None, MT, f=BLUE, fill=YEL)
    put(wsT, rr, 4, f"='Weekly targets USDA'!{L(TA['cCRW'])}{tr}", MT); put(wsT, rr, 5, f"='Weekly targets desk'!{L(TB['cCRW'])}{tr}", MT)
    put(wsT, rr, 6, f'=IF(C{rr}="","",IF(E{rr}=D{rr},"",(C{rr}-D{rr})/(E{rr}-D{rr})))', '0.00', f=font(10, True))
    put(wsT, rr, 7, f'=IF(F{rr}="","",IF(F{rr}<0,"Below USDA pace",IF(F{rr}<1,"Between USDA and desk","At/above desk")))')
    put(wsT, rr, 8, round(now_r['China'] + 0.6 * unk_now, 3) if w == WK else None, MT, f=BLUE, fill=YEL)
    put(wsT, rr, 9, '=ChinaFY', MT)
    put(wsT, rr, 10, round(ship_now, 3) if w == WK else None, MT, f=BLUE, fill=YEL)
    put(wsT, rr, 11, f"='Weekly targets USDA'!{L(TA['cCSH'])}{tr}", MT); put(wsT, rr, 12, f"='Weekly targets desk'!{L(TB['cCSH'])}{tr}", MT)
    put(wsT, rr, 13, f'=IF(J{rr}="","",IF(L{rr}=K{rr},"",(J{rr}-K{rr})/(L{rr}-K{rr})))', '0.00', f=font(10, True))
TL = 5 + NW - WK
for col in ('F', 'M'):
    rg = f'{col}6:{col}{TL}'
    wsT.conditional_formatting.add(rg, CellIsRule(operator='lessThan', formula=['0'], fill=RED))
    wsT.conditional_formatting.add(rg, CellIsRule(operator='between', formula=['0', '0.999'], fill=AMB))
    wsT.conditional_formatting.add(rg, CellIsRule(operator='greaterThanOrEqual', formula=['1'], fill=GRN))
wsT.column_dimensions['A'].width = 6; wsT.column_dimensions['B'].width = 9
for c in range(3, 14): wsT.column_dimensions[L(c)].width = 13
wsT.column_dimensions['G'].width = 20; wsT.freeze_panes = 'C5'
lc = LineChart(); lc.title = 'Rest-of-world bookings: actual vs paths (Mt)'; lc.y_axis.numFmt = '0'
lc.add_data(Reference(wsT, min_col=3, max_col=5, min_row=4, max_row=TL), titles_from_data=True)
lc.set_categories(Reference(wsT, min_col=1, min_row=5, max_row=TL)); lc.width = 18; lc.height = 9
lc.x_axis.delete = False; lc.y_axis.delete = False
for s_ in lc.series: s_.smooth = False
wsT.add_chart(lc, 'O4')
for ws in wb: ws.sheet_view.showGridLines = False
wb.save(out)
print('WK', WK, 'RoW now', round(sum(now_r[r] for r in REGS), 2), 'usual', round(sum(usual[r] for r in REGS), 2), {g: round(usual[g], 2) for g in REGS + ['China']})
