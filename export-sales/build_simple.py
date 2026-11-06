"""Build an easy-to-read summary workbook from the two analysis workbooks (recalculated).
Usage: python build_simple.py <main.xlsx> <europe.xlsx> <out.xlsx>
Numbers on the Data tab are taken from the analysis workbooks; every derived figure on the
presentation tabs (expected, gap, status, %, totals) is a live formula on that Data tab."""
import sys
import openpyxl
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.utils import get_column_letter as L
from openpyxl.formatting.rule import FormulaRule, CellIsRule
from openpyxl.workbook.defined_name import DefinedName

main_p, eur_p, out = sys.argv[1:4]
M = openpyxl.load_workbook(main_p, data_only=True)
E = openpyxl.load_workbook(eur_p, data_only=True)

def hdr_map(ws, row):
    return {str(c.value).strip(): c.column for c in ws[row] if c.value is not None}
# ---- pull from main workbook ----
dp = M['Demand & Pace']
dph = hdr_map(dp, 15)
rk = M['Ranking by Destination']; rkh = hdr_map(rk, 4)
CUR = [k for k in rkh if k[:2] == '20' and len(k) == 7][0]
LYc = [k for k in rkh if k[:2] == '20' and len(k) == 7][1]
WEEK = dp['B9'].value; USDA_BU = dp['B4'].value; CONV = dp['B5'].value; CHINA_T = dp['B7'].value
rkrow = {rk.cell(r, 2).value: r for r in range(5, rk.max_row + 1) if rk.cell(r, 2).value}
dates = {}
eh = E['Destinations']
eugrp = {eh.cell(r, 1).value: eh.cell(r, 2).value for r in range(5, 25) if eh.cell(r, 2).value in ('EU', 'non-EU')}
REGION_NICE = {'China': 'China', 'Unknown': 'Unknown'}
rows = []
for r in range(16, dp.max_row + 1):
    d = dp.cell(r, 2).value
    if not d or 'TOTAL' in str(d) or dp.cell(r, dph['Demand used']).value is None: continue
    rr = rkrow.get(d)
    g = lambda h: rk.cell(rr, rkh[h]).value if rr else 0
    rows.append(dict(
        dest=d.title().replace('China, Peoples Republic Of', 'China').replace('Korea, Republic Of', 'South Korea')
             .replace('South Africa, Republic Of', 'South Africa').replace('Georgia, Republic Of', 'Georgia'),
        raw=d, region=dp.cell(r, 3).value, eu=eugrp.get(d, ''),
        now=dp.cell(r, dph[f'Actual {CUR} at wk N']).value or 0,
        ly=g(LYc) or 0, avg5=g('5-yr avg') or 0, med5=g('5-yr median') if '5-yr median' in rkh else None,
        fy5=dp.cell(r, dph['Demand: own 5-yr avg full-MY exports']).value or 0,
        demand=dp.cell(r, dph['Demand used']).value or 0,
        pace=dp.cell(r, dph['Avg pace at wk N (% of year committed)']).value or 0,
        unshipped=g(f'Unshipped {CUR}') or 0))
unk_row = rkrow['UNKNOWN']
UNK = dict(now=rk.cell(unk_row, rkh[CUR]).value, ly=rk.cell(unk_row, rkh[LYc]).value, avg5=rk.cell(unk_row, rkh['5-yr avg']).value)
uc = M['USDA Check']
ROW_PACE5, ROW_PACE4 = uc['J11'].value, uc['J12'].value
CU_PACE5, CU_PACE4 = uc['G11'].value, uc['G12'].value
UNK_SHARE = uc['B3'].value if isinstance(uc['B3'].value, (int, float)) else 0.6
HIST = []
for r_ in range(6, 11):
    y_ = uc.cell(r_, 1).value
    tot_ = uc.cell(r_, 2).value
    cn_ = rk.cell(rkrow['CHINA, PEOPLES REPUBLIC OF'], rkh[y_]).value or 0
    un_ = rk.cell(rkrow['UNKNOWN'], rkh[y_]).value or 0
    HIST.append((y_, cn_, un_, tot_ - cn_ - un_, uc.cell(r_, 6).value, uc.cell(r_, 9).value))
# ---- pull from Europe workbook ----
win = E['Nov-Feb Window']
WIN = [(win.cell(r, 1).value, win.cell(r, 2).value, win.cell(r, 6).value, win.cell(r, 9).value) for r in range(9, 15)]
BENCH = {'normal': win['G20'].value, 'avg': win['G19'].value, 'needed': win['G22'].value}
WEAK = min(x[2] for x in WIN if x[0] != WIN[0][0]) / 17
we = E['Weekly Europe']
EW = [(we.cell(r, 1).value, we.cell(r, 3).value, we.cell(r, 4).value, we.cell(r, 5).value, we.cell(r, 9).value, we.cell(r, 10).value, we.cell(r, 11).value)
      for r in range(5, 57)]
DATE = E['Dashboard']['A1'].value.split('w/e ')[1].split(',')[0]

# ---------------- styles ----------------
F = 'Arial'
NAVY = '1F4E78'
HF = PatternFill('solid', fgColor=NAVY); TOT = PatternFill('solid', fgColor='EDEDED'); YEL = PatternFill('solid', fgColor='FFFF00')
KF = PatternFill('solid', fgColor='F2F2F2')
GRN = PatternFill('solid', fgColor='C6EFCE'); RED = PatternFill('solid', fgColor='FFC7CE'); GRY = PatternFill('solid', fgColor='F2F2F2')
thin = Side(style='thin', color='D9D9D9'); BOX = Border(bottom=thin)
KT = '#,##0;-#,##0;"–"'; SKT = '+#,##0;-#,##0;0'; PC = '+0%;-0%;0%'; P0 = '0%'
def font(sz=10, b=False, col='000000', i=False): return Font(name=F, size=sz, bold=b, color=col, italic=i)
def put(ws, r, c, v, fmt=None, f=None, fill=None, al=None, border=True):
    x = ws.cell(r, c, v); x.font = f or font()
    if fmt: x.number_format = fmt
    if fill: x.fill = fill
    if al: x.alignment = al
    if border: x.border = BOX
    return x
def title(ws, t, sub):
    ws['A1'] = t; ws['A1'].font = font(16, True, NAVY)
    ws['A2'] = sub; ws['A2'].font = font(10, False, '595959', True)
def header(ws, r, labels, c0=1, h=36):
    for i, t in enumerate(labels):
        x = ws.cell(r, c0 + i, t); x.font = font(10, True, 'FFFFFF'); x.fill = HF
        x.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    ws.row_dimensions[r].height = h
def note(ws, r, t, c=1, sz=9): x = ws.cell(r, c, t); x.font = font(sz, False, '404040', True); return x
def status_colors(ws, rng, first_cell):
    ws.conditional_formatting.add(rng, FormulaRule(formula=[f'LEFT({first_cell},5)="Ahead"'], fill=GRN, font=font(10, True, '006100')))
    ws.conditional_formatting.add(rng, FormulaRule(formula=[f'LEFT({first_cell},6)="Behind"'], fill=RED, font=font(10, True, '9C0006')))
    ws.conditional_formatting.add(rng, FormulaRule(formula=[f'LEFT({first_cell},2)="On"'], fill=GRY))
def kpi(ws, r, c, label, formula, fmt, sub=None):
    x = ws.cell(r, c, label); x.font = font(9, True, '595959'); x.fill = KF
    y = ws.cell(r + 1, c, formula); y.font = font(18, True, NAVY); y.number_format = fmt; y.fill = KF
    y.alignment = Alignment(horizontal='left')
    z = ws.cell(r + 2, c, sub); z.font = font(8, False, '595959', True); z.fill = KF
    if isinstance(sub, str) and sub.startswith('='): z.number_format = '@'

wb = Workbook()
wsR = wb.active; wsR.title = 'Read Me'
wsS = wb.create_sheet('Summary')
wsA = wb.create_sheet('All Destinations')
wsE = wb.create_sheet('Europe')
wsD = wb.create_sheet('Data')

# ================= Data =================
title(wsD, 'Data behind this file', f'Taken from the analysis workbooks (USDA Export Sales, data through w/e {DATE}, week {WEEK} of MY {CUR}). '
      'Blue = numbers from the analysis; black = formulas. Volumes in thousand tonnes (kt).')
put(wsD, 4, 1, 'Settings', f=font(11, True, NAVY), border=False)
settings = [('USDA US soybean export forecast (mln bu)', USDA_BU, '#,##0', 'User input'),
            ('Tonnes per bushel', CONV, '0.0000000', '60 lb bushel'),
            ('USDA forecast in kt', '=B5*B6*1000', KT, ''),
            ('China assumption (kt)', CHINA_T / 1000, KT, 'User assumption'),
            ('Rest of World demand (kt)', '=B7-B8', KT, 'USDA − China'),
            ('"On pace" band: one cargo (kt)', 65, KT, 'Gaps smaller than this are timing noise'),
            ('Week of the season', WEEK, '0', f'w/e {DATE}'),
            ('Rest of World: share usually booked by now (5-yr avg)', '=H26', P0, 'Incl. its share of Unknown — see history below'),
            ('Rest of World: share usually booked by now (excl. last year)', '=H27', P0, ''),
            ('Unknown destinations: booked now (kt)', UNK['now'] / 1000, KT, 'Sales not yet assigned to a destination'),
            ('Unknown: same week last year (kt)', UNK['ly'] / 1000, KT, ''),
            ('Unknown: same week 5-yr avg (kt)', UNK['avg5'] / 1000, KT, ''),
            ('Share of Unknown that ends up China', UNK_SHARE, P0, 'User indication (~60%, give or take); the rest is assigned to Rest of World')]
for i, (lab, v, fmt, nt) in enumerate(settings):
    r = 5 + i
    put(wsD, r, 1, lab, f=font(10, True))
    put(wsD, r, 2, v, fmt, f=font(10, False, '000000' if isinstance(v, str) else '0000FF'),
        fill=YEL if i in (0, 3, 5, 12) else None)
    note(wsD, r, nt, 3)
for nm, ref in [('USDA_kt', '$B$7'), ('China_kt', '$B$8'), ('RoW_kt', '$B$9'), ('Cargo', '$B$10'), ('Conv', '$B$6'),
                ('RoW_pace5', '$B$12'), ('RoW_pace4', '$B$13'), ('Unk_now', '$B$14'), ('Unk_ly', '$B$15'), ('Unk_avg', '$B$16'), ('Unk_CN', '$B$17'), ('CN_p5', '$G$26'), ('CN_p4', '$G$27')]:
    wb.defined_names[nm] = DefinedName(nm, attr_text=f"Data!{ref}")
HR0 = 19
put(wsD, HR0, 1, 'China / Unknown / Rest of World by season — bookings at this week and full-year exports (kt)', f=font(11, True, NAVY), border=False)
header(wsD, HR0 + 1, ['Season', 'China booked', 'Unknown booked', 'Rest of World booked (named)', 'China full year', 'Rest of World full year',
                      'China: share booked by now (incl. its share of Unknown)', 'Rest of World: share booked by now (incl. rest of Unknown)'], h=48)
for i, (y_, cn_, un_, rw_, cf_, rf_) in enumerate(HIST):
    r = HR0 + 2 + i
    put(wsD, r, 1, y_)
    for j, v in enumerate((cn_, un_, rw_, cf_, rf_), 2): put(wsD, r, j, v / 1000, KT, f=font(10, False, '0000FF'))
    put(wsD, r, 7, f'=(B{r}+Unk_CN*C{r})/E{r}', P0); put(wsD, r, 8, f'=(D{r}+(1-Unk_CN)*C{r})/F{r}', P0)
put(wsD, 26, 1, '5-yr average', f=font(10, True), fill=TOT); put(wsD, 27, 1, 'Excl. last year', f=font(10, True), fill=TOT)
for c in (7, 8):
    put(wsD, 26, c, f'=AVERAGE({L(c)}21:{L(c)}25)', P0, f=font(10, True), fill=TOT)
    put(wsD, 27, c, f'=AVERAGE({L(c)}22:{L(c)}25)', P0, f=font(10, True), fill=TOT)
for c in range(2, 7): put(wsD, 26, c, None, fill=TOT); put(wsD, 27, c, None, fill=TOT)
note(wsD, 28, 'Unknown is split between China and Rest of World in past seasons too, so "share booked by now" stays like-for-like with this season.')
DR = 31
put(wsD, DR - 1, 1, 'By destination', f=font(11, True, NAVY), border=False)
DH = ['Destination', 'Region', 'EU / non-EU', f'Booked so far {CUR}', 'Same week last year', 'Same week 5-yr avg',
      'Full-year 5-yr avg', 'Demand this year', 'Share usually booked by now', 'Not yet shipped']
header(wsD, DR, DH)
rows.sort(key=lambda x: -x['demand'])
for i, x in enumerate(rows):
    r = DR + 1 + i
    vals = [x['dest'], x['region'], x['eu'] or None, x['now'] / 1000, x['ly'] / 1000, x['avg5'] / 1000, x['fy5'] / 1000,
            x['demand'] / 1000, x['pace'], x['unshipped'] / 1000]
    for j, v in enumerate(vals, 1):
        put(wsD, r, j, v, P0 if j == 9 else (KT if j > 3 else None), f=font(10, False, '0000FF'))
DL = DR + len(rows)
for nm, col in [('d_dest', 'A'), ('d_reg', 'B'), ('d_eu', 'C'), ('d_now', 'D'), ('d_ly', 'E'), ('d_avg', 'F'), ('d_fy', 'G'),
                ('d_dem', 'H'), ('d_pace', 'I'), ('d_unsh', 'J')]:
    wb.defined_names[nm] = DefinedName(nm, attr_text=f"Data!${col}${DR+1}:${col}${DL}")
wsD.column_dimensions['A'].width = 52; wsD.column_dimensions['B'].width = 22
for c in 'CDEFGHIJ': wsD.column_dimensions[c].width = 14
# Europe window + weekly
ER = DL + 3
put(wsD, ER, 1, 'Europe — net sales Nov–Feb (weeks 10–26) by season, kt', f=font(11, True, NAVY), border=False)
header(wsD, ER + 1, ['Season', 'EU-27', 'All Europe', 'Share of Europe\'s year booked Nov–Feb'])
for i, (s, eu, eur, sh) in enumerate(WIN):
    r = ER + 2 + i
    put(wsD, r, 1, s); put(wsD, r, 2, eu / 1000, KT, f=font(10, False, '0000FF')); put(wsD, r, 3, eur / 1000, KT, f=font(10, False, '0000FF'))
    put(wsD, r, 4, sh, P0, f=font(10, False, '0000FF'))
WE = ER + 2 + len(WIN) - 1
EB = WE + 2
put(wsD, EB, 1, 'Europe — weekly net-sales benchmarks for Nov–Feb (kt per week)', f=font(11, True, NAVY), border=False)
for i, (lab, v) in enumerate([('Normal year (5-yr avg excl. last year)', BENCH['normal']), ('5-yr average', BENCH['avg']),
                              ('Needed to reach this year\'s demand', BENCH['needed']), ('Weak year (last year)', WEAK)]):
    put(wsD, EB + 1 + i, 1, lab, f=font(10, True)); put(wsD, EB + 1 + i, 2, v / 1000, KT, f=font(10, False, '0000FF'))
EWR = EB + 7
put(wsD, EWR, 1, 'Europe — cumulative bookings by week of the season (kt)', f=font(11, True, NAVY), border=False)
header(wsD, EWR + 1, ['Week', f'{CUR}', f'{LYc}', '5-yr avg', 'Expected path for this year\'s demand', f'Net sales {CUR}', 'Net sales 5-yr avg'])
for i, (w, c, l, a, ex, ns, nsa) in enumerate(EW):
    r = EWR + 2 + i
    put(wsD, r, 1, w, '0')
    for j, v in enumerate([c, l, a, ex, ns, nsa], 2):
        put(wsD, r, j, None if v is None else v / 1000, KT, f=font(10, False, '0000FF'))
EWL = EWR + 1 + len(EW)

# helper for SUMIFS on data
def S(col, *crit): return 'SUMIFS(' + col + ''.join(f',{a},{b}' for a, b in crit) + ')'

# ================= Summary =================
title(wsS, f'US soybean export sales — where we stand at week {WEEK} of {CUR}',
      f'Data through {DATE}. "Booked" = sold to that destination for this season (shipped + not yet shipped). Volumes in thousand tonnes (kt).')
kpi(wsS, 4, 1, 'USDA export forecast', '=Data!B5', '#,##0" mln bu"', None)
wsS['A6'] = '=TEXT(USDA_kt,"#,##0")&" kt"'; wsS['A6'].font = font(8, False, '595959', True); wsS['A6'].fill = KF
kpi(wsS, 4, 3, 'Booked so far (all buyers)', '=SUM(d_now)+Unk_now', KT, None)
wsS['C6'] = '=TEXT((SUM(d_now)+Unk_now)/USDA_kt,"0%")&" of USDA\'s total"'; wsS['C6'].font = font(8, False, '595959', True); wsS['C6'].fill = KF
kpi(wsS, 4, 5, 'China booked (+60% of Unknown)', f'=SUMIFS(d_now,d_reg,"China")+Unk_CN*Unk_now', KT, None)
wsS['E6'] = '=TEXT((SUMIFS(d_now,d_reg,"China")+Unk_CN*Unk_now)/China_kt,"0%")&" of the 25 Mt assumption"'; wsS['E6'].font = font(8, False, '595959', True); wsS['E6'].fill = KF
kpi(wsS, 4, 7, 'Rest of World vs pace', '=SUMIFS(d_now,d_reg,"<>China")+(1-Unk_CN)*Unk_now-RoW_kt*RoW_pace5', SKT, None)
wsS['G6'] = 'kt ahead (+) / behind (−) of the pace needed'; wsS['G6'].font = font(8, False, '595959', True); wsS['G6'].fill = KF
for c in range(1, 9): wsS.column_dimensions[L(c)].width = 16
wsS.column_dimensions['A'].width = 30
r0 = 9
put(wsS, r0, 1, '1) Bookings at this point of the season', f=font(12, True, NAVY), border=False)
header(wsS, r0 + 1, ['Buyer', 'Booked so far', 'Same week last year', 'Same week 5-yr avg', 'vs last year', 'vs 5-yr avg'])
grp = [('China', f'=SUMIFS(d_now,d_reg,"China")', f'=SUMIFS(d_ly,d_reg,"China")', f'=SUMIFS(d_avg,d_reg,"China")'),
       ('Unknown (≈60% China / 40% RoW)', '=Unk_now', '=Unk_ly', '=Unk_avg'),
       ('Rest of World', f'=SUMIFS(d_now,d_reg,"<>China")', f'=SUMIFS(d_ly,d_reg,"<>China")', f'=SUMIFS(d_avg,d_reg,"<>China")'),
       ('   of which Europe', f'=SUMIFS(d_now,d_reg,"Europe")', f'=SUMIFS(d_ly,d_reg,"Europe")', f'=SUMIFS(d_avg,d_reg,"Europe")')]
for i, (lab, a, b, c) in enumerate(grp):
    r = r0 + 2 + i
    put(wsS, r, 1, lab, f=font(10, True))
    put(wsS, r, 2, a, KT); put(wsS, r, 3, b, KT); put(wsS, r, 4, c, KT)
    put(wsS, r, 5, f'=IF(C{r}=0,"new",IF(B{r}/C{r}-1>5,"> +500%",B{r}/C{r}-1))', PC, al=Alignment(horizontal="right")); put(wsS, r, 6, f'=IF(D{r}=0,"new",B{r}/D{r}-1)', PC)
rt = r0 + 6
put(wsS, rt, 1, 'Total', f=font(10, True), fill=TOT)
for c in (2, 3, 4): put(wsS, rt, c, f'={L(c)}{r0+2}+{L(c)}{r0+3}+{L(c)}{r0+4}', KT, f=font(10, True), fill=TOT)
put(wsS, rt, 5, f'=B{rt}/C{rt}-1', PC, f=font(10, True), fill=TOT); put(wsS, rt, 6, f'=B{rt}/D{rt}-1', PC, f=font(10, True), fill=TOT)
note(wsS, rt + 1, 'Last year is a weak comparison: China bought almost no US beans early in 2025/26. The 5-yr average is the fairer benchmark.')
ch = BarChart(); ch.type = 'col'; ch.title = 'Booked by week ' + str(WEEK) + ' (kt)'; ch.y_axis.numFmt = '#,##0'
ch.add_data(Reference(wsS, min_col=2, max_col=4, min_row=r0 + 1, max_row=r0 + 4), titles_from_data=True)
ch.set_categories(Reference(wsS, min_col=1, min_row=r0 + 2, max_row=r0 + 4)); ch.width = 15; ch.height = 7.5
ch.x_axis.delete = False; ch.y_axis.delete = False
wsS.add_chart(ch, 'H9')

r1 = rt + 3
put(wsS, r1, 1, '2) What USDA\'s forecast implies for each buyer', f=font(12, True, NAVY), border=False)
header(wsS, r1 + 1, ['Buyer', 'Demand this year', 'Should have booked by now', 'Booked so far', 'Gap', 'Status'])
lines = [('China (+60% of Unknown)', '=China_kt', '=China_kt*CN_p5', f'=SUMIFS(d_now,d_reg,"China")+Unk_CN*Unk_now'),
         ('Rest of World (+40% of Unknown)', '=RoW_kt', '=RoW_kt*RoW_pace5', f'=SUMIFS(d_now,d_reg,"<>China")+(1-Unk_CN)*Unk_now'),
         ('   of which Europe', f'=SUMIFS(d_dem,d_reg,"Europe")', f'=SUMPRODUCT((d_reg="Europe")*d_dem*d_pace)', f'=SUMIFS(d_now,d_reg,"Europe")')]
for i, (lab, a, b, c) in enumerate(lines):
    r = r1 + 2 + i
    put(wsS, r, 1, lab, f=font(10, True)); put(wsS, r, 2, a, KT); put(wsS, r, 3, b, KT); put(wsS, r, 4, c, KT)
    put(wsS, r, 5, f'=D{r}-C{r}', SKT, f=font(10, True))
    put(wsS, r, 6, f'=IF(ABS(E{r})<Cargo,"On pace",IF(E{r}>0,"Ahead","Behind"))', f=font(10, True), al=Alignment(horizontal='center'))
status_colors(wsS, f'F{r1+2}:F{r1+4}', f'F{r1+2}')
wsS.cell(r1 + 5, 1, f'="China\'s history includes 2025/26 (almost no early buying). Without that year China should have booked "&TEXT(China_kt*CN_p4,"#,##0")&" kt by now — still ahead by "&TEXT(D{r1+2}-China_kt*CN_p4,"#,##0")&" kt."').font = font(9, False, '404040', True)
note(wsS, r1 + 6, 'Unknown sales are split 60% China / 40% Rest of World (Data tab, cell B17), in this season and in past seasons alike. Europe uses named sales only.')

r2 = r1 + 8
put(wsS, r2, 1, '3) What this means for USDA\'s 1,695 mln bu', f=font(12, True, NAVY), border=False)
header(wsS, r2 + 1, ['Scenario', 'Full-year exports (kt)', 'Full-year (mln bu)', 'vs USDA (mln bu)'])
sc = [('China at 25 Mt, Rest of World at its usual pace (5-yr)', f'=China_kt+(SUMIFS(d_now,d_reg,"<>China")+(1-Unk_CN)*Unk_now)/RoW_pace5'),
      ('China at 25 Mt, Rest of World at its usual pace (excl. last year)', f'=China_kt+(SUMIFS(d_now,d_reg,"<>China")+(1-Unk_CN)*Unk_now)/RoW_pace4')]
for i, (lab, f_) in enumerate(sc):
    r = r2 + 2 + i
    put(wsS, r, 1, lab, f=font(10, True)); put(wsS, r, 2, f_, KT)
    put(wsS, r, 3, f'=B{r}/Conv/1000', '#,##0'); put(wsS, r, 4, f'=C{r}-Data!$B$5', '+#,##0;-#,##0;0', f=font(10, True))
wsS.conditional_formatting.add(f'D{r2+2}:D{r2+3}', CellIsRule(operator='greaterThan', formula=['0'], fill=GRN))
note(wsS, r2 + 4, 'Mechanical estimate, not a forecast. In past normal seasons this method landed within about ±8% of the final number (≈ ±150 mln bu);')
note(wsS, r2 + 5, 'it failed when China\'s buying behaviour changed. Our view: a modest USDA increase in November (+25 to +75 mln bu) is more likely than a large one.')
wsS.column_dimensions['A'].width = 52

# ================= All Destinations =================
title(wsA, f'All destinations — are buyers ahead of or behind the pace needed for this year\'s demand? (week {WEEK})',
      'Sorted from most ahead to most behind. "On pace" = gap smaller than one cargo (65 kt). Volumes in kt. Destinations with demand under 100 kt '
      'and a gap under 20 kt are grouped in "All other destinations".')
AH = ['#', 'Destination', 'Region', 'Booked so far', 'Should have booked by now', 'Gap', 'Status', 'Demand this year',
      'Share usually booked by now', 'Same week last year', 'Same week 5-yr avg', 'vs 5-yr avg', 'Not yet shipped']
header(wsA, 4, AH, h=40)
def gap(x): return x['now'] - x['demand'] * x['pace']
show = [x for x in rows if x['raw'] != 'CHINA, PEOPLES REPUBLIC OF' and (x['demand'] >= 100000 or abs(gap(x)) >= 20000)]
show.sort(key=lambda x: -gap(x))
other = [x for x in rows if x['raw'] != 'CHINA, PEOPLES REPUBLIC OF' and x not in show]
def drow(ws, r, nm_cell):
    m = f'MATCH({nm_cell},d_dest,0)'
    return {k: f'INDEX({v},{m})' for k, v in [('reg', 'd_reg'), ('now', 'd_now'), ('ly', 'd_ly'), ('avg', 'd_avg'), ('dem', 'd_dem'),
                                                ('pace', 'd_pace'), ('unsh', 'd_unsh')]}
A1 = 5
for i, x in enumerate(show):
    r = A1 + i
    put(wsA, r, 1, i + 1, '0', f=font(9, False, '7F7F7F')); put(wsA, r, 2, x['dest'], f=font(10, True))
    d = drow(wsA, r, f'$B{r}')
    put(wsA, r, 3, '=' + d['reg']); put(wsA, r, 4, '=' + d['now'], KT, f=font(10, True)); put(wsA, r, 5, f"={d['dem']}*{d['pace']}", KT)
    put(wsA, r, 6, f'=D{r}-E{r}', SKT, f=font(10, True))
    put(wsA, r, 7, f'=IF(ABS(F{r})<Cargo,"On pace",IF(F{r}>0,"Ahead","Behind"))', al=Alignment(horizontal='center'))
    put(wsA, r, 8, '=' + d['dem'], KT); put(wsA, r, 9, '=' + d['pace'], P0)
    put(wsA, r, 10, '=' + d['ly'], KT); put(wsA, r, 11, '=' + d['avg'], KT)
    put(wsA, r, 12, f'=IF(K{r}=0,"new",D{r}/K{r}-1)', PC); put(wsA, r, 13, '=' + d['unsh'], KT)
A2 = A1 + len(show) - 1
ro = A2 + 1
put(wsA, ro, 2, f'All other destinations ({len(other)})', f=font(10, False, '595959', True))
put(wsA, ro, 3, '')
inlist = 'ISNA(MATCH(d_dest,$B$' + str(A1) + ':$B$' + str(A2) + ',0))*(d_reg<>"China")'
put(wsA, ro, 4, f'=SUMPRODUCT({inlist}*d_now)', KT); put(wsA, ro, 5, f'=SUMPRODUCT({inlist}*d_dem*d_pace)', KT)
put(wsA, ro, 6, f'=D{ro}-E{ro}', SKT); put(wsA, ro, 8, f'=SUMPRODUCT({inlist}*d_dem)', KT)
put(wsA, ro, 10, f'=SUMPRODUCT({inlist}*d_ly)', KT); put(wsA, ro, 11, f'=SUMPRODUCT({inlist}*d_avg)', KT)
put(wsA, ro, 12, f'=IF(K{ro}=0,"new",D{ro}/K{ro}-1)', PC); put(wsA, ro, 13, f'=SUMPRODUCT({inlist}*d_unsh)', KT)
rt = ro + 1
put(wsA, rt, 2, 'REST OF WORLD TOTAL', f=font(10, True), fill=TOT)
for c in (1, 3, 9): put(wsA, rt, c, None, fill=TOT)
for c in (4, 5, 6, 8, 10, 11, 13): put(wsA, rt, c, f'=SUM({L(c)}{A1}:{L(c)}{ro})', SKT if c == 6 else KT, f=font(10, True), fill=TOT)
put(wsA, rt, 7, f'=IF(ABS(F{rt})<Cargo,"On pace",IF(F{rt}>0,"Ahead","Behind"))', f=font(10, True), fill=TOT, al=Alignment(horizontal='center'))
put(wsA, rt, 12, f'=D{rt}/K{rt}-1', PC, f=font(10, True), fill=TOT)
note(wsA, rt + 2, 'Named destinations only. The Summary\'s Rest of World line also includes 40% of Unknown sales, so its gap vs pace is larger.', c=2)
put(wsA, rt + 1, 2, 'Check vs Summary (must be 0)', f=font(8, False, '7F7F7F', True))
put(wsA, rt + 1, 4, f"=D{rt}-Summary!B{r0+4}", KT, f=font(8, False, '7F7F7F'))
status_colors(wsA, f'G{A1}:G{rt}', f'G{A1}')
wsA.conditional_formatting.add(f'F{A1}:F{A2}', CellIsRule(operator='greaterThan', formula=['0'], font=font(10, True, '006100')))
wsA.conditional_formatting.add(f'F{A1}:F{A2}', CellIsRule(operator='lessThan', formula=['0'], font=font(10, True, '9C0006')))
for c, w in zip(range(1, 14), [4, 26, 22, 12, 13, 10, 11, 12, 12, 12, 12, 10, 11]): wsA.column_dimensions[L(c)].width = w
wsA.freeze_panes = 'C5'
gc = BarChart(); gc.type = 'bar'; gc.title = 'Gap vs pace needed (kt): ahead (+) / behind (−)'; gc.y_axis.numFmt = '#,##0'
gc.add_data(Reference(wsA, min_col=6, min_row=4, max_row=A2), titles_from_data=True)
gc.set_categories(Reference(wsA, min_col=2, min_row=A1, max_row=A2)); gc.legend = None
gc.x_axis.scaling.orientation = 'maxMin'; gc.x_axis.tickLblPos = 'low'; gc.width = 16; gc.height = 14
gc.x_axis.delete = False; gc.y_axis.delete = False
wsA.add_chart(gc, 'O4')

# ================= Europe =================
title(wsE, f'Europe — US soybean bookings at week {WEEK} of {CUR}',
      'Europe = EU-27 + UK, Norway, Switzerland. "On pace" = gap smaller than one cargo (65 kt). Volumes in kt.')
kpi(wsE, 4, 1, 'Europe booked so far', '=SUMIFS(d_now,d_reg,"Europe")', KT, None)
wsE['A6'] = '=TEXT(SUMIFS(d_now,d_reg,"Europe")/SUMIFS(d_avg,d_reg,"Europe")-1,"+0%;-0%")&" vs 5-yr avg"'; wsE['A6'].font = font(8, False, '595959', True); wsE['A6'].fill = KF
kpi(wsE, 4, 3, 'Europe demand this year', '=SUMIFS(d_dem,d_reg,"Europe")', KT, None)
wsE['C6'] = '=TEXT(SUMIFS(d_dem,d_reg,"Europe")/Conv/1000,"0")&" mln bu"'; wsE['C6'].font = font(8, False, '595959', True); wsE['C6'].fill = KF
kpi(wsE, 4, 5, 'Gap vs pace needed', '=SUMIFS(d_now,d_reg,"Europe")-SUMPRODUCT((d_reg="Europe")*d_dem*d_pace)', SKT, None)
wsE['E6'] = '=IF(ABS(E5)<Cargo,"On pace",IF(E5>0,"Ahead","Behind"))'; wsE['E6'].font = font(9, True, NAVY); wsE['E6'].fill = KF
kpi(wsE, 4, 7, 'Needed per week, Nov–Feb', f'=Data!B{EB+3}', KT, None)
wsE['G6'] = '="kt/week (normal year: "&TEXT(Data!B' + str(EB + 1) + ',"0")&")"'; wsE['G6'].font = font(8, False, '595959', True); wsE['G6'].fill = KF
for c in range(1, 9): wsE.column_dimensions[L(c)].width = 16
r0 = 9
put(wsE, r0, 1, '1) By destination', f=font(12, True, NAVY), border=False)
EHh = ['Destination', 'EU / non-EU', 'Booked so far', 'Should have booked by now', 'Gap', 'Status', 'Demand this year',
       'Share usually booked by now', 'Same week last year', 'Same week 5-yr avg', 'Not yet shipped']
header(wsE, r0 + 1, EHh, h=40)
eur = [x for x in rows if x['region'] == 'Europe' and (x['demand'] >= 50000 or x['now'] > 0 or x['avg5'] > 20000)]
eur.sort(key=lambda x: (x['eu'] != 'EU', -x['demand']))
E1 = r0 + 2
for i, x in enumerate(eur):
    r = E1 + i
    put(wsE, r, 1, x['dest'], f=font(10, True)); d = drow(wsE, r, f'$A{r}')
    put(wsE, r, 2, f'=INDEX(d_eu,MATCH($A{r},d_dest,0))', al=Alignment(horizontal='center'))
    put(wsE, r, 3, '=' + d['now'], KT, f=font(10, True)); put(wsE, r, 4, f"={d['dem']}*{d['pace']}", KT)
    put(wsE, r, 5, f'=C{r}-D{r}', SKT, f=font(10, True))
    put(wsE, r, 6, f'=IF(ABS(E{r})<Cargo,"On pace",IF(E{r}>0,"Ahead","Behind"))', al=Alignment(horizontal='center'))
    put(wsE, r, 7, '=' + d['dem'], KT); put(wsE, r, 8, '=' + d['pace'], P0)
    put(wsE, r, 9, '=' + d['ly'], KT); put(wsE, r, 10, '=' + d['avg'], KT); put(wsE, r, 11, '=' + d['unsh'], KT)
E2 = E1 + len(eur) - 1
eo = E2 + 1
put(wsE, eo, 1, 'Other European destinations', f=font(10, False, '595959', True)); put(wsE, eo, 2, '')
inl = f'ISNA(MATCH(d_dest,$A${E1}:$A${E2},0))*(d_reg="Europe")'
for c, f_ in ((3, f'SUMPRODUCT({inl}*d_now)'), (4, f'SUMPRODUCT({inl}*d_dem*d_pace)'), (7, f'SUMPRODUCT({inl}*d_dem)'),
              (9, f'SUMPRODUCT({inl}*d_ly)'), (10, f'SUMPRODUCT({inl}*d_avg)'), (11, f'SUMPRODUCT({inl}*d_unsh)')):
    put(wsE, eo, c, '=' + f_, KT)
put(wsE, eo, 5, f'=C{eo}-D{eo}', SKT)
et = eo + 1
put(wsE, et, 1, 'EUROPE TOTAL', f=font(10, True), fill=TOT)
for c in (2, 8): put(wsE, et, c, None, fill=TOT)
for c in (3, 4, 5, 7, 9, 10, 11): put(wsE, et, c, f'=SUM({L(c)}{E1}:{L(c)}{eo})', SKT if c == 5 else KT, f=font(10, True), fill=TOT)
put(wsE, et, 6, f'=IF(ABS(E{et})<Cargo,"On pace",IF(E{et}>0,"Ahead","Behind"))', f=font(10, True), fill=TOT, al=Alignment(horizontal='center'))
status_colors(wsE, f'F{E1}:F{et}', f'F{E1}')
for c, w in zip(range(1, 12), [24, 10, 12, 13, 10, 11, 12, 12, 12, 12, 11]): wsE.column_dimensions[L(c)].width = w
wsE.column_dimensions['A'].width = 30

r1 = et + 2
put(wsE, r1, 1, '2) The key months: November to February (weeks 10–26)', f=font(12, True, NAVY), border=False)
note(wsE, r1 + 1, 'Europe usually books about half of its US year in these weeks, before Brazil\'s harvest. Weekly net sales to watch (kt per week):')
bl = [('Normal year', f'=Data!B{EB+1}'), ('5-yr average (incl. weak last year)', f'=Data!B{EB+2}'),
      ('Needed for this year\'s demand', f'=Data!B{EB+3}'), ('Weak year (like last year)', f'=Data!B{EB+4}')]
header(wsE, r1 + 2, ['Benchmark', 'kt per week'], h=24)
for i, (lab, f_) in enumerate(bl):
    put(wsE, r1 + 3 + i, 1, lab, f=font(10, True)); put(wsE, r1 + 3 + i, 2, f_, KT, f=font(11, True, NAVY))
rr = r1 + 8
header(wsE, rr, ['Season', 'EU-27 net sales Nov–Feb', 'All Europe Nov–Feb', 'Share of year booked Nov–Feb'], h=36)
for i in range(len(WIN)):
    r = rr + 1 + i; dr = ER + 2 + i
    put(wsE, r, 1, f'=Data!A{dr}'); put(wsE, r, 2, f'=Data!B{dr}', KT); put(wsE, r, 3, f'=Data!C{dr}', KT); put(wsE, r, 4, f'=Data!D{dr}', P0)
rw = rr + len(WIN)
bc = BarChart(); bc.type = 'col'; bc.title = 'Europe: net sales Nov–Feb by season (kt)'; bc.y_axis.numFmt = '#,##0'
bc.add_data(Reference(wsE, min_col=3, min_row=rr, max_row=rw), titles_from_data=True)
bc.set_categories(Reference(wsE, min_col=1, min_row=rr + 1, max_row=rw)); bc.legend = None; bc.width = 15; bc.height = 7.5
bc.x_axis.delete = False; bc.y_axis.delete = False
wsE.add_chart(bc, f'M{r1}')
lc = LineChart(); lc.title = 'Europe: cumulative bookings by week (kt)'; lc.y_axis.numFmt = '#,##0'
lc.add_data(Reference(wsD, min_col=2, max_col=5, min_row=EWR + 1, max_row=EWL), titles_from_data=True)
lc.set_categories(Reference(wsD, min_col=1, min_row=EWR + 2, max_row=EWL)); lc.width = 15; lc.height = 8
lc.x_axis.title = 'Week of the season'; lc.x_axis.delete = False; lc.y_axis.delete = False
for s_ in lc.series: s_.smooth = False
wsE.add_chart(lc, 'M9')
r3 = rw + 2
put(wsE, r3, 1, '3) What to watch', f=font(12, True, NAVY), border=False)
for i, t in enumerate([
    '• Germany and the Netherlands (the big crushers) have nothing booked that is still waiting to ship. New forward buying there = first sign demand is back.',
    '• Spain, Portugal and Italy hold most of Europe\'s unshipped bookings.',
    '• Weekly net sales to Europe vs the benchmarks above, every Thursday.',
    '• EU deforestation rules (EUDR): may favour US beans over Brazilian from January 2027 (timing to be confirmed).',
    '• Cheaper Brazilian beans into Europe if China keeps buying in the US rather than in Brazil.']):
    note(wsE, r3 + 1 + i, t, sz=10)

# ================= Read Me =================
title(wsR, 'US soybean export sales by destination — how to read this file',
      f'Source: USDA Foreign Agricultural Service, weekly Export Sales Report (soybeans). Data through {DATE} (week {WEEK} of the {CUR} season, Sep–Aug).')
steps = [
    ('What we measure', '"Booked" = soybeans sold to a destination for this season, whether already shipped or not yet shipped.'),
    ('Step 1 — Same point in the season', f'We compare bookings at week {WEEK} of this season with week {WEEK} of the previous five seasons.'),
    ('Unknown destinations', 'Some sales are booked without a declared destination ("Unknown"). Historically about 60% ends up in China, so we assign 60% to China and 40% to the Rest of World — in this season and in past seasons, so comparisons stay like-for-like.'),
    ('Step 2 — Benchmarks', 'Last year, and the 5-year average. Last year is unusual: China bought almost no US beans early in 2025/26.'),
    ('Step 3 — This year\'s total', f'USDA forecast {USDA_BU:,.0f} mln bu = {USDA_BU*CONV*1000:,.0f} kt. We assume China takes {CHINA_T/1000:,.0f} kt; '
                                    'the rest goes to all other buyers ("Rest of World").'),
    ('Step 4 — Demand per destination', 'Rest of World is split with "floor + swing": every destination keeps what it bought in its second-weakest '
                                         'of the last five seasons (floor), plus the same share of its usual extra buying (swing). Steady buyers (Mexico, '
                                         'Japan) are cut little; volatile buyers (Egypt, Pakistan) absorb most of a smaller total.'),
    ('Step 5 — Seasonal pace', f'For each destination: what share of its full-year volume it usually has booked by week {WEEK} (5-season average).'),
    ('Step 6 — Ahead or behind', '"Should have booked by now" = demand × usual share. Gap = booked − should have. A gap smaller than one cargo '
                                  '(65 kt) counts as "On pace".'),
    ('Step 7 — Back to USDA', 'If China stays at 25 Mt and the Rest of World keeps its usual pace, what would full-year exports be in mln bu?'),
    ('Tabs', 'Summary — the big picture.  All Destinations — every material buyer, ahead/behind.  Europe — Europe in detail, with the '
             'November–February benchmarks.  Data — the numbers behind it (yellow cells can be changed).'),
    ('Colours', 'Green = ahead of pace, red = behind, grey = on pace. Blue numbers on the Data tab come from the analysis; black numbers are formulas.'),
    ('Caveats', 'Export-sales data runs slightly below USDA\'s official export figures. Five seasons is a short history. Buyers taking only 1–3 '
                'cargoes a year can look ahead or behind simply because of when one ship was booked.'),
]
for i, (k, v) in enumerate(steps):
    r = 4 + i
    a = wsR.cell(r, 1, k); a.font = font(11, True, NAVY); a.alignment = Alignment(vertical='top')
    b = wsR.cell(r, 2, v); b.font = font(11); b.alignment = Alignment(wrap_text=True, vertical='top')
    wsR.row_dimensions[r].height = 48 if len(v) > 150 else 32
wsR.column_dimensions['A'].width = 32; wsR.column_dimensions['B'].width = 120

for ws_ in (wsS, wsE):
    for rr_ in (4, 5, 6):
        for cc_ in range(1, 9):
            ws_.cell(rr_, cc_).fill = KF
    ws_.row_dimensions[5].height = 26
for ws in wb:
    ws.sheet_view.showGridLines = False
wb.save(out)
print('rows', len(rows), 'shown', len(show), 'europe', len(eur))
