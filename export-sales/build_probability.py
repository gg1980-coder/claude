"""US soybean exports: probability of ending above/below the USDA export forecast, based on sales pace.
Usage: python build_probability.py <ESR csv> <out.xlsx>
Source: USDA FAS Export Sales Reporting (ESR), SOYBEANS. Cross-check figures from the ESR pace report PDF (w/e 29-Oct-26).
Main model: Final (ESR basis) = China full-year (assumption) + RoW commitments at week N + RoW sales still to come,
where 'still to come' is drawn from the base seasons (mean, sd, Student-t with n-1 df). Converted to USDA (Census) basis
with an ESR/USDA ratio. All model maths are live Excel formulas on the Weekly Data tab."""
import sys
import numpy as np
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.chart import LineChart, BarChart, Reference
from openpyxl.utils import get_column_letter as L
from openpyxl.formatting.rule import CellIsRule, FormulaRule
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.datavalidation import DataValidation

src, out = sys.argv[1], sys.argv[2]
df = pd.read_csv(src, thousands=',')
df.columns = [c.strip() for c in df.columns]
df['Country'] = df['CountryName'].str.strip()
df['date'] = pd.to_datetime(df['WeekendingDate'])
df['Y'] = df['ReportingMarketingYearName'].str.extract(r'Sep (\d{4})').astype(int)
lo = pd.to_datetime(df.Y.astype(str) + '-08-31'); hi = pd.to_datetime((df.Y + 1).astype(str) + '-09-07')
ins = df[(df.date >= lo) & (df.date <= hi)].copy()
cs = int(ins.loc[ins.date.idxmax(), 'Y'])
WK = int(ins[ins.Y == cs].WeekNumber.max()); DATE = ins[ins.Y == cs].date.max()
def lab(y): return f'{y}/{str(y+1)[-2:]}'
YEARS = [y for y in sorted(ins.Y.unique()) if cs - 6 <= y <= cs]
PRIOR = [y for y in YEARS if y < cs]
CUR = lab(cs); NW = 52
ins = ins[ins.Y.isin(YEARS)]
ins['G'] = np.where(ins.Country.str.startswith('CHINA'), 'China', np.where(ins.Country == 'UNKNOWN', 'Unknown', 'Rest of World'))
cum = ins.pivot_table(index=['Y', 'G'], columns='WeekNumber', values='TotalCommitment', aggfunc='sum').reindex(columns=range(1, NW + 1)).fillna(0)
fin = ins[ins.Y < cs].sort_values('date').groupby(['Y', 'Country']).tail(1).groupby(['Y', 'G']).AccumulatedExports.sum()
GROUPS = ['China', 'Unknown', 'Rest of World']

# USDA (balance-sheet / Census basis) final exports, mln bu — APPROXIMATE, from memory; user should verify/replace.
USDA_FINAL = {2020: 2266, 2021: 2152, 2022: 1979, 2023: 1703, 2024: 1892, 2025: 1515}
USDA_SRC = {y: 'Official (Census) — monthly export file; matches WASDE' for y in USDA_FINAL}

# ---------------- styles ----------------
F = 'Arial'; NAVY = '1F4E78'
HF = PatternFill('solid', fgColor=NAVY); TOT = PatternFill('solid', fgColor='EDEDED'); YEL = PatternFill('solid', fgColor='FFFF00')
CURF = PatternFill('solid', fgColor='FFF2CC'); GRN = PatternFill('solid', fgColor='C6EFCE'); RED = PatternFill('solid', fgColor='FFC7CE')
thin = Side(style='thin', color='D9D9D9'); BOX = Border(bottom=thin)
MT = '#,##0.00'; KT = '#,##0'; BU = '#,##0'; SBU = '+#,##0;-#,##0;0'; P0 = '0%'
def font(sz=10, b=False, col='000000', i=False): return Font(name=F, size=sz, bold=b, color=col, italic=i)
BLUE = font(10, False, '0000FF'); GREENF = font(10, False, '008000')
def put(ws, r, c, v, fmt=None, f=None, fill=None, al=None, border=True):
    x = ws.cell(r, c, v); x.font = f or font()
    if fmt: x.number_format = fmt
    if fill: x.fill = fill
    if al: x.alignment = al
    if border: x.border = BOX
    return x
def title(ws, t, sub):
    ws['A1'] = t; ws['A1'].font = font(15, True, NAVY)
    ws['A2'] = sub; ws['A2'].font = font(9, False, '595959', True)
def header(ws, r, labels, c0=1, h=36):
    for i, t in enumerate(labels):
        x = ws.cell(r, c0 + i, t); x.font = font(9, True, 'FFFFFF'); x.fill = HF
        x.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    ws.row_dimensions[r].height = h
def note(ws, r, t, c=1, sz=9): x = ws.cell(r, c, t); x.font = font(sz, False, '404040', True); return x
def section(ws, r, t): x = ws.cell(r, 1, t); x.font = font(12, True, NAVY)
def name(wb, n, sh, ref): wb.defined_names[n] = DefinedName(n, attr_text=f"'{sh}'!{ref}")
def prob_ge(need, mean, se, df_):   # P(X >= need), X = mean + se * t(df)
    t = f'(({need})-({mean}))/({se})'
    return f'IF({t}>=0,TDIST({t},{df_},1),1-TDIST(-{t},{df_},1))'

wb = Workbook()
wsM = wb.active; wsM.title = 'Model'
wsH = wb.create_sheet('History')
wsB = wb.create_sheet('Backtest')
wsR = wb.create_sheet('USDA vs ESR basis')
wsW = wb.create_sheet('Weekly Data')
wsN = wb.create_sheet('Method')

# ================= Weekly Data =================
title(wsW, 'Cumulative soybean export commitments by group and in-season week (Mt), with full-year exports',
      f'Hardcoded from the ESR file (blue). Commitments = shipped + unshipped. Full-year = accumulated exports on the last in-season report. {CUR} through week {WK}.')
header(wsW, 4, ['Marketing year', 'Group', 'Full-year exports (Mt)'] + [f'W{w}' for w in range(1, NW + 1)])
r = 5
for y in YEARS:
    for g in GROUPS:
        put(wsW, r, 1, lab(y), f=BLUE, border=False); put(wsW, r, 2, g, f=BLUE, border=False)
        put(wsW, r, 3, None if y == cs else float(fin.get((y, g), 0)) / 1e6, MT, f=BLUE, border=False)
        for w in range(1, NW + 1):
            if y == cs and w > WK: continue
            x = wsW.cell(r, 3 + w, float(cum.loc[(y, g), w]) / 1e6 if (y, g) in cum.index else 0.0)
            x.number_format = MT; x.font = BLUE
        r += 1
WL = r - 1
for n_, c_ in [('WD_MY', 'A'), ('WD_G', 'B'), ('WD_F', 'C')]: name(wb, n_, 'Weekly Data', f'${c_}$5:${c_}${WL}')
name(wb, 'WD_W', 'Weekly Data', f'$D$5:${L(3+NW)}${WL}')
wsW.column_dimensions['A'].width = 12; wsW.column_dimensions['B'].width = 14; wsW.column_dimensions['C'].width = 12
for w in range(1, NW + 1): wsW.column_dimensions[L(3 + w)].width = 8
wsW.freeze_panes = 'D5'
def wk(g, my, w='Wk_N'): return f'SUMIFS(INDEX(WD_W,0,{w}),WD_G,"{g}",WD_MY,{my})'
def fy(g, my): return f'SUMIFS(WD_F,WD_G,"{g}",WD_MY,{my})'

# ================= USDA vs ESR basis =================
title(wsR, 'USDA (balance-sheet) vs export-sales (ESR) full-year exports — conversion ratio',
      'USDA (Census-basis) finals from the monthly export file you provided (they match the WASDE). ESR finals come from the weekly data.')
header(wsR, 4, ['Marketing year', 'USDA final exports (mln bu)', 'USDA final (Mt)', 'ESR full-year (Mt)', 'ESR ÷ USDA', 'Use? (Y/N)', 'Source of USDA number'])
for i, y in enumerate(PRIOR):
    r = 5 + i
    put(wsR, r, 1, lab(y))
    put(wsR, r, 2, USDA_FINAL.get(y), BU, f=BLUE, fill=YEL)
    put(wsR, r, 3, f'=IF(B{r}="","",B{r}*Conv)', MT)
    put(wsR, r, 4, '=' + '+'.join(fy(g, f'$A{r}') for g in GROUPS), MT)
    put(wsR, r, 5, f'=IF(C{r}="","",D{r}/C{r})', '0.000')
    put(wsR, r, 6, 'Y' if y in USDA_FINAL else 'N', f=BLUE, fill=YEL, al=Alignment(horizontal='center'))
    put(wsR, r, 7, USDA_SRC.get(y, 'From memory — please verify'), f=font(9, False, '404040', True))
RL = 5 + len(PRIOR) - 1
put(wsR, RL + 1, 1, 'Average ratio (rows marked Y)', f=font(10, True), fill=TOT)
put(wsR, RL + 1, 5, f'=AVERAGEIFS(E5:E{RL},F5:F{RL},"Y")', '0.000', f=font(10, True), fill=TOT)
for c in (2, 3, 4, 6): put(wsR, RL + 1, c, None, fill=TOT)
put(wsR, RL + 3, 1, 'Ratio used in the model', f=font(10, True)); put(wsR, RL + 3, 2, f'=E{RL+1}', '0.000', f=font(11, True, NAVY))
note(wsR, RL + 4, 'ESR usually runs ~2–3% below USDA\'s official (Census-based) exports. A ±0.01 change in the ratio moves the result by ~±0.5 Mt (≈ ±18 mln bu).')
name(wb, 'Ratio', wsR.title, f'$B${RL+3}')
wsR.column_dimensions['A'].width = 30
for c in 'BCDEF': wsR.column_dimensions[c].width = 16
wsR.column_dimensions['G'].width = 28

# ================= Model: inputs =================
title(wsM, f'US soybean exports {CUR}: probability of ending above USDA\'s forecast, based on sales pace',
      f'Data through w/e {DATE:%d %b %Y} (week {WK}). Yellow = inputs. ESR = USDA weekly export sales; USDA forecast is on balance-sheet (Census) basis — converted with the ratio on "USDA vs ESR basis".')
inp = [('USDA export forecast (mln bu)', 1720, BU, True, 'Nov-2026 WASDE (Oct: 1,695)'),
       ('Bushels → tonnes (MT per bu)', 0.0272155, '0.0000000', False, '60 lb bushel'),
       ('USDA forecast (Mt, balance-sheet basis)', '=B4*B5', MT, False, ''),
       ('ESR ÷ USDA ratio', '=Ratio', '0.000', False, 'From "USDA vs ESR basis" tab'),
       ('USDA forecast expressed on ESR basis (Mt)', '=B6*B7', MT, False, 'This is the bar the ESR data must clear'),
       ('China full-year assumption (Mt, ESR basis)', 25, MT, True, 'Purchase commitment; change to test scenarios'),
       ('Share of "Unknown" that ends up China', 0.6, P0, True, 'User indication'),
       ('Week N (in-season week)', WK, '0', True, f'Latest = {WK}. Lower it to re-run the model as of an earlier week')]
for i, (lb, v, fmt, isin, nt) in enumerate(inp):
    r = 4 + i
    put(wsM, r, 1, lb, f=font(10, True))
    put(wsM, r, 2, v, fmt, f=BLUE if isin else (GREENF if isinstance(v, str) and v.startswith('=Ratio') else font()), fill=YEL if isin else None)
    note(wsM, r, nt, 3)
wb.defined_names['Conv'] = DefinedName('Conv', attr_text="'Model'!$B$5")
for n_, ref in [('USDA_Mt', '$B$6'), ('Need_ESR', '$B$8'), ('China_A', '$B$9'), ('Unk_CN', '$B$10'), ('Wk_N', '$B$11')]: name(wb, n_, 'Model', ref)

# ================= History =================
title(wsH, 'Base seasons at week N: commitments, full-year exports, sales still to come and multipliers (Mt)',
      'Unknown split by the share on the Model tab (China / Rest of World). "Still to come" = full-year − committed at week N. Set Include = N to drop a season.')
HH = ['Marketing year', 'Include? (Y/N)', 'China at wk N', 'Unknown at wk N', 'RoW at wk N', 'China + share of Unknown', 'RoW + rest of Unknown',
      'Total at wk N', 'China full-year', 'RoW full-year', 'Total full-year', 'RoW still to come', 'Total still to come',
      'RoW multiplier', 'Total multiplier', 'China multiplier', 'inc']
header(wsH, 4, HH, h=44)
for i, y in enumerate(PRIOR):
    r = 5 + i; my = f'$A{r}'
    put(wsH, r, 1, lab(y), f=font(10, True))
    put(wsH, r, 2, 'Y', f=BLUE, fill=YEL, al=Alignment(horizontal='center'))
    put(wsH, r, 3, '=' + wk('China', my), MT); put(wsH, r, 4, '=' + wk('Unknown', my), MT); put(wsH, r, 5, '=' + wk('Rest of World', my), MT)
    put(wsH, r, 6, f'=C{r}+Unk_CN*D{r}', MT); put(wsH, r, 7, f'=E{r}+(1-Unk_CN)*D{r}', MT); put(wsH, r, 8, f'=C{r}+D{r}+E{r}', MT)
    put(wsH, r, 9, '=' + fy('China', my), MT); put(wsH, r, 10, '=' + fy('Rest of World', my) + '+' + fy('Unknown', my), MT)
    put(wsH, r, 11, f'=I{r}+J{r}', MT)
    put(wsH, r, 12, f'=J{r}-G{r}', MT, f=font(10, True)); put(wsH, r, 13, f'=K{r}-H{r}', MT, f=font(10, True))
    put(wsH, r, 14, f'=IF(G{r}=0,"",J{r}/G{r})', '0.00'); put(wsH, r, 15, f'=IF(H{r}=0,"",K{r}/H{r})', '0.00')
    put(wsH, r, 16, f'=IF(F{r}=0,"",I{r}/F{r})', '0.00')
    put(wsH, r, 17, f'=IF(B{r}="Y",1,0)', '0', f=font(8, False, '7F7F7F'))
H1, H2 = 5, 5 + len(PRIOR) - 1
INC = f'$Q${H1}:$Q${H2}'
stats = [('n (included)', 'n'), ('Mean', 'mean'), ('Std dev', 'sd'), ('Lowest', 'min'), ('Highest', 'max')]
for i, (lb, k) in enumerate(stats):
    r = H2 + 2 + i
    put(wsH, r, 1, lb, f=font(10, True), fill=TOT); put(wsH, r, 2, None, fill=TOT)
    for c in range(3, 17):
        rg = f'{L(c)}{H1}:{L(c)}{H2}'
        if k == 'n': f_ = f'=SUM({INC})'
        elif k == 'mean': f_ = f'=SUMPRODUCT({INC},{rg})/SUM({INC})'
        elif k == 'sd': f_ = f'=SQRT(SUMPRODUCT({INC},({rg}-{L(c)}{H2+3})^2)/(SUM({INC})-1))'
        elif k == 'min': f_ = f'=MIN(IF({INC}=1,{rg}))'
        else: f_ = f'=MAX(IF({INC}=1,{rg}))'
        put(wsH, r, c, f_, '0' if k == 'n' else ('0.00' if c >= 14 else MT), f=font(10, True), fill=TOT)
SN, SMEAN, SSD, SMIN, SMAX = (H2 + 2 + i for i in range(5))
# MIN(IF()) / MAX(IF()) need array entry: write as array formulas
from openpyxl.worksheet.formula import ArrayFormula
for c in range(3, 17):
    for rr in (SMIN, SMAX):
        cell = wsH.cell(rr, c); f_ = cell.value
        cell.value = ArrayFormula(f'{L(c)}{rr}', f_)
note(wsH, SMAX + 2, 'Mean and std dev use only the seasons marked Y. 2020/21 (record China buying) and 2025/26 (China absent) are unusual; '
                    'try excluding them to see how sensitive the result is.')
wsH.column_dimensions['A'].width = 16
for c in range(2, 18): wsH.column_dimensions[L(c)].width = 11
wsH.column_dimensions['Q'].width = 4
for n_, ref in [('n_yrs', f'$C${SN}'), ('RoWrem_mean', f'$L${SMEAN}'), ('RoWrem_sd', f'$L${SSD}'), ('Totrem_mean', f'$M${SMEAN}'),
                ('Totrem_sd', f'$M${SSD}'), ('RoWmul_mean', f'$N${SMEAN}'), ('RoWmul_sd', f'$N${SSD}'),
                ('Totmul_mean', f'$O${SMEAN}'), ('Totmul_sd', f'$O${SSD}')]:
    name(wb, n_, 'History', ref)

# ================= Model: current position + results =================
r0 = 13
section(wsM, r0, f'1) Where we are at week N ({CUR}, Mt, ESR basis)')
header(wsM, r0 + 1, ['Group', 'Committed (sold) so far', 'of which Unknown assigned', 'Full-year assumption / outlook'], h=30)
cn_now = f"{wk('China', 'CUR_MY')}+Unk_CN*{wk('Unknown', 'CUR_MY')}"
rw_now = f"{wk('Rest of World', 'CUR_MY')}+(1-Unk_CN)*{wk('Unknown', 'CUR_MY')}"
name(wb, 'CUR_MY', 'Model', f'$B${r0+6}')
put(wsM, r0 + 2, 1, 'China (+ its share of Unknown)', f=font(10, True)); put(wsM, r0 + 2, 2, '=' + cn_now, MT)
put(wsM, r0 + 2, 3, f"=Unk_CN*{wk('Unknown', 'CUR_MY')}", MT); put(wsM, r0 + 2, 4, '=China_A', MT)
put(wsM, r0 + 3, 1, 'Rest of World (+ rest of Unknown)', f=font(10, True)); put(wsM, r0 + 3, 2, '=' + rw_now, MT)
put(wsM, r0 + 3, 3, f"=(1-Unk_CN)*{wk('Unknown', 'CUR_MY')}", MT); put(wsM, r0 + 3, 4, f'=B{r0+3}+RoWrem_mean', MT)
put(wsM, r0 + 4, 1, 'Total', f=font(10, True), fill=TOT); put(wsM, r0 + 4, 2, f'=B{r0+2}+B{r0+3}', MT, f=font(10, True), fill=TOT)
put(wsM, r0 + 4, 3, f'=C{r0+2}+C{r0+3}', MT, fill=TOT); put(wsM, r0 + 4, 4, f'=D{r0+2}+D{r0+3}', MT, f=font(10, True), fill=TOT)
put(wsM, r0 + 6, 1, 'Current marketing year label', f=font(8, False, '7F7F7F')); put(wsM, r0 + 6, 2, CUR, f=font(8, False, '7F7F7F'))
note(wsM, r0 + 5, 'China already has (committed) vs assumption: ', sz=9)
wsM.cell(r0 + 5, 3, f'=TEXT(B{r0+2}/China_A,"0%")&" of the assumption booked"').font = font(9, True, NAVY)
CNN, RWN, TTN = f'$B${r0+2}', f'$B${r0+3}', f'$B${r0+4}'

r1 = r0 + 8
section(wsM, r1, '2) Probability of ending ABOVE the USDA forecast — four models')
RH = ['Model', 'Expected full year (Mt, ESR)', 'Uncertainty ±1 sd (Mt)', 'Expected full year (mln bu, USDA basis)', 'vs USDA (mln bu)',
      '80% range low (mln bu)', '80% range high (mln bu)', 'Probability ABOVE USDA', 'Probability BELOW USDA', 'Backtest error, normal seasons (Mt)']
header(wsM, r1 + 1, RH, h=48)
SE = 'SQRT(1+1/n_yrs)'
T80 = 'TINV(0.2,n_yrs-1)'
models = [
    ('A. RECOMMENDED — China = assumption; RoW = sold so far + its usual sales still to come', f'China_A+{RWN}+RoWrem_mean', f'RoWrem_sd*{SE}', 3.0),
    ('B. China = assumption; RoW = sold so far × its usual multiplier', f'China_A+{RWN}*RoWmul_mean', f'{RWN}*RoWmul_sd*{SE}', 2.9),
    ('C. All destinations: sold so far + usual total still to come', f'{TTN}+Totrem_mean', f'Totrem_sd*{SE}', 2.9),
    ('D. All destinations: sold so far × usual multiplier (classic "implied FY")', f'{TTN}*Totmul_mean', f'{TTN}*Totmul_sd*{SE}', 3.2),
]
for i, (lb, mean, se, bt) in enumerate(models):
    r = r1 + 2 + i
    put(wsM, r, 1, lb, f=font(10, i == 0), al=Alignment(wrap_text=True, vertical='center'))
    put(wsM, r, 2, '=' + mean, MT, f=font(10, True)); put(wsM, r, 3, '=' + se, MT)
    put(wsM, r, 4, f'=B{r}/Ratio/Conv', BU, f=font(10, True)); put(wsM, r, 5, f'=D{r}-$B$4', SBU, f=font(10, True))
    put(wsM, r, 6, f'=(B{r}-{T80}*C{r})/Ratio/Conv', BU); put(wsM, r, 7, f'=(B{r}+{T80}*C{r})/Ratio/Conv', BU)
    put(wsM, r, 8, '=' + prob_ge('Need_ESR', f'B{r}', f'C{r}', 'n_yrs-1'), P0, f=font(12, True, NAVY), al=Alignment(horizontal='center'))
    put(wsM, r, 9, f'=1-H{r}', P0, al=Alignment(horizontal='center'))
    put(wsM, r, 10, bt, '0.0', f=BLUE)
    wsM.row_dimensions[r].height = 30
RA = r1 + 2
wsM.conditional_formatting.add(f'H{RA}:H{RA+3}', CellIsRule(operator='greaterThanOrEqual', formula=['0.7'], fill=GRN))
wsM.conditional_formatting.add(f'H{RA}:H{RA+3}', CellIsRule(operator='lessThanOrEqual', formula=['0.3'], fill=RED))
note(wsM, RA + 4, 'Probabilities use a Student-t distribution on the base seasons (n−1 degrees of freedom), widened for the fact that next season is not in the sample. '
                  'Backtest error = leave-one-out RMSE at week 9 in the four "normal" seasons (see Backtest tab). Models C and D ignore the China deal and are shown as context.')
note(wsM, RA + 5, 'Model D is what most "pace" reports (incl. the PDF) show. It failed badly in 2020/21 (+31 Mt) and 2025/26 (−13 Mt) because China changed behaviour; '
                  'model A isolates China as an assumption and models only the rest of the world, which has been far more stable.')

r2 = RA + 7
section(wsM, r2, '3) China is the swing factor — probability above USDA for different China outcomes (model A)')
header(wsM, r2 + 1, ['China full-year (Mt)', 'Total expected (mln bu)', 'vs USDA (mln bu)', 'Probability ABOVE USDA'], h=30)
CH0 = r2 + 2
for i, cn in enumerate(range(18, 33)):
    r = CH0 + i
    put(wsM, r, 1, cn, '0', f=BLUE)
    put(wsM, r, 2, f'=(A{r}+{RWN}+RoWrem_mean)/Ratio/Conv', BU); put(wsM, r, 3, f'=B{r}-$B$4', SBU)
    put(wsM, r, 4, '=' + prob_ge('Need_ESR', f'A{r}+{RWN}+RoWrem_mean', f'RoWrem_sd*{SE}', 'n_yrs-1'), P0, f=font(10, True))
CH1 = CH0 + 14
wsM.conditional_formatting.add(f'D{CH0}:D{CH1}', CellIsRule(operator='greaterThanOrEqual', formula=['0.7'], fill=GRN))
wsM.conditional_formatting.add(f'D{CH0}:D{CH1}', CellIsRule(operator='lessThanOrEqual', formula=['0.3'], fill=RED))
put(wsM, CH1 + 1, 1, 'Break-even China (50% chance of beating USDA)', f=font(10, True), fill=TOT)
put(wsM, CH1 + 1, 2, f'=Need_ESR-{RWN}-RoWrem_mean', MT, f=font(11, True, NAVY), fill=TOT)
put(wsM, CH1 + 1, 3, f'="China already has "&TEXT({CNN},"0.0")&" Mt"', f=font(9, False, '404040', True), fill=TOT)
put(wsM, CH1 + 1, 4, None, fill=TOT)

r3 = CH1 + 3
section(wsM, r3, '4) How likely is any USDA number? (exceedance curve, model A with China at the assumption, and model C)')
header(wsM, r3 + 1, ['USDA export number (mln bu)', 'Bar on ESR basis (Mt)', 'P(above) — model A', 'P(above) — model C'], h=30)
EX0 = r3 + 2
for i, t in enumerate(range(1550, 2001, 25)):
    r = EX0 + i
    put(wsM, r, 1, t, BU, f=BLUE); put(wsM, r, 2, f'=A{r}*Conv*Ratio', MT)
    put(wsM, r, 3, '=' + prob_ge(f'B{r}', f'China_A+{RWN}+RoWrem_mean', f'RoWrem_sd*{SE}', 'n_yrs-1'), P0)
    put(wsM, r, 4, '=' + prob_ge(f'B{r}', f'{TTN}+Totrem_mean', f'Totrem_sd*{SE}', 'n_yrs-1'), P0)
EX1 = EX0 + len(range(1550, 2001, 25)) - 1
lc = LineChart(); lc.title = 'Probability that final exports beat a given number'; lc.y_axis.numFmt = '0%'
lc.add_data(Reference(wsM, min_col=3, max_col=4, min_row=r3 + 1, max_row=EX1), titles_from_data=True)
lc.set_categories(Reference(wsM, min_col=1, min_row=EX0, max_row=EX1)); lc.width = 18; lc.height = 9
lc.x_axis.title = 'Export number (mln bu)'; lc.y_axis.scaling.min = 0; lc.y_axis.scaling.max = 1
lc.x_axis.delete = False; lc.y_axis.delete = False
for s_ in lc.series: s_.smooth = False
wsM.add_chart(lc, f'F{r3}')
bc = BarChart(); bc.type = 'col'; bc.title = 'P(above USDA) by China outcome (model A)'; bc.y_axis.numFmt = '0%'
bc.add_data(Reference(wsM, min_col=4, min_row=r2 + 1, max_row=CH1), titles_from_data=True)
bc.set_categories(Reference(wsM, min_col=1, min_row=CH0, max_row=CH1)); bc.legend = None; bc.width = 18; bc.height = 8
bc.y_axis.scaling.min = 0; bc.y_axis.scaling.max = 1; bc.x_axis.title = 'China full-year (Mt)'
bc.x_axis.delete = False; bc.y_axis.delete = False
wsM.add_chart(bc, f'F{r2}')

r4 = EX1 + 2
section(wsM, r4, '5) Cross-check with the 10-year figures in the pace report (PDF, week 9 only)')
cx = [('10-yr avg total commitments at week 9 (Mt, 2016/17–2025/26)', 29.205, MT, True),
      ('10-yr avg full-year exports, ESR (Mt)', 51.176, MT, True),
      ('10-yr avg still to come after week 9 (Mt)', f'=B{r4+2}-B{r4+1}', MT, False),
      ('Implied full year = sold so far + 10-yr still to come (Mt, ESR)', f'={TTN}+B{r4+3}', MT, False),
      ('   in mln bu, USDA basis', f'=B{r4+4}/Ratio/Conv', BU, False),
      ('Implied full year = sold so far × 10-yr multiplier (Mt, ESR) — as in the PDF', f'={TTN}*B{r4+2}/B{r4+1}', MT, False),
      ('   in mln bu, USDA basis', f'=B{r4+6}/Ratio/Conv', BU, False)]
for i, (lb, v, fmt, isin) in enumerate(cx):
    r = r4 + 1 + i
    put(wsM, r, 1, lb, f=font(10, True)); put(wsM, r, 2, v, fmt, f=BLUE if isin else font(10, True, NAVY))
note(wsM, r4 + 8, 'Source of the two blue numbers: ESR pace report, soybeans region summary (page 21), w/e 29-Oct-26. Valid for week 9 only.')

wsM.column_dimensions['A'].width = 58
for c, w_ in zip('BCDEFGHIJ', [15, 13, 16, 11, 12, 12, 13, 12, 13]): wsM.column_dimensions[c].width = w_
wsM.column_dimensions['J'].width = 14

# ================= Backtest (static, computed here) =================
title(wsB, 'Backtest — how accurate is each model at each week of the season? (leave-one-out, Mt, ESR basis)',
      'Each base season is predicted using only the other seasons. Error = predicted − actual full-year exports. China-based models use the actual China outcome '
      '(the model takes China as an assumption, so this tests the rest-of-world part). Computed when the file was built.')
s_ = 0.6
def grp(y, w):
    c = cum.loc[(y, 'China'), w]; u = cum.loc[(y, 'Unknown'), w]; rw = cum.loc[(y, 'Rest of World'), w]
    return c + s_ * u, rw + (1 - s_) * u, c + u + rw
def finals(y):
    return fin.get((y, 'China'), 0), fin.get((y, 'Rest of World'), 0) + fin.get((y, 'Unknown'), 0)
BW = [5, 9, 13, 17, 22, 26, 30, 35, 40]
header(wsB, 4, ['Week', 'A: China known + RoW still to come — all seasons', 'A — normal seasons', 'C: total still to come — all seasons',
                'C — normal seasons', 'D: total × multiplier — all seasons', 'D — normal seasons'], h=48)
normal = [y for y in PRIOR if y not in (cs - 6, cs - 1)]
for i, w in enumerate(BW):
    errs = {'A': [], 'C': [], 'D': []}
    for y in PRIOR:
        oth = [o for o in PRIOR if o != y]
        cn, rw, t = grp(y, w); fc, fr = finals(y)
        rem_rw = np.mean([finals(o)[1] - grp(o, w)[1] for o in oth])
        rem_t = np.mean([sum(finals(o)) - grp(o, w)[2] for o in oth])
        mul_t = np.mean([sum(finals(o)) / grp(o, w)[2] for o in oth])
        act = fc + fr
        errs['A'].append((fc + rw + rem_rw - act) / 1e6); errs['C'].append((t + rem_t - act) / 1e6); errs['D'].append((t * mul_t - act) / 1e6)
    r = 5 + i
    put(wsB, r, 1, w, '0')
    for j, k in enumerate('ACD'):
        e = np.array(errs[k]); en = np.array([errs[k][PRIOR.index(y)] for y in normal])
        put(wsB, r, 2 + 2 * j, float(np.sqrt((e ** 2).mean())), '0.0', f=BLUE); put(wsB, r, 3 + 2 * j, float(np.sqrt((en ** 2).mean())), '0.0', f=BLUE)
note(wsB, 5 + len(BW), f'"Normal seasons" = {", ".join(lab(y) for y in normal)} (excludes 2020/21 record China buying and 2025/26 China absent).')
r = 7 + len(BW)
put(wsB, r, 1, f'Errors by season at week {WK} (Mt)', f=font(11, True, NAVY), border=False)
header(wsB, r + 1, ['Season', 'Model A', 'Model C', 'Model D', 'Actual full-year'], h=24)
for i, y in enumerate(PRIOR):
    oth = [o for o in PRIOR if o != y]
    cn, rw, t = grp(y, WK); fc, fr = finals(y)
    rem_rw = np.mean([finals(o)[1] - grp(o, WK)[1] for o in oth]); rem_t = np.mean([sum(finals(o)) - grp(o, WK)[2] for o in oth])
    mul_t = np.mean([sum(finals(o)) / grp(o, WK)[2] for o in oth]); act = fc + fr
    rr = r + 2 + i
    put(wsB, rr, 1, lab(y))
    for j, v in enumerate([fc + rw + rem_rw - act, t + rem_t - act, t * mul_t - act, act]):
        put(wsB, rr, 2 + j, float(v) / 1e6, '+0.0;-0.0;0.0' if j < 3 else '0.0', f=BLUE)
wsB.column_dimensions['A'].width = 14
for c in 'BCDEFG': wsB.column_dimensions[c].width = 16
lb_ = LineChart(); lb_.title = 'Backtest error by week, normal seasons (Mt)'; lb_.y_axis.numFmt = '0.0'
for col in (3, 5, 7): lb_.add_data(Reference(wsB, min_col=col, min_row=4, max_row=4 + len(BW)), titles_from_data=True)
lb_.set_categories(Reference(wsB, min_col=1, min_row=5, max_row=4 + len(BW))); lb_.width = 16; lb_.height = 8
lb_.x_axis.title = 'Week'; lb_.x_axis.delete = False; lb_.y_axis.delete = False
for s2 in lb_.series: s2.smooth = False
wsB.add_chart(lb_, 'I4')

# ================= Regions: data =================
RMAP = {'CHINA, PEOPLES REPUBLIC OF': 'China', 'MEXICO': 'Mexico', 'CANADA': 'Canada', 'UNKNOWN': 'Unknown'}
for lst, nm in [(['JAPAN', 'TAIWAN', 'KOREA, REPUBLIC OF', 'HONG KONG'], 'East Asia ex-China'),
                (['INDONESIA', 'VIETNAM', 'MALAYSIA', 'THAILAND', 'PHILIPPINES', 'BURMA', 'CAMBODIA', 'LAOS', 'SINGAPORE'], 'Southeast Asia'),
                (['BANGLADESH', 'PAKISTAN', 'NEPAL', 'SRI LANKA'], 'South Asia'),
                (['EGYPT', 'ALGERIA', 'TURKEY', 'TUNISIA', 'MOROCCO', 'IRAN', 'IRAQ', 'ISRAEL', 'JORDAN', 'LEBANON', 'LIBYA', 'SAUDI ARABIA', 'QATAR'], 'Middle East & N. Africa'),
                (['SPAIN', 'NETHERLANDS', 'PORTUGAL', 'ITALY', 'GERMANY', 'UNITED KINGDOM', 'GREECE', 'BELGIUM-LUXEMBOURG', 'FRANCE', 'IRELAND', 'NORWAY',
                  'ROMANIA', 'SWITZERLAND', 'UKRAINE', 'GEORGIA, REPUBLIC OF'], 'Europe'),
                (['COSTA RICA', 'EL SALVADOR', 'GUATEMALA', 'HONDURAS', 'NICARAGUA', 'PANAMA', 'DOMINICAN REPUBLIC', 'CUBA', 'BARBADOS',
                  'LEEWARD-WINDWARD ISLANDS', 'VIRGIN ISLANDS OF THE U.S.'], 'Central America & Caribbean'),
                (['COLOMBIA', 'VENEZUELA', 'PERU', 'ARGENTINA', 'BRAZIL', 'CHILE'], 'South America'),
                (['GHANA', 'NIGER', 'NIGERIA', 'SENEGAL', 'SOUTH AFRICA, REPUBLIC OF'], 'Sub-Saharan Africa')]:
    for c_ in lst: RMAP[c_] = nm
REGS = ['China', 'Mexico', 'Canada', 'East Asia ex-China', 'Southeast Asia', 'South Asia', 'Middle East & N. Africa', 'Europe',
        'Central America & Caribbean', 'South America', 'Sub-Saharan Africa', 'Other']
ins['R'] = ins.Country.map(RMAP).fillna('Other')
rcum = ins.pivot_table(index=['Y', 'R'], columns='WeekNumber', values='TotalCommitment', aggfunc='sum').reindex(columns=range(1, NW + 1)).fillna(0)
rfin = ins[ins.Y < cs].sort_values('date').groupby(['Y', 'Country']).tail(1).groupby(['Y', 'R']).AccumulatedExports.sum()
# 10-yr region figures from the ESR pace report (page 21, week 9): 10y avg commitments at wk 9 and 10y avg full-year (kt)
PDF10 = {'China': (13239, 24964), 'Mexico': (2403, 4699), 'Canada': (117, 242), 'East Asia ex-China': (1736, 4244),
         'Southeast Asia': (1480, 4370), 'South Asia': (860, 1718), 'Middle East & N. Africa': (1223, 4105), 'Europe': (1398, 5452),
         'Central America & Caribbean': (158, 395), 'South America': (308, 951), 'Sub-Saharan Africa': (7, 33), 'Other': (1, 2),
         'Unknown': (6274, 0)}
wsRD = wb.create_sheet('Weekly by Region')
title(wsRD, 'Cumulative commitments by region and in-season week (Mt), with full-year exports',
      f'Hardcoded from the ESR file (blue). Regions follow the pace report. {CUR} through week {WK}.')
header(wsRD, 4, ['Marketing year', 'Region', 'Full-year exports (Mt)'] + [f'W{w}' for w in range(1, NW + 1)])
r = 5
for y in YEARS:
    for g in REGS + ['Unknown']:
        put(wsRD, r, 1, lab(y), f=BLUE, border=False); put(wsRD, r, 2, g, f=BLUE, border=False)
        put(wsRD, r, 3, None if y == cs else float(rfin.get((y, g), 0)) / 1e6, MT, f=BLUE, border=False)
        for w in range(1, NW + 1):
            if y == cs and w > WK: continue
            x = wsRD.cell(r, 3 + w, float(rcum.loc[(y, g), w]) / 1e6 if (y, g) in rcum.index else 0.0); x.number_format = MT; x.font = BLUE
        r += 1
RDL = r - 1
for n_, c_ in [('RD_MY', 'A'), ('RD_G', 'B'), ('RD_F', 'C')]: name(wb, n_, 'Weekly by Region', f'${c_}$5:${c_}${RDL}')
name(wb, 'RD_W', 'Weekly by Region', f'$D$5:${L(3+NW)}${RDL}')
wsRD.column_dimensions['A'].width = 12; wsRD.column_dimensions['B'].width = 26; wsRD.column_dimensions['C'].width = 12
wsRD.freeze_panes = 'D5'
def rwk(g, my): return f'SUMIFS(INDEX(RD_W,0,Wk_N),RD_G,{g},RD_MY,{my})'
def rfy(g, my): return f'SUMIFS(RD_F,RD_G,{g},RD_MY,{my})'

# backtest by week (named regions; LOO over all prior seasons) -> RMSE used as the model's uncertainty
def yt(y, w): return {g: (float(rcum.loc[(y, g), w]) if (y, g) in rcum.index else 0.0) / 1e6 for g in REGS}
def fn(y): return {g: float(rfin.get((y, g), 0)) / 1e6 for g in REGS}
FT = {y: sum(float(rfin.get((y, g), 0)) for g in REGS + ['Unknown']) / 1e6 for y in PRIOR}
rmse_add, rmse_mul, rmse_addC, rmse_mulC = {}, {}, {}, {}
for w in range(1, NW + 1):
    ea, em, eac, emc = [], [], [], []
    for y in PRIOR:
        oth = [o for o in PRIOR if o != y]
        cur_ = yt(y, w)
        pa = pm = 0.0
        for g in REGS:
            ay = np.mean([yt(o, w)[g] for o in oth]); af = np.mean([fn(o)[g] for o in oth])
            pa += cur_[g] + af - ay
            pm += cur_[g] * af / ay if ay > 0 else 0
        ea.append(pa - FT[y]); em.append(pm - FT[y])
        ayc = np.mean([yt(o, w)['China'] for o in oth]); afc = np.mean([fn(o)['China'] for o in oth])
        cpa = cur_['China'] + afc - ayc; cpm = cur_['China'] * afc / ayc if ayc > 0 else 0
        eac.append(pa - cpa + fn(y)['China'] - FT[y]); emc.append(pm - cpm + fn(y)['China'] - FT[y])
    rmse_add[w] = float(np.sqrt(np.mean(np.square(ea)))); rmse_mul[w] = float(np.sqrt(np.mean(np.square(em))))
    rmse_addC[w] = float(np.sqrt(np.mean(np.square(eac)))); rmse_mulC[w] = float(np.sqrt(np.mean(np.square(emc))))
c0b = 10
put(wsB, 3, c0b, 'Regional pace methods — leave-one-out RMSE by week (Mt, all seasons; used as the uncertainty on the Pace Model tab)', f=font(10, True, NAVY), border=False)
header(wsB, 4, ['Week', 'Remaining-sales method', 'Multiplier method', 'Remaining-sales, China known', 'Multiplier, China known'], c0b, h=48)
for w in range(1, NW + 1):
    put(wsB, 4 + w, c0b, w, '0'); put(wsB, 4 + w, c0b + 1, rmse_add[w], '0.0', f=BLUE); put(wsB, 4 + w, c0b + 2, rmse_mul[w], '0.0', f=BLUE)
    put(wsB, 4 + w, c0b + 3, rmse_addC[w], '0.0', f=BLUE); put(wsB, 4 + w, c0b + 4, rmse_mulC[w], '0.0', f=BLUE)
name(wb, 'RMSE_addC', 'Backtest', f'${L(c0b+3)}$5:${L(c0b+3)}${4+NW}'); name(wb, 'RMSE_mulC', 'Backtest', f'${L(c0b+4)}$5:${L(c0b+4)}${4+NW}')
name(wb, 'RMSE_add', 'Backtest', f'${L(c0b+1)}$5:${L(c0b+1)}${4+NW}'); name(wb, 'RMSE_mul', 'Backtest', f'${L(c0b+2)}$5:${L(c0b+2)}${4+NW}')
wsB.column_dimensions[L(c0b)].width = 8; wsB.column_dimensions[L(c0b+1)].width = 14; wsB.column_dimensions[L(c0b+2)].width = 14
for _ch in wsB._charts: _ch.anchor = 'N4'

# ================= Pace Model (no China assumption) =================
wsP = wb.create_sheet('Pace Model', index=0)
title(wsP, f'Pace view: probability that {CUR} US soybean exports beat USDA, and how each region gets us there',
      'Every region is projected from its own 5-yr and 10-yr seasonal pace; China can be fixed in B9 (leave blank to project China from its pace too). USDA number, ratio and week N are set on the "Model" tab.')
put(wsP, 4, 1, 'USDA export forecast (mln bu)', f=font(10, True)); put(wsP, 4, 2, '=Model!B4', BU, f=GREENF)
put(wsP, 5, 1, 'USDA forecast on ESR basis (Mt) — the bar to clear', f=font(10, True)); put(wsP, 5, 2, '=Need_ESR', MT, f=GREENF)
put(wsP, 6, 1, 'Week N', f=font(10, True)); put(wsP, 6, 2, '=Wk_N', '0', f=GREENF)
put(wsP, 7, 1, '5-yr base seasons', f=font(10, True)); put(wsP, 7, 2, f'{lab(PRIOR[-5])} – {lab(PRIOR[-1])}')
put(wsP, 8, 1, '10-yr figures', f=font(10, True)); put(wsP, 8, 2, 'ESR pace report p.21 (2016/17–2025/26), valid for week 9 only')
put(wsP, 9, 1, 'China full-year (Mt, ESR) — blank = use China\'s own pace', f=font(10, True)); put(wsP, 9, 2, 25, MT, f=BLUE, fill=YEL)
put(wsP, 9, 3, 'Fixed at 25 Mt: China\'s uncertainty is removed; the probability then reflects the rest of the world only.', f=font(9, False, '404040', True), border=False)
name(wb, 'China_fix', 'Pace Model', '$B$9')
B5Y = [lab(y) for y in PRIOR[-5:]]
# regional table first (totals reference it)
rr0 = 22
put(wsP, rr0, 1, '2) Regional view — what each region is on track for, and how that compares with its share of USDA\'s number (Mt, ESR basis)', f=font(12, True, NAVY), border=False)
RH = ['Region', 'Committed so far', '5-yr avg committed by wk N', '5-yr avg full year', '5-yr usual still to come', 'Expected full year (5-yr)',
      '10-yr avg committed by wk 9', '10-yr avg full year', 'Expected full year (10-yr)', 'Share of USDA total (5-yr share)', 'Gap vs USDA share (5-yr)',
      'Gap vs USDA share (10-yr)', 'Still to sell to reach USDA share', 'Needed ÷ usual still to come', 'Chance region reaches its USDA share',
      'Gap in mln bu (5-yr)', 'Std dev of still-to-come (6 seasons)', 'Expected full year — multiplier (5-yr)']
header(wsP, rr0 + 1, RH, h=56)
hb = len(RH) + 2
for k in range(6):
    x = wsP.cell(rr0 + 1, hb + k, f'Still to come {lab(PRIOR[k])}'); x.font = font(8, True, '595959'); x.alignment = Alignment(wrap_text=True)
R1 = rr0 + 2
for i, g in enumerate(REGS):
    r = R1 + i; G = f'$A{r}'
    put(wsP, r, 1, g, f=font(10, True))
    put(wsP, r, 2, f'=IF(Wk_N>{WK},"",{rwk(G, chr(34)+CUR+chr(34))})', MT, fill=CURF)
    put(wsP, r, 3, '=(' + '+'.join(rwk(G, f'"{m}"') for m in B5Y) + ')/5', MT)
    put(wsP, r, 4, '=(' + '+'.join(rfy(G, f'"{m}"') for m in B5Y) + ')/5', MT)
    put(wsP, r, 5, f'=D{r}-C{r}', MT)
    FX = (lambda e: f'=IF(ISNUMBER(China_fix),China_fix,{e})') if g == 'China' else (lambda e: '=' + e)
    put(wsP, r, 6, FX(f'B{r}+E{r}'), MT, f=font(10, True), fill=CURF if g == 'China' else None)
    a10, f10 = PDF10[g]
    put(wsP, r, 7, a10 / 1000, MT, f=BLUE); put(wsP, r, 8, f10 / 1000, MT, f=BLUE)
    put(wsP, r, 9, FX(f'IF(Wk_N=9,B{r}+H{r}-G{r},"n/a (wk 9 only)")'), MT, f=font(10, True), fill=CURF if g == 'China' else None)
    put(wsP, r, 10, f'=$B$5*D{r}/$D${R1+len(REGS)}', MT)
    put(wsP, r, 11, f'=F{r}-J{r}', SBU.replace('0;', '0.00;').replace('0;0', '0.00;0') if False else '+0.00;-0.00;0.00', f=font(10, True))
    put(wsP, r, 12, f'=IF(ISNUMBER(I{r}),I{r}-J{r},"")', '+0.00;-0.00;0.00')
    put(wsP, r, 13, f'=MAX(0,J{r}-B{r})', MT)
    put(wsP, r, 14, f'=IF(E{r}<=0,"",M{r}/E{r})', '0%')
    for k, y in enumerate(PRIOR):
        put(wsP, r, hb + k, f'={rfy(G, chr(34)+lab(y)+chr(34))}-{rwk(G, chr(34)+lab(y)+chr(34))}', MT, f=font(8, False, '595959'), border=False)
    put(wsP, r, 17, f'=STDEV({L(hb)}{r}:{L(hb+5)}{r})', MT)
    pf = f'IF(Q{r}=0,IF(E{r}>=M{r},1,0),' + prob_ge(f'M{r}', f'E{r}', f'Q{r}*SQRT(1+1/6)', 5) + ')'
    put(wsP, r, 15, (f'=IF(ISNUMBER(China_fix),"fixed",{pf})' if g == 'China' else '=' + pf), P0, f=font(10, True), al=Alignment(horizontal='right'))
    put(wsP, r, 16, f'=K{r}/Ratio/Conv', '+#,##0;-#,##0;0', f=font(10, True))
    put(wsP, r, 18, FX(f'IF(C{r}>0,B{r}*D{r}/C{r},0)'), MT)
R2 = R1 + len(REGS) - 1
RT = R2 + 1
put(wsP, RT, 1, 'TOTAL (named destinations)', f=font(10, True), fill=TOT)
for c in [2, 3, 4, 5, 6, 7, 8, 10, 11, 13, 16, 18]:
    put(wsP, RT, c, f'=SUM({L(c)}{R1}:{L(c)}{R2})', '+#,##0;-#,##0;0' if c == 16 else ('+0.00;-0.00;0.00' if c == 11 else MT), f=font(10, True), fill=TOT)
put(wsP, RT, 9, f'=IF(Wk_N=9,SUM(I{R1}:I{R2}),"n/a")', MT, f=font(10, True), fill=TOT)
put(wsP, RT, 12, f'=IF(ISNUMBER(I{RT}),I{RT}-J{RT},"")', '+0.00;-0.00;0.00', f=font(10, True), fill=TOT)
for c in (14, 15, 17): put(wsP, RT, c, None, fill=TOT)
RU = RT + 1
put(wsP, RU, 1, 'Memo: Unknown (not counted — its usual level is already inside each region\'s "still to come")', f=font(9, False, '404040', True))
put(wsP, RU, 2, f'=IF(Wk_N>{WK},"",{rwk(chr(34)+"Unknown"+chr(34), chr(34)+CUR+chr(34))})', MT)
put(wsP, RU, 3, '=(' + '+'.join(rwk('"Unknown"', f'"{m}"') for m in B5Y) + ')/5', MT)
put(wsP, RU, 7, PDF10['Unknown'][0] / 1000, MT, f=BLUE)
note(wsP, RU + 1, '"Share of USDA total": USDA publishes no split by destination, so its total is shared out by each region\'s 5-yr share of exports. '
                  '"Gap" = expected full year − that share: positive = the region is on track to deliver more than a USDA-consistent amount.')
note(wsP, RU + 2, '"Chance region reaches its USDA share" uses the spread of that region\'s "still to come" across the six seasons (Student-t, 5 df).')
for rg_ in (f'K{R1}:K{R2}', f'L{R1}:L{R2}', f'P{R1}:P{R2}'):
    wsP.conditional_formatting.add(rg_, CellIsRule(operator='greaterThan', formula=['0'], fill=GRN))
    wsP.conditional_formatting.add(rg_, CellIsRule(operator='lessThan', formula=['0'], fill=RED))
wsP.conditional_formatting.add(f'O{R1}:O{R2}', CellIsRule(operator='greaterThanOrEqual', formula=['0.7'], fill=GRN))
wsP.conditional_formatting.add(f'O{R1}:O{R2}', CellIsRule(operator='lessThanOrEqual', formula=['0.3'], fill=RED))

# totals section (rows 10-20)
t0 = 10
put(wsP, t0, 1, '1) Probability that full-year exports beat USDA — pace methods, all destinations', f=font(12, True, NAVY), border=False)
TH = ['Method', 'Expected full year (Mt, ESR)', 'Expected (mln bu, USDA basis)', 'vs USDA (mln bu)', 'Historical error ±1 sd (Mt)',
      '80% range low (mln bu)', '80% range high (mln bu)', 'Probability ABOVE USDA', 'Probability BELOW USDA']
header(wsP, t0 + 1, TH, h=48)
SDA, SDM = 'IF(ISNUMBER(China_fix),INDEX(RMSE_addC,Wk_N),INDEX(RMSE_add,Wk_N))', 'IF(ISNUMBER(China_fix),INDEX(RMSE_mulC,Wk_N),INDEX(RMSE_mul,Wk_N))'
tm = [('Remaining sales, 5-yr (RECOMMENDED)', f'=F{RT}', SDA),
      ('Remaining sales, 10-yr (pace report)', f'=IF(ISNUMBER(I{RT}),I{RT},"")', SDA),
      ('Multiplier ("implied FY"), 5-yr', f'=R{RT}', SDM),
      ('Multiplier ("implied FY"), 10-yr — as in the pace report', f'=IF(Wk_N=9,IF(ISNUMBER(China_fix),China_fix+SUMPRODUCT(B{R1+1}:B{R2},H{R1+1}:H{R2}/G{R1+1}:G{R2}),SUMPRODUCT(B{R1}:B{R2},H{R1}:H{R2}/G{R1}:G{R2})),"")', SDM)]
for i, (lb, ex, sd) in enumerate(tm):
    r = t0 + 2 + i
    put(wsP, r, 1, lb, f=font(10, i == 0)); put(wsP, r, 2, ex, MT, f=font(10, True)); put(wsP, r, 5, '=' + sd, '0.0')
    put(wsP, r, 3, f'=IF(B{r}="","",B{r}/Ratio/Conv)', BU, f=font(10, True)); put(wsP, r, 4, f'=IF(C{r}="","",C{r}-$B$4)', SBU, f=font(10, True))
    put(wsP, r, 6, f'=IF(B{r}="","",(B{r}-TINV(0.2,5)*E{r}*SQRT(1+1/6))/Ratio/Conv)', BU)
    put(wsP, r, 7, f'=IF(B{r}="","",(B{r}+TINV(0.2,5)*E{r}*SQRT(1+1/6))/Ratio/Conv)', BU)
    put(wsP, r, 8, f'=IF(B{r}="","",' + prob_ge('Need_ESR', f'B{r}', f'E{r}*SQRT(1+1/6)', 5) + ')', P0, f=font(12, True, NAVY), al=Alignment(horizontal='center'))
    put(wsP, r, 9, f'=IF(H{r}="","",1-H{r})', P0, al=Alignment(horizontal='center'))
T1 = t0 + 2
wsP.conditional_formatting.add(f'H{T1}:H{T1+3}', CellIsRule(operator='greaterThanOrEqual', formula=['0.7'], fill=GRN))
wsP.conditional_formatting.add(f'H{T1}:H{T1+3}', CellIsRule(operator='lessThanOrEqual', formula=['0.3'], fill=RED))
rs = t0 + 7
put(wsP, rs, 1, 'Season-by-season test: this year\'s bookings + each past season\'s regional "still to come"', f=font(10, True))
header(wsP, rs + 1, ['Season pattern'] + [lab(y) for y in PRIOR] + ['Seasons above USDA'], h=24)
put(wsP, rs + 2, 1, 'Full year (mln bu, USDA basis)', f=font(10, True))
for k in range(6):
    put(wsP, rs + 2, 2 + k, f'=IF(ISNUMBER(China_fix),China_fix+B{RT}-B{R1}+SUM({L(hb+k)}{R1+1}:{L(hb+k)}{R2}),B{RT}+SUM({L(hb+k)}{R1}:{L(hb+k)}{R2}))/Ratio/Conv', BU)
put(wsP, rs + 2, 8, f'=COUNTIF(B{rs+2}:G{rs+2},">"&$B$4)&" of 6"', f=font(11, True, NAVY))
wsP.conditional_formatting.add(f'B{rs+2}:G{rs+2}', CellIsRule(operator='greaterThan', formula=['$B$4'], fill=GRN))
wsP.conditional_formatting.add(f'B{rs+2}:G{rs+2}', CellIsRule(operator='lessThanOrEqual', formula=['$B$4'], fill=RED))

wsP.column_dimensions['A'].width = 46
for c in range(2, len(RH) + 1): wsP.column_dimensions[L(c)].width = 12
for k in range(6): wsP.column_dimensions[L(hb + k)].width = 9
gc = BarChart(); gc.type = 'bar'; gc.title = 'Gap vs USDA-consistent share by region (mln bu, 5-yr pace)'; gc.y_axis.numFmt = '+#,##0;-#,##0'
gc.add_data(Reference(wsP, min_col=16, min_row=rr0 + 1, max_row=R2), titles_from_data=True)
gc.set_categories(Reference(wsP, min_col=1, min_row=R1, max_row=R2)); gc.legend = None
gc.x_axis.scaling.orientation = 'maxMin'; gc.x_axis.tickLblPos = 'low'; gc.width = 18; gc.height = 10
gc.x_axis.delete = False; gc.y_axis.delete = False
wsP.add_chart(gc, f'A{RU + 5}')

# ================= Method =================
title(wsN, 'Method — how the probability is calculated', 'Plain-language summary; all numbers on the Model tab are live formulas.')
steps = [
    ('Question', 'Given what has been sold so far, how likely is it that US soybean exports end the season above (or below) USDA\'s forecast?'),
    ('Step 1 — split the buyers', 'China (plus 60% of "Unknown") and the Rest of World (plus 40% of Unknown) behave very differently, so they are modelled separately.'),
    ('Step 2 — China as an assumption', 'China\'s purchases this season are driven by policy (purchase commitment), not by normal seasonality. We take China\'s full year as an input '
                                         '(default 25 Mt) and show how the probability changes for other China outcomes.'),
    ('Step 3 — Rest of World from history', 'For each base season we measure how much the Rest of World bought AFTER week N ("still to come"). '
                                            'This year\'s Rest of World = sold so far + a draw from that history (average and spread).'),
    ('Step 4 — same basis as USDA', 'Export sales data runs ~2–3% below USDA\'s official export number. We convert with the ESR ÷ USDA ratio of past seasons.'),
    ('Step 5 — probability', 'With only a few seasons of history we use a Student-t distribution (wider tails than a normal curve) to turn the average and spread '
                             'into the chance of beating USDA\'s number.'),
    ('Why not just "pace"', 'The classic pace method (sold so far ÷ share usually sold by now) gives very unstable answers when China changes behaviour: '
                            'its backtest error was +31 Mt in 2020/21 and −13 Mt in 2025/26. Modelling only the Rest of World from history cut the error to ~±3 Mt.'),
    ('Limitations', 'Six seasons of history (the ESR file starts in 2020/21). The China number is an assumption. USDA\'s past finals on the basis tab are '
                    'approximate and should be checked. Cancellations, logistics or a crop shortfall are not modelled explicitly.'),
    ('Weekly update', 'Re-run the build script on the new ESR export: week N, commitments and history update automatically.'),
]
for i, (k, v) in enumerate(steps):
    r = 4 + i
    a = wsN.cell(r, 1, k); a.font = font(11, True, NAVY); a.alignment = Alignment(vertical='top')
    b = wsN.cell(r, 2, v); b.font = font(11); b.alignment = Alignment(wrap_text=True, vertical='top')
    wsN.row_dimensions[r].height = 46
wsN.column_dimensions['A'].width = 32; wsN.column_dimensions['B'].width = 120

for ws in wb: ws.sheet_view.showGridLines = False
wb.save(out)
print('CUR', CUR, 'WK', WK, 'years', [lab(y) for y in PRIOR])
