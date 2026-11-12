"""US vs Brazil soybean origin switching by destination (10 seasons, monthly customs exports).
Usage: python build_switching.py <trade csv> <out.xlsx> [base_case_mbu]
Input: monthly export file (Product SBS, reporters United States & Brazil, crop year Sep-Aug, Quantity in kt).
Forecast rows (current season) are ignored."""
import sys
import numpy as np
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.utils import get_column_letter as L
from openpyxl.formatting.rule import CellIsRule, FormulaRule
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.datavalidation import DataValidation

src, out = sys.argv[1], sys.argv[2]
BASE_MBU = float(sys.argv[3]) if len(sys.argv) > 3 else 1825
ESR_CSV = sys.argv[4] if len(sys.argv) > 4 else None
ESR = {}
if ESR_CSV:   # this season's US position by destination from the weekly export sales (ESR) data
    e = pd.read_csv(ESR_CSV, thousands=','); e.columns = [c.strip() for c in e.columns]
    e['C'] = e.CountryName.str.strip(); e['dt'] = pd.to_datetime(e.WeekendingDate)
    e['Y'] = e.ReportingMarketingYearName.str.extract(r'Sep (\d{4})').astype(int)
    e = e[(e.dt >= pd.to_datetime(e.Y.astype(str) + '-08-31')) & (e.dt <= pd.to_datetime((e.Y + 1).astype(str) + '-09-07'))]
    ecs = int(e.loc[e.dt.idxmax(), 'Y']); EWK = int(e[e.Y == ecs].WeekNumber.max())
    now_ = e[(e.Y == ecs) & (e.WeekNumber == EWK)].groupby('C').TotalCommitment.sum() / 1000
    fin_ = e[(e.Y < ecs) & (e.Y >= ecs - 5)].sort_values('dt').groupby(['Y', 'C']).tail(1).groupby('C').AccumulatedExports.sum() / 5000
    ESR = {'now': now_, 'fy5': fin_, 'wk': EWK}
ESRNAME = {'South Korea': 'KOREA, REPUBLIC OF'}
def esr_name(p): return ESRNAME.get(p, p.upper())
REG = {'China': 'China', 'Mexico': 'Mexico', 'Canada': 'Canada'}
for _l, _r in [(['Japan', 'South Korea', 'Taiwan', 'Hong Kong'], 'East Asia ex-China'),
               (['Indonesia', 'Malaysia', 'Myanmar', 'Philippines', 'Thailand', 'Vietnam', 'Singapore'], 'Southeast Asia'),
               (['Bangladesh', 'India', 'Pakistan', 'Sri Lanka'], 'South Asia'),
               (['Algeria', 'Egypt', 'Morocco', 'Tunisia', 'Iran', 'Iraq', 'Israel', 'Jordan', 'Lebanon', 'Saudi Arabia', 'Syrian Arab Republic',
                 'United Arab Emirates', 'Yemen', 'Turkey'], 'Middle East & N. Africa'),
               (['Austria', 'Belgium', 'Bulgaria', 'Croatia', 'Cyprus', 'Czech Republic', 'Denmark', 'Estonia', 'Finland', 'France', 'Germany', 'Greece',
                 'Hungary', 'Ireland', 'Italy', 'Latvia', 'Lithuania', 'Luxembourg', 'Malta', 'Netherlands', 'Poland', 'Portugal', 'Romania', 'Slovakia',
                 'Slovenia', 'Spain', 'Sweden', 'United Kingdom', 'Norway'], 'Europe'),
               (['Russian Federation', 'Ukraine', 'Uzbekistan'], 'Russia & CIS'),
               (['Costa Rica', 'Cuba', 'Dominican Republic', 'El Salvador', 'Guatemala', 'Honduras', 'Jamaica', 'Nicaragua', 'Panama'], 'Central America & Caribbean'),
               (['Argentina', 'Chile', 'Colombia', 'Ecuador', 'Peru', 'Uruguay', 'Venezuela'], 'South America'),
               (['Ethiopia', 'South Africa', 'Tanzania', 'Ghana', 'Nigeria', 'Senegal', 'Togo'], 'Sub-Saharan Africa')]:
    for _n in _l: REG[_n] = _r
REGS = ['China', 'Mexico', 'Canada', 'East Asia ex-China', 'Southeast Asia', 'South Asia', 'Middle East & N. Africa', 'Europe', 'Russia & CIS',
        'Central America & Caribbean', 'South America', 'Sub-Saharan Africa', 'Other']
def reg(p): return REG.get(p, 'Other')
d = pd.read_csv(src)
d = d[(d.Product == 'SBS') & (d.Reporter_MappedName.isin(['United States', 'Brazil'])) & (~d.Forecast)].copy()
d['O'] = d.Reporter_MappedName.map({'United States': 'US', 'Brazil': 'BR'}); d['M'] = d.Month.astype(int)
last = sorted(d.CropYear.unique())[-1]
SEAS = sorted(d.CropYear.unique())[-10:]          # last 10 complete seasons
d = d[d.CropYear.isin(SEAS)]
B5 = SEAS[-5:]
MONTHS = [9, 10, 11, 12, 1, 2, 3, 4, 5, 6, 7, 8]; MN = ['Sep', 'Oct', 'Nov', 'Dec', 'Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug']
CONV = 0.0272155

ann = d.pivot_table(index=['Partner_Name', 'CropYear'], columns='O', values='Quantity', aggfunc='sum').fillna(0)
ann['T'] = ann.US + ann.BR; ann['S'] = ann.US / ann['T'].where(ann['T'] > 0)
brx = d[(d.O == 'BR') & (d.Partner_Name != 'China')].groupby('CropYear').Quantity.sum().reindex(SEAS) / 1000
X = np.column_stack([np.ones(len(SEAS)), brx.values, np.arange(len(SEAS))])
half = d.assign(H=np.where(d.M.isin([9, 10, 11, 12, 1, 2]), 'SF', 'MA')).pivot_table(index=['Partner_Name', 'H'], columns='O', values='Quantity', aggfunc='sum').fillna(0)
rem = d[d.CropYear.isin(B5) & ~d.M.isin([9, 10])].groupby('Partner_Name').Quantity.sum() / 5
rows = []
for p, g in ann.groupby(level=0):
    g = g.droplevel(0).reindex(SEAS).fillna(0)
    if g.US.mean() < 150 or p == 'China': continue
    S = g.S.fillna(0).values
    lnUS = np.log(g.US.clip(lower=1)); lnS = lnUS - np.log(g['T'].clip(lower=1))
    vfs = float(np.var(lnS) / np.var(lnUS)) if np.var(lnUS) > 0 else np.nan
    b = np.linalg.lstsq(X, S, rcond=None)[0]; res = S - X @ b
    se = np.sqrt((res @ res) / (len(SEAS) - 3) * np.linalg.inv(X.T @ X)[1, 1])
    braw = np.polyfit(brx.values, S, 1)[0]
    def hs(h):
        try: z = half.loc[(p, h)]; return z.US / (z.US + z.BR) if (z.US + z.BR) > 0 else np.nan
        except KeyError: return np.nan
    corr = float(np.corrcoef(g.US, g.BR)[0, 1]) if g.BR.std() > 0 else np.nan
    rows.append(dict(dest=p, US=g.US.mean(), US5=g.US.loc[B5].mean(), BR=g.BR.mean(), share=np.nanmean(g.S), smin=np.nanmin(g.S), smax=np.nanmax(g.S),
                     sf=hs('SF'), ma=hs('MA'), rho=corr, vfs=vfs, raw=braw * 100, detr=b[1] * 100, t=b[1] / se if se > 0 else 0,
                     comb=float(rem.get(p, 0))))
T = pd.DataFrame(rows).set_index('dest')
T['class'] = np.where((T.vfs >= 0.6) & (T.rho < 0), 'Switcher', np.where(T.vfs < 0.3, 'Demand-driven', 'Mixed'))
T = T.sort_values('US5', ascending=False)

# ---------------- styles ----------------
F = 'Arial'; NAVY = '1F4E78'
HF = PatternFill('solid', fgColor=NAVY); TOT = PatternFill('solid', fgColor='EDEDED'); YEL = PatternFill('solid', fgColor='FFFF00')
GRN = PatternFill('solid', fgColor='C6EFCE'); RED = PatternFill('solid', fgColor='FFC7CE'); AMB = PatternFill('solid', fgColor='FFEB9C')
thin = Side(style='thin', color='D9D9D9'); BOX = Border(bottom=thin)
KT = '#,##0'; P0 = '0%'
def font(sz=10, b=False, col='000000', i=False): return Font(name=F, size=sz, bold=b, color=col, italic=i)
BLUE = font(10, False, '0000FF')
def put(ws, r, c, v, fmt=None, f=None, fill=None, al=None, border=True):
    x = ws.cell(r, c, v); x.font = f or font()
    if fmt: x.number_format = fmt
    if fill: x.fill = fill
    if al: x.alignment = al
    if border: x.border = BOX
    return x
def title(ws, t, sub):
    ws['A1'] = t; ws['A1'].font = font(15, True, NAVY); ws['A2'] = sub; ws['A2'].font = font(9, False, '595959', True)
def header(ws, r, labels, c0=1, h=40):
    for i, t in enumerate(labels):
        x = ws.cell(r, c0 + i, t); x.font = font(9, True, 'FFFFFF'); x.fill = HF
        x.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    ws.row_dimensions[r].height = h
def note(ws, r, t, c=1, sz=9): x = ws.cell(r, c, t); x.font = font(sz, False, '404040', True); return x
def section(ws, r, t): ws.cell(r, 1, t).font = font(12, True, NAVY)

wb = Workbook()
wsS = wb.active; wsS.title = 'Switching Ranking'
wsC = wb.create_sheet('Country Profile')
wsH = wb.create_sheet('History')
wsD = wb.create_sheet('Data')
wsN = wb.create_sheet('Method')

# ================= Data (monthly, by destination) =================
title(wsD, f'Monthly soybean exports by destination: United States and Brazil, {SEAS[0]}–{SEAS[-1]} (kt)',
      'From the uploaded monthly export file (customs basis; US totals match USDA\'s official finals). Crop year Sep–Aug. Forecast rows excluded.')
header(wsD, 4, ['Origin', 'Destination', 'Crop year'] + MN + ['Total'])
mt = d.pivot_table(index=['O', 'Partner_Name', 'CropYear'], columns='M', values='Quantity', aggfunc='sum').reindex(columns=MONTHS).fillna(0)
keep = set(T.index) | {'China'}
r = 5
for (o, p, cy), row in mt.iterrows():
    if p not in keep: continue
    put(wsD, r, 1, o, f=BLUE, border=False); put(wsD, r, 2, p, f=BLUE, border=False); put(wsD, r, 3, cy, f=BLUE, border=False)
    for j, m in enumerate(MONTHS): x = wsD.cell(r, 4 + j, float(row[m])); x.number_format = KT; x.font = BLUE
    x = wsD.cell(r, 16, f'=SUM(D{r}:O{r})'); x.number_format = KT
    r += 1
# Brazil total ex-China rows (availability) per season
for cy in SEAS:
    put(wsD, r, 1, 'BR', f=BLUE, border=False); put(wsD, r, 2, 'ALL EX-CHINA', f=BLUE, border=False); put(wsD, r, 3, cy, f=BLUE, border=False)
    z = d[(d.O == 'BR') & (d.Partner_Name != 'China') & (d.CropYear == cy)].groupby('M').Quantity.sum().reindex(MONTHS).fillna(0)
    for j, m in enumerate(MONTHS): x = wsD.cell(r, 4 + j, float(z[m])); x.number_format = KT; x.font = BLUE
    wsD.cell(r, 16, f'=SUM(D{r}:O{r})').number_format = KT
    r += 1
DL = r - 1
for n_, c_ in [('D_O', 'A'), ('D_P', 'B'), ('D_Y', 'C'), ('D_T', 'P')]:
    wb.defined_names[n_] = DefinedName(n_, attr_text=f"Data!${c_}$5:${c_}${DL}")
wb.defined_names['D_M'] = DefinedName('D_M', attr_text=f"Data!$D$5:$O${DL}")
wsD.column_dimensions['B'].width = 22; wsD.freeze_panes = 'D5'
def ann_f(o, p, y): return f'SUMIFS(D_T,D_O,"{o}",D_P,{p},D_Y,{y})'

# ================= History (annual, formulas) =================
title(wsH, 'Annual exports by destination and origin, and the US share of each destination\'s US + Brazil purchases',
      'All values are SUMIFS on the Data tab. US share = US ÷ (US + Brazil).')
header(wsH, 4, ['Destination'] + [f'US {s}' for s in SEAS] + [f'Brazil {s}' for s in SEAS] + [f'US share {s}' for s in SEAS], h=48)
for i, p in enumerate(T.index.tolist() + ['China']):
    r = 5 + i; P = f'$A{r}'
    put(wsH, r, 1, p, f=font(10, True))
    for k, s in enumerate(SEAS):
        put(wsH, r, 2 + k, '=' + ann_f('US', P, f'"{s}"'), KT)
        put(wsH, r, 12 + k, '=' + ann_f('BR', P, f'"{s}"'), KT)
        put(wsH, r, 22 + k, f'=IF({L(2+k)}{r}+{L(12+k)}{r}=0,"",{L(2+k)}{r}/({L(2+k)}{r}+{L(12+k)}{r}))', P0)
HL = 5 + len(T)
r = HL + 1
put(wsH, r, 1, 'US total (all destinations)', f=font(10, True), fill=TOT)
put(wsH, r + 1, 1, 'US to China', f=font(10, True), fill=TOT)
put(wsH, r + 2, 1, 'US to rest of world', f=font(10, True), fill=TOT)
put(wsH, r + 3, 1, '   in mln bu', f=font(10, True), fill=TOT)
put(wsH, r + 4, 1, 'Brazil to all destinations except China', f=font(10, True), fill=TOT)
for k, s in enumerate(SEAS):
    tot = float(d[(d.O == 'US') & (d.CropYear == s)].Quantity.sum())
    put(wsH, r, 2 + k, tot, KT, f=BLUE, fill=TOT)
    put(wsH, r + 1, 2 + k, f'={L(2+k)}{HL}', KT, fill=TOT)
    put(wsH, r + 2, 2 + k, f'={L(2+k)}{r}-{L(2+k)}{r+1}', KT, f=font(10, True), fill=TOT)
    put(wsH, r + 3, 2 + k, f'={L(2+k)}{r+2}/{CONV}/1000', KT, fill=TOT)
    put(wsH, r + 4, 2 + k, '=' + ann_f('BR', '"ALL EX-CHINA"', f'"{s}"'), KT, fill=TOT)
note(wsH, r + 5, 'US totals are on USDA\'s official (Census) basis — they match the WASDE finals, e.g. 2024/25 = 1,892 and 2025/26 = 1,515 mln bu.')
wsH.column_dimensions['A'].width = 34
for c in range(2, 32): wsH.column_dimensions[L(c)].width = 9
wsH.freeze_panes = 'B5'
UST, USR, BRXR = r, r + 2, r + 4

# ================= Switching Ranking =================
title(wsS, f'Which destinations switch between US and Brazilian soybeans? — {SEAS[0]}–{SEAS[-1]} (10 seasons, monthly customs data)',
      'Ranked by average US volume over the last 5 seasons. Metrics computed by the build script from the Data tab; the risk scenario (yellow inputs) is live.')
put(wsS, 4, 1, 'Scenario: Brazil exports to all destinations except China this season (Mt)', f=font(10, True))
put(wsS, 4, 2, round(float(brx.loc[B5].mean()) + 3, 1), '0.0', f=BLUE, fill=YEL)
put(wsS, 5, 1, f'Brazil except China — 5-season average {B5[0]}–{B5[-1]} (Mt)', f=font(10, True))
put(wsS, 5, 2, f'=AVERAGE(History!{L(2+5)}{BRXR}:{L(2+9)}{BRXR})/1000', '0.0')
put(wsS, 6, 1, 'Extra Brazilian supply to non-China buyers vs normal (Mt)', f=font(10, True)); put(wsS, 6, 2, '=B4-B5', '+0.0;-0.0;0.0', f=font(10, True, NAVY))
put(wsS, 7, 1, 'Sensitivity to use (1 = detrended, recommended; 2 = simple)', f=font(10, True)); put(wsS, 7, 2, 1, '0', f=BLUE, fill=YEL)
put(wsS, 8, 1, 'Only count sensitivities that are statistically meaningful (|t| ≥ 1)? (Y/N)', f=font(10, True)); put(wsS, 8, 2, 'Y', f=BLUE, fill=YEL, al=Alignment(horizontal='center'))
put(wsS, 9, 1, 'Base-case US export number (mln bu, China at 25 Mt)', f=font(10, True)); put(wsS, 9, 2, BASE_MBU, KT, f=BLUE, fill=YEL)
put(wsS, 10, 1, 'USDA forecast (mln bu)', f=font(10, True)); put(wsS, 10, 2, 1720, KT, f=BLUE, fill=YEL)
dv = DataValidation(type='list', formula1='"Y,N"'); wsS.add_data_validation(dv); dv.add('B8')
wsS.cell(4, 3, f'10-season range {brx.min():.1f}–{brx.max():.1f} Mt (record {brx.idxmax()}). Default = 5-season average + 3 Mt.').font = font(9, False, '404040', True)
SH = ['Destination', 'Type', 'US avg (kt, 5 seasons)', 'Brazil avg (kt, 10 seasons)', 'US share of US+Brazil (avg)', 'US share — lowest season',
      'US share — highest season', 'US share Sep–Feb', 'US share Mar–Aug', 'Correlation US vs Brazil volumes',
      'Share of US volume swings due to switching', 'Sensitivity: US share pts per +1 Mt Brazil (simple)',
      'Sensitivity: detrended', 't-stat (detrended)', 'US+Brazil volume Nov–Aug (kt, 5-season avg)', 'US volume at risk in scenario (kt)', 'At risk (mln bu)',
      'US booked so far this season (kt, export sales)', 'US 5-yr avg full year (kt, export sales)', 'US still to book on usual pace (kt)',
      'At risk, capped at what is still to book (kt)', 'At risk, capped (mln bu)', 'Region']
R0 = 13
section(wsS, R0 - 1, '1) Ranking by destination')
header(wsS, R0, SH, h=64)
R1 = R0 + 1
for i, (p, x) in enumerate(T.iterrows()):
    r = R1 + i
    vals = [p, x['class'], x.US5, x.BR, x.share, x.smin, x.smax, x.sf, x.ma, x.rho, x.vfs, x.raw, x.detr, x.t, x.comb]
    fmts = [None, None, KT, KT, P0, P0, P0, P0, P0, '0.00', '0%', '0.00', '0.00', '0.0', KT]
    for j, (v, fm) in enumerate(zip(vals, fmts)):
        v = None if isinstance(v, float) and np.isnan(v) else v
        put(wsS, r, 1 + j, v, fm, f=font(10, j in (0, 1)) if j < 2 else BLUE)
    put(wsS, r, 16, f'=MIN(0,IF(OR($B$8="N",ABS(N{r})>=1),CHOOSE($B$7,M{r},L{r}),0)/100*$B$6*O{r})', '#,##0;-#,##0;0', f=font(10, True))
    put(wsS, r, 17, f'=P{r}/{CONV}/1000', '#,##0;-#,##0;0')
    en = esr_name(p)
    put(wsS, r, 18, float(ESR['now'].get(en, 0)) if ESR else None, KT, f=BLUE)
    put(wsS, r, 19, float(ESR['fy5'].get(en, 0)) if ESR else None, KT, f=BLUE)
    put(wsS, r, 20, f'=MAX(0,S{r}-R{r})', KT)
    put(wsS, r, 21, f'=IF(S{r}="",P{r},MAX(P{r},-T{r}))', '#,##0;-#,##0;0', f=font(10, True))
    put(wsS, r, 22, f'=U{r}/{CONV}/1000', '#,##0;-#,##0;0')
    put(wsS, r, 23, reg(p), f=font(9, False, '595959'))
R2 = R1 + len(T) - 1
RT = R2 + 1
put(wsS, RT, 1, 'TOTAL US volume at risk', f=font(10, True), fill=TOT)
for c in range(2, 24): put(wsS, RT, c, None, fill=TOT)
put(wsS, RT, 16, f'=SUM(P{R1}:P{R2})', '#,##0;-#,##0;0', f=font(11, True, NAVY), fill=TOT)
put(wsS, RT, 17, f'=SUM(Q{R1}:Q{R2})', '#,##0;-#,##0;0', f=font(11, True, NAVY), fill=TOT)
for c in (18, 19, 20): put(wsS, RT, c, f'=SUM({L(c)}{R1}:{L(c)}{R2})', KT, f=font(10, True), fill=TOT)
for c in (21, 22): put(wsS, RT, c, f'=SUM({L(c)}{R1}:{L(c)}{R2})', '#,##0;-#,##0;0', f=font(11, True, NAVY), fill=TOT)
wsS.conditional_formatting.add(f'B{R1}:B{R2}', FormulaRule(formula=[f'$B{R1}="Switcher"'], fill=RED))
wsS.conditional_formatting.add(f'B{R1}:B{R2}', FormulaRule(formula=[f'$B{R1}="Mixed"'], fill=AMB))
wsS.conditional_formatting.add(f'B{R1}:B{R2}', FormulaRule(formula=[f'$B{R1}="Demand-driven"'], fill=GRN))
wsS.conditional_formatting.add(f'P{R1}:P{R2}', CellIsRule(operator='lessThan', formula=['-50'], fill=RED))

s0 = RT + 2
section(wsS, s0, '2) What it means for the US export number')
res = [('Base case (China 25 Mt, rest of world at its usual pace)', '=B9', KT),
       ('Volume at risk from extra Brazilian competition (mln bu, capped)', f'=V{RT}', '+#,##0;-#,##0;0'),
       ('Adjusted US export number (mln bu)', f'=B{s0+1}+B{s0+2}', KT),
       ('vs USDA (mln bu)', f'=B{s0+3}-B10', '+#,##0;-#,##0;0')]
for i, (lb, f_, fm) in enumerate(res):
    put(wsS, s0 + 1 + i, 1, lb, f=font(10, True)); put(wsS, s0 + 1 + i, 2, f_, fm, f=font(12, True, NAVY))
header(wsS, s0 + 6, ['Extra Brazil supply to non-China buyers (Mt)', 'US volume at risk (mln bu)', 'Adjusted US number (mln bu)', 'vs USDA'], h=36)
for i, dlt in enumerate([0, 1, 2, 3, 4, 5, 6]):
    r = s0 + 7 + i
    put(wsS, r, 1, dlt, '+0;-0;0', f=BLUE)
    put(wsS, r, 2, f'=IF($B$6=0,0,SUMPRODUCT(-(-1)*($P${R1}:$P${R2}*A{r}/$B$6<-$T${R1}:$T${R2})*(-$T${R1}:$T${R2})+($P${R1}:$P${R2}*A{r}/$B$6>=-$T${R1}:$T${R2})*$P${R1}:$P${R2}*A{r}/$B$6)/{CONV}/1000)', '+#,##0;-#,##0;0')
    put(wsS, r, 3, f'=$B$9+B{r}', KT); put(wsS, r, 4, f'=C{r}-$B$10', '+#,##0;-#,##0;0', f=font(10, True))
wsS.conditional_formatting.add(f'D{s0+7}:D{s0+13}', CellIsRule(operator='lessThan', formula=['0'], fill=RED))
wsS.conditional_formatting.add(f'D{s0+7}:D{s0+13}', CellIsRule(operator='greaterThanOrEqual', formula=['0'], fill=GRN))
n0 = s0 + 15
for k, t in enumerate([
    'How to read the type: Switcher = US volume moves mainly because the buyer changes origin (US and Brazil volumes move in opposite directions). '
    'Demand-driven = US volume moves with the buyer\'s total imports (e.g. currency, crush margins), not because of Brazil. Mixed = both.',
    '"US share Mar–Aug" shows how much US business a buyer keeps in Brazil\'s season: low values (Spain, Turkey, UK, Portugal) mean it already switches every year — '
    'that normal switch is already in our pace model; the risk is switching EARLIER or MORE than usual.',
    'Sensitivity = change in the US share (percentage points) when Brazil ships 1 Mt more than usual to non-China buyers. "Detrended" removes the 10-year drift '
    '(Brazil\'s exports grew every year) so it measures year-to-year switching. With 10 seasons these estimates are noisy: country-level numbers are indicative, '
    'the total is more robust (simple and detrended totals are within ~10% of each other).',
    'This year\'s context: if China takes 25 Mt from the US it needs less from Brazil, so more Brazilian beans are available for everyone else — the scenario input.',
    'Not included: price spreads (US Gulf vs Paranaguá FOB). Brazil\'s non-China export volume is used as the proxy for how aggressively Brazil competes.']):
    note(wsS, n0 + k, t)
wsS.column_dimensions['A'].width = 62; wsS.column_dimensions['B'].width = 14
for c in range(3, 23): wsS.column_dimensions[L(c)].width = 11
wsS.column_dimensions['W'].width = 24
wsS.freeze_panes = f'C{R1}'
bc = BarChart(); bc.type = 'bar'; bc.title = 'US volume at risk by destination (kt)'; bc.y_axis.numFmt = '#,##0'
bc.add_data(Reference(wsS, min_col=21, min_row=R0, max_row=R2), titles_from_data=True)
bc.set_categories(Reference(wsS, min_col=1, min_row=R1, max_row=R2)); bc.legend = None
bc.x_axis.scaling.orientation = 'maxMin'; bc.x_axis.tickLblPos = 'low'; bc.width = 16; bc.height = 12
bc.x_axis.delete = False; bc.y_axis.delete = False
wsS.add_chart(bc, f'Y{R0}')

# ================= Country Profile =================
title(wsC, 'Destination profile — US share of US + Brazil purchases by month and season (pick a destination in B4)',
      'Live SUMIFS on the Data tab. Shows when in the year each buyer moves from US to Brazilian beans, and how that changes from season to season.')
put(wsC, 4, 1, 'Destination', f=font(10, True)); put(wsC, 4, 2, 'Netherlands', f=BLUE, fill=YEL)
dvc = DataValidation(type='list', formula1=f"='Switching Ranking'!$A${R1}:$A${R2}"); wsC.add_data_validation(dvc); dvc.add('B4')
put(wsC, 5, 1, 'Type', f=font(10, True)); put(wsC, 5, 2, f"=INDEX('Switching Ranking'!$B:$B,MATCH($B$4,'Switching Ranking'!$A:$A,0))")
header(wsC, 7, ['Season'] + MN + ['Full year'], h=24)
def msum(o, m_idx, y): return f'SUMPRODUCT((D_O="{o}")*(D_P=$B$4)*(D_Y={y})*INDEX(D_M,0,{m_idx}))'
for k, s in enumerate(SEAS):
    r = 8 + k
    put(wsC, r, 1, s, f=font(10, True))
    for j in range(12):
        us_, br_ = msum('US', j + 1, f'$A{r}'), msum('BR', j + 1, f'$A{r}')
        put(wsC, r, 2 + j, f'=IF(({us_})+({br_})=0,"",({us_})/(({us_})+({br_})))', P0)
    put(wsC, r, 14, f"=IFERROR(INDEX(History!$V:$AE,MATCH($B$4,History!$A:$A,0),{k+1}),\"\")", P0, f=font(10, True))
ra = 8 + len(SEAS)
put(wsC, ra, 1, 'Average (10 seasons)', f=font(10, True), fill=TOT)
put(wsC, ra + 1, 1, f'Average (5 seasons {B5[0]}–{B5[-1]})', f=font(10, True), fill=TOT)
for j in range(13):
    cl = L(2 + j)
    put(wsC, ra, 2 + j, f'=IFERROR(AVERAGE({cl}8:{cl}{ra-1}),"")', P0, f=font(10, True), fill=TOT)
    put(wsC, ra + 1, 2 + j, f'=IFERROR(AVERAGE({cl}{ra-5}:{cl}{ra-1}),"")', P0, f=font(10, True), fill=TOT)
for rg in (f'B8:N{ra+1}',):
    from openpyxl.formatting.rule import ColorScaleRule
    wsC.conditional_formatting.add(rg, ColorScaleRule(start_type='num', start_value=0, start_color='F8696B', mid_type='num', mid_value=0.5, mid_color='FFEB84', end_type='num', end_value=1, end_color='63BE7B'))
header(wsC, ra + 3, ['Season'] + [f'US {m}' for m in MN] + ['US total'], h=24)
for k, s in enumerate(SEAS):
    r = ra + 4 + k
    put(wsC, r, 1, s)
    for j in range(12): put(wsC, r, 2 + j, '=' + msum('US', j + 1, f'$A{r}'), KT)
    put(wsC, r, 14, f'=SUM(B{r}:M{r})', KT, f=font(10, True))
rb = ra + 4 + len(SEAS) + 1
header(wsC, rb, ['Season'] + [f'Brazil {m}' for m in MN] + ['Brazil total'], h=24)
for k, s in enumerate(SEAS):
    r = rb + 1 + k
    put(wsC, r, 1, s)
    for j in range(12): put(wsC, r, 2 + j, '=' + msum('BR', j + 1, f'$A{r}'), KT)
    put(wsC, r, 14, f'=SUM(B{r}:M{r})', KT, f=font(10, True))
wsC.column_dimensions['A'].width = 24
for c in range(2, 15): wsC.column_dimensions[L(c)].width = 9
lc = LineChart(); lc.title = 'US share of US + Brazil purchases by month'; lc.y_axis.numFmt = '0%'
for rr in (ra, ra + 1, ra - 1):
    from openpyxl.chart import Series
    lc.series.append(Series(Reference(wsC, min_col=2, max_col=13, min_row=rr), title={ra: '10-season avg', ra + 1: '5-season avg', ra - 1: f'{SEAS[-1]}'}[rr]))
lc.set_categories(Reference(wsC, min_col=2, max_col=13, min_row=7)); lc.y_axis.scaling.min = 0; lc.y_axis.scaling.max = 1
lc.width = 18; lc.height = 9; lc.x_axis.delete = False; lc.y_axis.delete = False
for s_ in lc.series: s_.smooth = False
wsC.add_chart(lc, 'P4')

# ================= By Region =================
wsR = wb.create_sheet('By Region', index=1)
title(wsR, 'By region — how much US business Brazil can take, and the Sep–Feb vs Mar–Aug split of US and Brazilian exports',
      'Table A is live (follows the scenario on "Switching Ranking"). Tables B–C are computed from the monthly customs data (kt / Mt). Regions as in the pace model.')
section(wsR, 4, 'A) US volume Brazil could take this season, by region (kt)')
AH = ['Region', 'US still to book on usual pace (kt)', 'At risk — scenario on Ranking tab (kt)', 'At risk if Brazil +3 Mt (kt)', 'At risk if Brazil +5 Mt (kt)',
      'At risk if Brazil +7.2 Mt (provider forecast) (kt)', 'Share of what is still to book at risk (+7.2 Mt)', 'Main switchers']
header(wsR, 5, AH, h=52)
RNK = "'Switching Ranking'"
def risk_static(D):
    out = {}
    for p_, x_ in T.iterrows():
        still = float(max(0, ESR['fy5'].get(esr_name(p_), 0) - ESR['now'].get(esr_name(p_), 0))) if ESR else 1e9
        v = min(0.0, x_.detr / 100 * D * x_.comb) if abs(x_.t) >= 1 else 0.0
        out[p_] = max(v, -still)
    return out
RS = {D: risk_static(D) for D in (3, 5, 7.2)}
regs_sw = [g for g in REGS if g != 'China' and any(reg(p_) == g for p_ in T.index)]
for i, g in enumerate(regs_sw):
    r = 6 + i
    put(wsR, r, 1, g, f=font(10, True))
    put(wsR, r, 2, f"=SUMIF({RNK}!$W${R1}:$W${R2},$A{r},{RNK}!$T${R1}:$T${R2})", KT)
    put(wsR, r, 3, f"=SUMIF({RNK}!$W${R1}:$W${R2},$A{r},{RNK}!$U${R1}:$U${R2})", '#,##0;-#,##0;0', f=font(10, True))
    for j, D in enumerate((3, 5, 7.2)):
        put(wsR, r, 4 + j, sum(v for p_, v in RS[D].items() if reg(p_) == g), '#,##0;-#,##0;0', f=BLUE)
    put(wsR, r, 7, f'=IF(B{r}=0,"",-F{r}/B{r})', P0)
    sw_ = sorted([(v, p_) for p_, v in RS[7.2].items() if reg(p_) == g and v < -20])
    put(wsR, r, 8, ', '.join(p_ for _, p_ in sw_), f=font(9))
ra_ = 6 + len(regs_sw)
put(wsR, ra_, 1, 'TOTAL (destinations ≥150 kt/yr US)', f=font(10, True), fill=TOT)
for c in range(2, 7): put(wsR, ra_, c, f'=SUM({L(c)}6:{L(c)}{ra_-1})', '#,##0;-#,##0;0', f=font(10, True), fill=TOT)
put(wsR, ra_, 7, f'=IF(B{ra_}=0,"",-F{ra_}/B{ra_})', P0, f=font(10, True), fill=TOT); put(wsR, ra_, 8, None, fill=TOT)
put(wsR, ra_ + 1, 1, '   in mln bu', f=font(10, True))
for c in range(3, 7): put(wsR, ra_ + 1, c, f'={L(c)}{ra_}/{CONV}/1000', '#,##0;-#,##0;0', f=font(10, True, NAVY))
note(wsR, ra_ + 2, '"Still to book" = 5-yr average full-year US sales minus what is already booked (weekly export sales). A destination cannot switch more than that. '
                   'China is excluded (fixed at 25 Mt). Brazil +7.2 Mt = provider forecast of 33.5 Mt to non-China buyers vs 26.3 Mt 5-yr average.')

b0 = ra_ + 5
section(wsR, b0, f'B) Who ships when — 5-season average {B5[0]}–{B5[-1]} (Mt; customs basis; Sep–Feb = US season, Mar–Aug = Brazil season)')
BH = ['Region', 'US Sep–Feb', 'US Mar–Aug', 'Brazil Sep–Feb', 'Brazil Mar–Aug', 'US share Sep–Feb', 'US share Mar–Aug', 'US share full year', 'US volume in Mar–Aug as % of its year']
header(wsR, b0 + 1, BH, h=44)
hh = d[~d.Partner_Name.isin(['United States', 'Brazil'])].copy()
hh['R'] = hh.Partner_Name.map(reg); hh['H'] = np.where(hh.M.isin([9, 10, 11, 12, 1, 2]), 'H1', 'H2')
def split(years):
    x = hh[hh.CropYear.isin(years)].pivot_table(index='R', columns=['O', 'H'], values='Quantity', aggfunc='sum').fillna(0) / len(years) / 1000
    return x.reindex(REGS).fillna(0)
def write_split(r0, x, label_total):
    for i, g in enumerate(REGS):
        r = r0 + i
        put(wsR, r, 1, g, f=font(10, g == 'China'))
        for j, k in enumerate([('US', 'H1'), ('US', 'H2'), ('BR', 'H1'), ('BR', 'H2')]):
            put(wsR, r, 2 + j, float(x.loc[g, k]) if k in x.columns else 0.0, '0.0', f=BLUE)
        put(wsR, r, 6, f'=IF(B{r}+D{r}=0,"",B{r}/(B{r}+D{r}))', P0); put(wsR, r, 7, f'=IF(C{r}+E{r}=0,"",C{r}/(C{r}+E{r}))', P0)
        put(wsR, r, 8, f'=IF(SUM(B{r}:E{r})=0,"",(B{r}+C{r})/SUM(B{r}:E{r}))', P0); put(wsR, r, 9, f'=IF(B{r}+C{r}=0,"",C{r}/(B{r}+C{r}))', P0)
    rt = r0 + len(REGS)
    for k_, (lab, rows_) in enumerate([(label_total, f'{r0}:{rt-1}'), ('Rest of world (excl. China)', None)]):
        r = rt + k_
        put(wsR, r, 1, lab, f=font(10, True), fill=TOT)
        for c in range(2, 6):
            cl = L(c)
            f_ = f'=SUM({cl}{r0}:{cl}{rt-1})' if rows_ else f'={cl}{rt}-{cl}{r0}'
            put(wsR, r, c, f_, '0.0', f=font(10, True), fill=TOT)
        put(wsR, r, 6, f'=B{r}/(B{r}+D{r})', P0, f=font(10, True), fill=TOT); put(wsR, r, 7, f'=C{r}/(C{r}+E{r})', P0, f=font(10, True), fill=TOT)
        put(wsR, r, 8, f'=(B{r}+C{r})/SUM(B{r}:E{r})', P0, f=font(10, True), fill=TOT); put(wsR, r, 9, f'=C{r}/(B{r}+C{r})', P0, f=font(10, True), fill=TOT)
    return rt
bt = write_split(b0 + 2, split(B5), 'TOTAL (all destinations)')
c0_ = bt + 4
section(wsR, c0_, f'C) Same split for the last season {SEAS[-1]} (Mt) — China bought almost no US beans')
header(wsR, c0_ + 1, BH, h=44)
ct = write_split(c0_ + 2, split([SEAS[-1]]), 'TOTAL (all destinations)')
note(wsR, ct + 3, 'Reading: the US ships about three-quarters of its year in Sep–Feb; Brazil ships about three-quarters of its year in Mar–Aug. '
                  'Outside China the US holds ~3/4 of the Sep–Feb market but only ~1/3 of Mar–Aug — that second-half share is what extra Brazilian supply attacks.')
wsR.column_dimensions['A'].width = 34
for c in range(2, 8): wsR.column_dimensions[L(c)].width = 14
wsR.column_dimensions['H'].width = 40; wsR.column_dimensions['I'].width = 14
from openpyxl.chart import BarChart as _BC, Series as _S
ch_ = _BC(); ch_.type = 'col'; ch_.grouping = 'stacked'; ch_.overlap = 100; ch_.title = 'Rest of world (excl. China): who ships when (Mt, 5-season avg)'
rr_ = bt + 1
for c, t_ in ((2, 'US Sep–Feb'), (3, 'US Mar–Aug'), (4, 'Brazil Sep–Feb'), (5, 'Brazil Mar–Aug')):
    ch_.series.append(_S(Reference(wsR, min_col=c, min_row=rr_), title=t_))
ch_.width = 14; ch_.height = 8; ch_.x_axis.delete = True; ch_.y_axis.delete = False
wsR.add_chart(ch_, f'K{b0}')

# ================= Method =================
title(wsN, 'Method', 'How the switching metrics and the risk scenario are calculated.')
steps = [
    ('Data', f'Monthly customs exports of soybeans by destination for the US and Brazil, crop years {SEAS[0]}–{SEAS[-1]} (Sep–Aug). Destinations with ≥150 kt/yr of US business.'),
    ('US share', 'US ÷ (US + Brazil) for each destination, by season and by half (Sep–Feb = US season, Mar–Aug = Brazil season).'),
    ('Switching vs demand', 'US volume = total (US + Brazil) × US share. The share of the year-to-year variance of US volume that comes from the share (not from the total) '
                            'tells whether a buyer switches origin (high) or just buys more or less overall (low). Combined with the US–Brazil volume correlation (negative = substitution).'),
    ('Sensitivity', 'Regression across the 10 seasons of the US share on Brazil\'s exports to non-China buyers (Mt), with and without a time trend. '
                    'Result = percentage points of US share lost per extra 1 Mt of Brazilian supply to non-China buyers.'),
    ('Volume at risk', 'Sensitivity × extra Brazilian supply this season × the destination\'s usual US + Brazil volume from November to August (the part of the year still to come). '
                       'Optionally only sensitivities with |t| ≥ 1.'),
    ('Limits', 'Ten seasons; no price data (FOB spreads) — Brazil\'s non-China export volume is a proxy for competition. Country numbers are indicative; the total is more stable.'),
]
for i, (k, v) in enumerate(steps):
    r = 4 + i
    a = wsN.cell(r, 1, k); a.font = font(11, True, NAVY); a.alignment = Alignment(vertical='top')
    b = wsN.cell(r, 2, v); b.font = font(11); b.alignment = Alignment(wrap_text=True, vertical='top'); wsN.row_dimensions[r].height = 46
wsN.column_dimensions['A'].width = 24; wsN.column_dimensions['B'].width = 120

for ws in wb: ws.sheet_view.showGridLines = False
wb.save(out)
print('seasons', SEAS[0], SEAS[-1], 'destinations', len(T))
