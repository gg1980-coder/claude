import pandas as pd, datetime as dt
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.formatting.rule import CellIsRule, FormulaRule, ColorScaleRule
from openpyxl.utils import get_column_letter as L
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.chart.series import SeriesLabel
from openpyxl.chart.shapes import GraphicalProperties
from openpyxl.drawing.line import LineProperties
def load(f):
    d=pd.read_csv(f,dtype=str); d.columns=[c.strip() for c in d.columns]
    d['MT']=pd.to_numeric(d['Metric Ton']); return d
d25=load(CY2025); d26=load(CY2026)
x25=d25[d25.Grain=='SOYBEANS']; x26=d26[d26.Grain=='SOYBEANS']
def agg(d,wks,grain='SOYBEANS'):
    x=d[(d.Grain==grain)&d.Thursday.isin(wks)]
    return x.groupby('Destination').agg(MT=('MT','sum'),N=('MT','size'))
for _,_,w25,w26,_ in PERIODS:
    assert all(w in set(d25.Thursday) for w in w25), w25
    assert all(w in set(d26.Thursday) for w in w26), w26
NUM='#,##0;\\-#,##0;\\-'; PCT='\\+0.0%;\\-0.0%;0.0%'
HF=Font(name='Calibri',bold=True,color='FFFFFFFF'); HFILL=PatternFill('solid',fgColor='FF1F4E78')
TFILL=PatternFill('solid',fgColor='FFFFF2CC'); B=Font(name='Calibri',bold=True)
GREEN=Font(name='Calibri',color='FF007A33'); RED=Font(name='Calibri',color='FFC00000')
def cf(ws,rng):
    first=rng.split(':')[0]
    ws.conditional_formatting.add(rng,FormulaRule(formula=[f'AND(ISNUMBER({first}),{first}>0)'],font=GREEN))
    ws.conditional_formatting.add(rng,FormulaRule(formula=[f'AND(ISNUMBER({first}),{first}<0)'],font=RED))
periods=PERIODS
weeks=[]
for name,short,w25,w26,_ in periods:
    for a,b in zip(w25,w26): weeks.append((name,a,b))
wb=Workbook(); summ=wb.active; summ.title='Summary China vs RoW'
refs={}
for name,short,w25,w26,desc in periods:
    ws=wb.create_sheet(name)
    ws['A1']=f'US Soybean Export Inspections — {desc} (weekly-report basis)'; ws['A1'].font=Font(name='Calibri',bold=True,size=12)
    ws['A2']=('Source: USDA AMS/FGIS export grain inspections, CY2025 and CY2026 files (Grain = SOYBEANS, grouped by "Thursday" week-ending date). '
              f'CY2026 data as of file dated wk ending {dt.datetime.strptime(weeks[-1][2],"%Y%m%d").strftime("%-d %b %Y")}, incl. revisions.')
    ws['A2'].font=Font(name='Calibri',italic=True,size=9)
    ws.merge_cells('A4:A5'); ws.merge_cells('B4:E4'); ws.merge_cells('F4:H4')
    ws['A4']='Destination'; ws['B4']='Metric tons'; ws['F4']='Number of inspections'
    for c,v in zip('BCDEFGH',['2025','2026','Δ MT','Δ %','2025','2026','Δ']): ws[f'{c}5']=v
    for r in (4,5):
        for c in 'ABCDEFGH': ws[f'{c}{r}'].font=HF; ws[f'{c}{r}'].fill=HFILL; ws[f'{c}{r}'].alignment=Alignment(horizontal='center')
    a=agg(d25,w25); b=agg(d26,w26)
    t=a.join(b,lsuffix='25',rsuffix='26',how='outer').fillna(0).astype(int)
    t['ch']=t.index!='CHINA'; t=t.sort_values(['ch','MT26','MT25'],ascending=[True,False,False])
    r=6
    for dest,row in t.iterrows():
        ws[f'A{r}']=dest; ws[f'B{r}']=row.MT25; ws[f'C{r}']=row.MT26; ws[f'F{r}']=row.N25; ws[f'G{r}']=row.N26; r+=1
    last=r-1; tot=r; chn=r+1; row_=r+2
    ws[f'A{tot}']='TOTAL'; ws[f'A{chn}']='China'; ws[f'A{row_}']='Rest of World excl. China'
    for c in 'BCFG':
        ws[f'{c}{tot}']=f'=SUM({c}6:{c}{last})'
        ws[f'{c}{chn}']=f'=SUMIF($A$6:$A${last},"CHINA",{c}6:{c}{last})'
        ws[f'{c}{row_}']=f'={c}{tot}-{c}{chn}'
    for i in range(6,row_+1):
        ws[f'D{i}']=f'=C{i}-B{i}'; ws[f'H{i}']=f'=G{i}-F{i}'
        ws[f'E{i}']=f'=IF(B{i}=0,IF(C{i}=0,0,"new"),C{i}/B{i}-1)'
        for c in 'BCDFGH': ws[f'{c}{i}'].number_format=NUM
        ws[f'E{i}'].number_format=PCT; ws[f'E{i}'].alignment=Alignment(horizontal='right')
        for c in 'ABCDEFGH':
            ws[f'{c}{i}'].font=Font(name='Calibri',bold=i>=tot)
            if i==tot: ws[f'{c}{i}'].fill=TFILL
    cf(ws,f'D6:E{row_}'); cf(ws,f'H6:H{row_}')
    ws.column_dimensions['A'].width=26
    for c in 'BCDEFGH': ws.column_dimensions[c].width=13
    ws.freeze_panes='B6'
    refs[name]=(chn,row_,tot)
