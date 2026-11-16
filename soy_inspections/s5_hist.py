def week1(y):
    d=dt.date(y,9,1)
    while d.weekday()!=3: d+=dt.timedelta(1)
    return d
def hload(path):
    d=pd.read_csv(path,dtype=str,usecols=lambda c:c.strip() in ('Thursday','Grain','Destination','Metric Ton'))
    d.columns=[c.strip() for c in d.columns]; d=d[d.Grain=='SOYBEANS'].copy(); d['MT']=pd.to_numeric(d['Metric Ton']); return d
NW=len(weeks); YEARS=list(range(2015,2027))
HG=['Total','China','Rest of World excl. China']+[x for x in regions if x!='China']
def ydata(d,y):
    w1=week1(y); ws_=[(w1+dt.timedelta(7*k)).strftime('%Y%m%d') for k in range(NW)]
    out={g:[0]*NW for g in HG}; unm=set()
    for k,w in enumerate(ws_):
        z=d[d.Thursday==w]
        for dest,mt in z.groupby('Destination').MT.sum().items():
            mt=int(mt); out['Total'][k]+=mt
            if dest=='CHINA': out['China'][k]+=mt; continue
            out['Rest of World excl. China'][k]+=mt
            r_=REG.get(dest)
            if r_ is None: unm.add(dest)
            elif r_ in out: out[r_][k]+=mt
    return out,ws_,unm
HD={}; HW={}
for y in YEARS:
    d=x25 if y==2025 else (x26 if y==2026 else hload(HIST[y]))
    HD[y],HW[y],u=ydata(d,y)
    if u: print('UNMAPPED',y,u)
    if y<2026: assert all(w in set(d.Thursday) for w in HW[y]), (y,'missing weeks')
assert HW[2026]==[b for _,a,b in weeks] and HW[2025]==[a for _,a,b in weeks]
hd=wb.create_sheet('History Data')
hd['A1']='Soybean inspections — weekly MT by group, season weeks 1–%d, 2015–2026 (inputs)'%NW; hd['A1'].font=Font(name='Calibri',bold=True,size=12)
hd['A2']=('Week 1 = first Thursday on/after 1 Sep (soybean marketing year start); weeks are FGIS "Thursday" week-ending dates. Source: FGIS CY files — '+DATA_NOTE+
          ' Grain = SOYBEANS. Regions as on Country by Country tab. Cumulative columns are formulas.')
hd['A2'].font=Font(name='Calibri',italic=True,size=9)
hh=['Group','Year']+[f'Wk{k+1}' for k in range(NW)]+[f'Cum wk{k+1}' for k in range(NW)]
for c,h in enumerate(hh):
    x=hd.cell(4,1+c,h); x.font=HF; x.fill=HFILL; x.alignment=Alignment(horizontal='center')
r=5; HD0=r
for g in HG:
    for y in YEARS:
        hd.cell(r,1,g); hd.cell(r,2,y)
        for k in range(NW):
            x=hd.cell(r,3+k,HD[y][g][k]); x.number_format=NUM; x.font=Font(name='Calibri',color='FF0000FF')
            cc=3+NW+k; hd.cell(r,cc,f'={L(3)}{r}' if k==0 else f'={L(cc-1)}{r}+{L(3+k)}{r}').number_format=NUM
        r+=1
HD1=r-1
r+=1; hd.cell(r,1,'Week-ending dates by year').font=B; r+=1
hd.cell(r,1,'Year').font=B
for k in range(NW): hd.cell(r,3+k,f'Wk{k+1}').font=B
for y in YEARS:
    r+=1; hd.cell(r,1,y)
    for k in range(NW): x=hd.cell(r,3+k,dt.datetime.strptime(HW[y][k],'%Y%m%d').date()); x.number_format='dd-mmm-yy'
hd.column_dimensions['A'].width=26
for c in range(3,3+2*NW): hd.column_dimensions[L(c)].width=11
hd.freeze_panes='C5'
sp=wb.create_sheet('Seasonal Pace',4)
sp['A1']='Soybean inspections — 2026 seasonal pace vs history (5-yr and 10-yr averages)'; sp['A1'].font=Font(name='Calibri',bold=True,size=12)
sp['A2']=('Cumulative inspections from week 1 (first Thursday on/after 1 Sep). Benchmarks (editable, row 4): simple 5-yr and 10-yr averages of the same season weeks, no years excluded. '
          'Min/max = range over the 10-yr window. Pace % = 2026 cumulative / benchmark − 1. Note: both windows include trade-war years (2018–19, 2025) when China was largely absent and RoW took more US beans.')
sp['A2'].font=Font(name='Calibri',italic=True,size=9)
for c,v in [('A4','5-yr from'),('C4',2021),('D4','to'),('E4',2025),('F4','10-yr from'),('H4',2016),('I4','to'),('J4',2025)]:
    sp[c]=v
    if isinstance(v,int): sp[c].font=Font(name='Calibri',color='FF0000FF',bold=True); sp[c].fill=PatternFill('solid',fgColor='FFFFFF00')
    else: sp[c].font=Font(name='Calibri',size=9,italic=True)
HDs="'History Data'"; GA=f'{HDs}!$A${HD0}:$A${HD1}'; YA=f'{HDs}!$B${HD0}:$B${HD1}'
def cumcol(k): c=L(3+NW+k); return f'{HDs}!${c}${HD0}:${c}${HD1}'
def wkcol(k): c=L(3+k); return f'{HDs}!${c}${HD0}:${c}${HD1}'
def f_year(g,y,k): return f'SUMIFS({cumcol(k)},{GA},"{g}",{YA},{y})'
def f_avg(g,k,a,b): return f'AVERAGEIFS({cumcol(k)},{GA},"{g}",{YA},">="&{a},{YA},"<="&{b})'
def f_mm(fn,g,k): return f'_xlfn.{fn}({cumcol(k)},{GA},"{g}",{YA},">="&$H$4,{YA},"<="&$J$4)'
r=6; sp.cell(r,1,f'A. Season-to-date inspections, weeks 1–{NW} (to ~{dt.datetime.strptime(HW[2026][-1],"%Y%m%d").strftime("%-d %b")}), MT — by year').font=Font(name='Calibri',bold=True,size=11)
r+=1; hdrA=['Group / region','']+[str(y) for y in YEARS]+['5-yr avg','10-yr avg','2026 vs 5-yr','2026 vs 10-yr']
for c,h in enumerate(hdrA):
    x=sp.cell(r,1+c,h); x.font=HF; x.fill=HFILL; x.alignment=Alignment(horizontal='center',wrap_text=True)
sp.row_dimensions[r].height=30
A0=r+1; K=NW-1; c5=3+len(YEARS); c26=L(3+YEARS.index(2026))
for i,g in enumerate(HG):
    rr=A0+i; sp.cell(rr,1,g).font=Font(name='Calibri',bold=g in ('Total','China','Rest of World excl. China'))
    if g not in ('Total','China','Rest of World excl. China'): sp.cell(rr,1).alignment=Alignment(indent=2); sp.cell(rr,1).font=Font(name='Calibri',color='FF595959')
    for j,y in enumerate(YEARS): sp.cell(rr,3+j,'='+f_year(g,y,K)).number_format=NUM
    sp.cell(rr,c5,'='+f_avg(g,K,'$C$4','$E$4')).number_format=NUM
    sp.cell(rr,c5+1,'='+f_avg(g,K,'$H$4','$J$4')).number_format=NUM
    for q in range(2):
        bc_=L(c5+q); sp.cell(rr,c5+2+q,f'=IF({bc_}{rr}=0,"",{c26}{rr}/{bc_}{rr}-1)').number_format=PCT
    for c in range(3,c5+4): sp.cell(rr,c).font=Font(name='Calibri',bold=g=='Total')
A1=A0+len(HG)-1
cf(sp,f'{L(c5+2)}{A0}:{L(c5+3)}{A1}')
for rr in range(A0-1,A1+1):
    sp.cell(rr,3+YEARS.index(2026)).fill=PatternFill('solid',fgColor='FFDDEBF7') if rr>=A0 else HFILL
sp.cell(A1+1,1,'Trade-war years: 2018–19 (China largely absent) and 2025 (China zero). Both the 5-yr and 10-yr averages include some of them, which inflates the RoW benchmark and deflates China/Total.').font=Font(name='Calibri',italic=True,size=9)
r=A1+4; sp.cell(r,1,'B. Cumulative pace by week — 2026 vs benchmarks').font=Font(name='Calibri',bold=True,size=11); r+=1
hdrB=['Group / series','']+[dt.datetime.strptime(b,'%Y%m%d').strftime('%-d %b') for b in HW[2026]]
BB={}
SER=['2026','2025','5-yr avg','10-yr avg','10-yr min','10-yr max','2026 vs 5-yr','2026 vs 10-yr']
for g in HG:
    if g=='Sub-Saharan Africa': continue
    for c,h in enumerate(hdrB[1:],start=2):
        x=sp.cell(r,c,h if c>2 else ''); x.font=HF; x.fill=HFILL; x.alignment=Alignment(horizontal='center')
    sp.cell(r,1,g); sp.cell(r,1).fill=HFILL; sp.cell(r,1).font=HF
    BB[g]=r+1
    for s_i,sname in enumerate(SER):
        rr=r+1+s_i; sp.cell(rr,1,sname).alignment=Alignment(indent=1)
        for k in range(NW):
            col=3+k; cl=L(col)
            if sname=='2026': v='='+f_year(g,2026,k)
            elif sname=='2025': v='='+f_year(g,2025,k)
            elif sname=='5-yr avg': v='='+f_avg(g,k,'$C$4','$E$4')
            elif sname=='10-yr avg': v='='+f_avg(g,k,'$H$4','$J$4')
            elif sname=='10-yr min': v='='+f_mm('MINIFS',g,k)
            elif sname=='10-yr max': v='='+f_mm('MAXIFS',g,k)
            else:
                ref={'2026 vs 5-yr':r+3,'2026 vs 10-yr':r+4}[sname]
                v=f'=IF({cl}{ref}<$F$5,"",{cl}{r+1}/{cl}{ref}-1)'
            x=sp.cell(rr,col,v); x.number_format=PCT if sname.startswith('2026 vs') else NUM
            x.font=Font(name='Calibri',bold=sname=='2026')
        if sname.startswith('2026 vs'): sp.cell(rr,1).font=Font(name='Calibri',bold=True)
    sp.conditional_formatting.add(f'C{r+7}:{L(2+NW)}{r+8}',ColorScaleRule(start_type='num',start_value=-1,start_color='E34948',mid_type='num',mid_value=0,mid_color='F0EFEC',end_type='num',end_value=1,end_color='2A78D6'))
    r+=len(SER)+2
sp['F5']=50000; sp['A5']='Min. benchmark (MT) to show a pace %:'; sp['A5'].font=Font(name='Calibri',size=9,italic=True)
sp['F5'].font=Font(name='Calibri',color='FF0000FF',bold=True); sp['F5'].fill=PatternFill('solid',fgColor='FFFFFF00'); sp['F5'].number_format=NUM
sp.column_dimensions['A'].width=28; sp.column_dimensions['B'].width=2
for c in range(3,3+len(YEARS)+6): sp.column_dimensions[L(c)].width=11.5
sp.freeze_panes='C6'
ch0=3+len(YEARS)+7
hb=BarChart(); hb.type='col'; hb.grouping='stacked'; hb.overlap=100; hb.gapWidth=50
for g,col in (('Rest of World excl. China','2A78D6'),('China','EB6834')):
    rr=A0+HG.index(g)
    hb.add_data(Reference(sp,min_col=3,max_col=2+len(YEARS),min_row=rr),from_rows=True,titles_from_data=False)
    s_=hb.series[-1]; s_.tx=SeriesLabel(v=g); s_.graphicalProperties.solidFill=col; s_.graphicalProperties.line.solidFill='FFFFFF'
hb.set_categories(Reference(sp,min_col=3,max_col=2+len(YEARS),min_row=A0-1))
hb.title=f'Season-to-date inspections (weeks 1–{NW}) by year: China vs RoW'; hb.y_axis.numFmt='#,##0,,"M"'; hb.y_axis.title='MT'
hb.y_axis.majorGridlines.spPr=GraphicalProperties(ln=LineProperties(solidFill='E5E4E0')); hb.legend.position='b'
hb.x_axis.delete=False; hb.y_axis.delete=False; hb.width=22; hb.height=10
sp.add_chart(hb,f'{L(ch0)}6')
pb=BarChart(); pb.type='bar'; pb.title='2026 season-to-date vs benchmarks'
grp=[g for g in HG if g not in ('China','Sub-Saharan Africa')]
hr=A1+2
sp.cell(hr,ch0+14,'chart helper').font=Font(name='Calibri',size=8,color='FF999999')
for i,g in enumerate(grp):
    rr=A0+HG.index(g); sp.cell(hr+1+i,ch0+14,g)
    for q in range(2): sp.cell(hr+1+i,ch0+15+q,f'={L(c5+2+q)}{rr}').number_format=PCT
for q,(lab_,col) in enumerate((('vs 5-yr avg','2A78D6'),('vs 10-yr avg','EB6834'))):
    pb.add_data(Reference(sp,min_col=ch0+15+q,min_row=hr+1,max_row=hr+len(grp)),titles_from_data=False)
    s_=pb.series[-1]; s_.tx=SeriesLabel(v=lab_); s_.graphicalProperties.solidFill=col; s_.graphicalProperties.line.solidFill='FFFFFF'
pb.set_categories(Reference(sp,min_col=ch0+14,min_row=hr+1,max_row=hr+len(grp)))
pb.x_axis.scaling.orientation='maxMin'; pb.x_axis.tickLblPos='low'; pb.y_axis.numFmt='0%'; pb.legend.position='b'
pb.y_axis.majorGridlines.spPr=GraphicalProperties(ln=LineProperties(solidFill='E5E4E0'))
pb.x_axis.delete=False; pb.y_axis.delete=False; pb.width=22; pb.height=10
sp.add_chart(pb,f'{L(ch0)}27')
sm=[g for g in HG if g!='Sub-Saharan Africa']
for n_,g in enumerate(sm):
    b0=BB[g]; c=LineChart(); c.title=g
    for s_i,col,wd,dash in ((4,'C9C8C2',12700,None),(5,'C9C8C2',12700,None),(3,'EB6834',19050,'dash'),(2,'1BAF7A',19050,'dash'),(1,'8C8B86',19050,'sysDot'),(0,'2A78D6',31750,None)):
        nm={0:'2026',1:'2025',2:'5-yr avg',3:'10-yr avg',4:'10-yr min',5:'10-yr max'}[s_i]
        c.add_data(Reference(sp,min_col=3,max_col=2+NW,min_row=b0+s_i),from_rows=True,titles_from_data=False)
        s_=c.series[-1]; s_.tx=SeriesLabel(v=nm); s_.smooth=False; s_.marker.symbol='none'
        s_.graphicalProperties.line.solidFill=col; s_.graphicalProperties.line.width=wd
        if dash: s_.graphicalProperties.line.dashStyle=dash
    c.set_categories(Reference(sp,min_col=3,max_col=2+NW,min_row=BB[sm[0]]-1))
    c.y_axis.numFmt='#,##0,"k"'; c.legend.position='b'
    c.y_axis.majorGridlines.spPr=GraphicalProperties(ln=LineProperties(solidFill='E5E4E0'))
    c.x_axis.delete=False; c.y_axis.delete=False; c.width=13; c.height=8
    sp.add_chart(c,f'{L(ch0+(n_%2)*9)}{48+(n_//2)*17}')
sp.cell(46,ch0,'Cumulative MT by group: 2026 solid blue; 2025 grey dotted; 5-yr avg green dashed; 10-yr avg orange dashed; 10-yr min/max thin grey. Own scale per panel.').font=Font(name='Calibri',italic=True,size=9)
