from openpyxl.worksheet.datavalidation import DataValidation
ps=wb.create_sheet('Region Progression',3)
GR=['Total','China','Rest of World excl. China']+[x for x in regions if x!='China']
CHARTR=[g for g in GR if g!='Sub-Saharan Africa']
W=len(weeks); C0=3
PP='+0.0" pp";-0.0" pp";0.0" pp"'
SPs="'Seasonal Pace'"
def normw(g,j):
    cond=f'({GA}="{g}")*({YA}>=$E$4)*({YA}<=$G$4)'
    return f'SUMPRODUCT({cond}*{wkcol(j)})/SUMPRODUCT({cond})'
def wkdata(x,w):
    y=x[x.Thursday==w]; out={g:[0,0] for g in GR}
    for dest,grp in y.groupby('Destination'):
        mt=int(grp.MT.sum()); n=len(grp); reg=REG.get(dest,'Other')
        for g in ('Total',('China' if dest=='CHINA' else 'Rest of World excl. China'), reg if dest!='CHINA' else None):
            if g and g in out: out[g][0]+=mt; out[g][1]+=n
    return out
D25=[wkdata(x25,a) for _,a,b in weeks]; D26=[wkdata(x26,b) for _,a,b in weeks]
ps['A1']='Soybean inspections by region — pace vs 5-yr / 10-yr average (selector in C4) and vs 2025'; ps['A1'].font=Font(name='Calibri',bold=True,size=12)
ps['A2']=('Benchmark = simple average of the same season week (week 1 = first Thursday on/after 1 Sep) over the window chosen in C4 — 5-yr or 10-yr, as defined on the Seasonal Pace tab (row 4); no years excluded. '
          'pp = percentage points. A % is blank when the benchmark is below the row-3 thresholds. Regions as on Country by Country tab.')
ps['A2'].font=Font(name='Calibri',italic=True,size=9)
ps['A3']='Min. benchmark to show a % — weekly & 3-wk (MT):'; ps['G3']='cumulative (MT):'
ps['F3']=25000; ps['I3']=50000; ps['J3']='signal threshold (pp):'; ps['L3']=5
ps['M3']='min. 2026 lots/week to trust a weekly read:'; ps['R3']=10
for c in ('F3','I3','L3','R3'): ps[c].font=Font(name='Calibri',color='FF0000FF',bold=True); ps[c].fill=PatternFill('solid',fgColor='FFFFFF00')
for c in ('F3','I3'): ps[c].number_format=NUM
for c in ('A3','G3','J3','M3'): ps[c].font=Font(name='Calibri',size=9,italic=True)
ps['G3'].alignment=Alignment(horizontal='right')
ps['A4']='Benchmark (choose):'; ps['A4'].font=Font(name='Calibri',bold=True)
ps['C4']='5-yr avg'; ps['C4'].font=Font(name='Calibri',color='FF0000FF',bold=True); ps['C4'].fill=PatternFill('solid',fgColor='FFFFFF00')
dv=DataValidation(type='list',formula1='"5-yr avg,10-yr avg"',allow_blank=False); ps.add_data_validation(dv); dv.add('C4')
ps['D4']='years:'; ps['E4']=f'=IF($C$4="10-yr avg",{SPs}!$H$4,{SPs}!$C$4)'; ps['F4']='to'; ps['G4']=f'=IF($C$4="10-yr avg",{SPs}!$J$4,{SPs}!$E$4)'
for c in ('D4','F4'): ps[c].font=Font(name='Calibri',size=9,italic=True); ps[c].alignment=Alignment(horizontal='right')
for c in ('E4','G4'): ps[c].font=Font(name='Calibri',bold=True); ps[c].number_format='0'
SC0=5; blocks={}; r=SC0+len(GR)+7
def header(r,title):
    ps.cell(r,1,title).font=Font(name='Calibri',bold=True,size=11)
    ps.cell(r+1,1,'Group / region'); ps.cell(r+1,2,'')
    for j,(_,a,b) in enumerate(weeks): ps.cell(r+1,C0+j,dt.datetime.strptime(b,'%Y%m%d').strftime('%-d %b'))
    for c in range(1,C0+W):
        x=ps.cell(r+1,c); x.font=HF; x.fill=HFILL; x.alignment=Alignment(horizontal='center')
    return r+2
MAIN=('Total','China','Rest of World excl. China')
def label(i,g):
    ps.cell(i,1,g).font=Font(name='Calibri',bold=g in MAIN,color=None if g in MAIN else 'FF595959')
    if g not in MAIN: ps.cell(i,1).alignment=Alignment(indent=2)
spec=[('w26','Weekly MT — 2026 (input, from raw files)'),
      ('wB','Weekly MT — benchmark average (per C4, from History Data)'),
      ('w25','Weekly MT — 2025 (input, reference)'),
      ('n26','Weekly # inspections (lots) — 2026 (input)'),
      ('c26','Cumulative MT — 2026'),('cB','Cumulative MT — benchmark average'),('c25','Cumulative MT — 2025'),
      ('yc','Cumulative vs benchmark % (headline)'),
      ('y3','3-week rolling vs benchmark %'),
      ('yw','Weekly vs benchmark % (noisy — reference)'),
      ('dyc','Gap change: week-on-week change in cumulative vs benchmark (pp) — positive = gap closing'),
      ('dy3','Week-on-week change in 3-week rolling vs benchmark (pp) — positive = recent pace improving'),
      ('dyw','Week-on-week change in weekly vs benchmark (pp) — reference; grey = few 2026 lots'),
      ('gw','Weekly MT gap vs benchmark (2026 − benchmark) = change in cumulative MT gap'),
      ('g3','3-week MT gap vs benchmark'),
      ('gc','Cumulative MT gap vs benchmark'),
      ('yc25','Cumulative vs 2025 % (2025 = trade-war year: China zero, RoW unusually high)'),
      ('y3_25','3-week rolling vs 2025 %'),
      ('dyc25','Gap change vs 2025: week-on-week change in cumulative vs 2025 (pp)'),
      ('dy3_25','Week-on-week change in 3-week rolling vs 2025 (pp)'),
      ('gw25','Weekly MT gap vs 2025 (2026 − 2025)'),('g3_25','3-week MT gap vs 2025'),('gc25','Cumulative MT gap vs 2025')]
PCTK=('yc','y3','yw','yc25','y3_25'); PPK=('dyc','dy3','dyw','dyc25','dy3_25')
PPSRC={'dyc':'yc','dy3':'y3','dyw':'yw','dyc25':'yc25','dy3_25':'y3_25'}
for key,title in spec:
    r0=header(r,title); blocks[key]=r0
    for k,g in enumerate(GR):
        i=r0+k; label(i,g)
        for j in range(W):
            col=C0+j; cl=L(col); pc=L(col-1); s0=L(col-2) if j>=2 else None
            a26,aB,a25=blocks['w26']+k,blocks.get('wB',0)+k,blocks.get('w25',0)+k
            if key=='w26': v=D26[j][g][0]
            elif key=='w25': v=D25[j][g][0]
            elif key=='n26': v=D26[j][g][1]
            elif key=='wB': v='='+normw(g,j)
            elif key in ('c26','cB','c25'):
                src_=blocks[{'c26':'w26','cB':'wB','c25':'w25'}[key]]+k
                v=f'={cl}{src_}' if j==0 else f'={pc}{i}+{cl}{src_}'
            elif key=='yc': v=f'=IF({cl}{blocks["cB"]+k}<$I$3,"",{cl}{blocks["c26"]+k}/{cl}{blocks["cB"]+k}-1)'
            elif key=='yc25': v=f'=IF({cl}{blocks["c25"]+k}<$I$3,"",{cl}{blocks["c26"]+k}/{cl}{blocks["c25"]+k}-1)'
            elif key=='yw': v=f'=IF({cl}{aB}<$F$3,"",{cl}{a26}/{cl}{aB}-1)'
            elif key=='y3': v=None if j<2 else f'=IF(SUM({s0}{aB}:{cl}{aB})<$F$3,"",SUM({s0}{a26}:{cl}{a26})/SUM({s0}{aB}:{cl}{aB})-1)'
            elif key=='y3_25': v=None if (j<2 or g=='China') else f'=IF(SUM({s0}{a25}:{cl}{a25})<$F$3,"",SUM({s0}{a26}:{cl}{a26})/SUM({s0}{a25}:{cl}{a25})-1)'
            elif key in PPK:
                srcb=blocks[PPSRC[key]]+k
                v=None if j==0 else f'=IF(OR({cl}{srcb}="",{pc}{srcb}=""),"",({cl}{srcb}-{pc}{srcb})*100)'
            elif key=='gw': v=f'={cl}{a26}-{cl}{aB}'
            elif key=='g3': v=None if j<2 else f'=SUM({s0}{a26}:{cl}{a26})-SUM({s0}{aB}:{cl}{aB})'
            elif key=='gc': v=f'={cl}{blocks["c26"]+k}-{cl}{blocks["cB"]+k}'
            elif key=='gw25': v=f'={cl}{a26}-{cl}{a25}'
            elif key=='g3_25': v=None if j<2 else f'=SUM({s0}{a26}:{cl}{a26})-SUM({s0}{a25}:{cl}{a25})'
            elif key=='gc25': v=f'={cl}{blocks["c26"]+k}-{cl}{blocks["c25"]+k}'
            x=ps.cell(i,col,v); x.font=Font(name='Calibri',bold=g=='Total')
            x.number_format=PCT if key in PCTK else (PP if key in PPK else NUM)
            if key in ('w26','w25','n26'): x.font=Font(name='Calibri',color='FF0000FF',bold=g=='Total')
    rng=f'{L(C0)}{r0}:{L(C0+W-1)}{r0+len(GR)-1}'
    if key in PCTK:
        ps.conditional_formatting.add(rng,ColorScaleRule(start_type='num',start_value=-1,start_color='E34948',mid_type='num',mid_value=0,mid_color='F0EFEC',end_type='num',end_value=1,end_color='2A78D6'))
    if key=='dyw':
        nb=blocks['n26']
        ps.conditional_formatting.add(rng,FormulaRule(formula=[f'AND(ISNUMBER({L(C0)}{r0}),OR({L(C0)}{nb}<$R$3,{L(C0-1)}{nb}<$R$3))'],font=Font(color='FFA6A6A6',italic=True),fill=PatternFill('solid',fgColor='FFF2F2F2',bgColor='FFF2F2F2'),stopIfTrue=True))
        cf(ps,rng)
    if key in ('dyc','dy3','dyc25','dy3_25'):
        ps.conditional_formatting.add(rng,ColorScaleRule(start_type='num',start_value=-15,start_color='E34948',mid_type='num',mid_value=0,mid_color='F0EFEC',end_type='num',end_value=15,end_color='2A78D6'))
    if key in ('gw','g3','gc','gw25','g3_25','gc25'): cf(ps,rng)
    r=r0+len(GR)+1
ck=r+1; ps.cell(ck,1,'Tie-out vs Summary season-to-date (should be 0)').font=B
ps.cell(ck+1,1,'Group / region').font=B; ps.cell(ck+1,2,'Δ 2025').font=B; ps.cell(ck+1,3,'Δ 2026').font=B
lastc=L(C0+W-1); SS="'Summary China vs RoW'"
for k,g in enumerate(GR):
    i=ck+2+k; ps.cell(i,1,g)
    tr={'Total':cum+TOTK,'China':cum,'Rest of World excl. China':cum+1}.get(g, cum+2+RG.index(g) if g in RG else None)
    ps.cell(i,2,f'={lastc}{blocks["c25"]+k}-{SS}!C{tr}').number_format='0;-0;0'
    ps.cell(i,3,f'={lastc}{blocks["c26"]+k}-{SS}!D{tr}').number_format='0;-0;0'
cl_=L(C0+W-1); pv=L(C0+W-2); p3=L(C0+W-4)
lastlab=dt.datetime.strptime(weeks[-1][2],'%Y%m%d').strftime('%-d %b %Y')
ps.cell(SC0,1,f'="Pace scorecard vs "&$C$4&" ("&$E$4&"–"&$G$4&") and vs 2025 — latest week ending {lastlab}"').font=Font(name='Calibri',bold=True,size=11)
sh=['Group / region','','Cumulative vs benchmark','Gap change vs last week (pp)','Gap change vs 3 weeks ago (pp)','3-wk rolling vs benchmark','3-wk change vs last week (pp)',
    'Recent vs season (3-wk − cum, pp)','MT gap this week','MT gap last 3 weeks','Cumulative MT gap','Signal (3-wk based)','Weekly vs benchmark (ref.)','Lots this week 2026',
    'vs 2025: Cumulative','vs 2025: Gap change vs last week (pp)','vs 2025: Gap change vs 3 weeks ago (pp)','vs 2025: 3-wk rolling','vs 2025: Recent vs season (pp)',
    'vs 2025: MT gap this week','vs 2025: MT gap last 3 weeks','vs 2025: Cumulative MT gap','vs 2025: Signal']
for c,h in enumerate(sh):
    x=ps.cell(SC0+1,1+c,h); x.font=HF; x.fill=HFILL; x.alignment=Alignment(horizontal='center',vertical='center',wrap_text=True)
for c in range(15,24): ps.cell(SC0+1,c).fill=PatternFill('solid',fgColor='FF7F6000')
ps.cell(SC0,15,'Columns C–N: vs benchmark (C4)  |  columns O–W: vs 2025').font=Font(name='Calibri',bold=True,size=9,color='FF7F6000')
ps.row_dimensions[SC0+1].height=58
SIG=lambda a,b_: f'=IF(OR({a}="",{b_}=""),"n/a",IF(AND({a}>$L$3,{b_}>$L$3),"▲ Improving",IF(AND({a}<-$L$3,{b_}<-$L$3),"▼ Deteriorating",IF(AND(ABS({a})<=$L$3,ABS({b_})<=$L$3),"■ Stable","◆ Mixed"))))'
for k,g in enumerate(GR):
    i=SC0+2+k; label(i,g); b=lambda key: blocks[key]+k
    F_=[(3,f'={cl_}{b("yc")}',PCT),(4,f'=IF(OR({cl_}{b("yc")}="",{pv}{b("yc")}=""),"",({cl_}{b("yc")}-{pv}{b("yc")})*100)',PP),
        (5,f'=IF(OR({cl_}{b("yc")}="",{p3}{b("yc")}=""),"",({cl_}{b("yc")}-{p3}{b("yc")})*100)',PP),
        (6,f'={cl_}{b("y3")}',PCT),(7,f'={cl_}{b("dy3")}',PP),(8,f'=IF(OR(C{i}="",F{i}=""),"",(F{i}-C{i})*100)',PP),
        (9,f'={cl_}{b("gw")}',NUM),(10,f'={cl_}{b("g3")}',NUM),(11,f'={cl_}{b("gc")}',NUM),(12,SIG(f'E{i}',f'H{i}'),None),
        (13,f'={cl_}{b("yw")}',PCT),(14,f'={cl_}{b("n26")}','0'),
        (15,f'={cl_}{b("yc25")}',PCT),(16,f'=IF(OR({cl_}{b("yc25")}="",{pv}{b("yc25")}=""),"",({cl_}{b("yc25")}-{pv}{b("yc25")})*100)',PP),
        (17,f'=IF(OR({cl_}{b("yc25")}="",{p3}{b("yc25")}=""),"",({cl_}{b("yc25")}-{p3}{b("yc25")})*100)',PP),
        (18,f'={cl_}{b("y3_25")}',PCT),(19,f'=IF(OR(O{i}="",R{i}=""),"",(R{i}-O{i})*100)',PP),
        (20,f'={cl_}{b("gw25")}',NUM),(21,f'={cl_}{b("g3_25")}',NUM),(22,f'={cl_}{b("gc25")}',NUM),(23,SIG(f'Q{i}',f'S{i}'),None)]
    for c,v,fmt in F_:
        x=ps.cell(i,c,v); x.font=Font(name='Calibri',bold=g=='Total' or c in (3,12,15,23))
        if fmt: x.number_format=fmt
    ps.cell(i,12).alignment=Alignment(horizontal='center'); ps.cell(i,23).alignment=Alignment(horizontal='center')
    if g=='China':
        ps.cell(i,15,'n/a — no 2025 base'); ps.cell(i,23,'n/a')
        for c in range(16,20): ps.cell(i,c).value=None
n_=SC0+1+len(GR)
cf(ps,f'C{SC0+2}:K{n_}'); cf(ps,f'O{SC0+2}:V{n_}')
ps.conditional_formatting.add(f'M{SC0+2}:M{n_}',FormulaRule(formula=[f'AND(ISNUMBER(M{SC0+2}),N{SC0+2}<$R$3)'],font=Font(color='FFA6A6A6',italic=True),stopIfTrue=True))
for col_ in ('L','W'):
    ps.conditional_formatting.add(f'{col_}{SC0+2}:{col_}{n_}',FormulaRule(formula=[f'LEFT({col_}{SC0+2},1)="▲"'],font=Font(color='FF2A78D6',bold=True)))
    ps.conditional_formatting.add(f'{col_}{SC0+2}:{col_}{n_}',FormulaRule(formula=[f'LEFT({col_}{SC0+2},1)="▼"'],font=Font(color='FFC00000',bold=True)))
for i in range(SC0+2,n_+1): ps.cell(i,13).font=Font(name='Calibri',color='FF7F7F7F',italic=True)
notes=['How to read: cumulative vs benchmark is the headline. Gap change = how much it moved (positive = closing on the benchmark). The 3-wk measures react faster but ignore single vessels. MT gaps say "how much": this week\'s MT gap is exactly the change in the cumulative MT gap.',
       'Signal (threshold T in L3): ▲ Improving = cumulative vs benchmark up > T pp vs 3 weeks ago AND last 3 weeks running > T pp ahead of the season average; ▼ Deteriorating = both < −T; ■ Stable = both within ±T; ◆ Mixed = otherwise.',
       'Weekly vs benchmark (grey) is reference only — dominated by vessel timing. Columns O–W repeat the same measures vs 2025 (same signal rule). 2025 was a trade-war year (China zero, RoW unusually high); the 5-yr and 10-yr averages also include trade-war years.']
for q,t in enumerate(notes): ps.cell(n_+1+q,1,t).font=Font(name='Calibri',italic=True,size=9)
ps.column_dimensions['A'].width=28; ps.column_dimensions['B'].width=3
for j in range(21): ps.column_dimensions[L(C0+j)].width=12.5
ps.column_dimensions['L'].width=16; ps.column_dimensions['K'].width=13.5; ps.column_dimensions['W'].width=16
ps.freeze_panes=f'{L(C0)}5'
CHC=max(25,C0+W+1); ch_col=L(CHC); cats=Reference(ps,min_col=C0,max_col=C0+W-1,min_row=blocks['yc']-1)
PAL=['2A78D6','EB6834','1BAF7A','EDA100','E87BA4','008300','6250D6','E34948']
bc=BarChart(); bc.type='bar'; bc.title='Gap change vs benchmark (pp) — positive = closing'
scr=[GR.index(g) for g in GR if g not in ('Sub-Saharan Africa','Total')]
r1=SC0+2+min(scr); r2=SC0+2+max(scr)
for colx,colr,lab in ((4,'2A78D6','vs last week'),(5,'EB6834','vs 3 weeks ago')):
    bc.add_data(Reference(ps,min_col=colx,min_row=r1,max_row=r2),titles_from_data=False)
    sx=bc.series[-1]; sx.tx=SeriesLabel(v=lab); sx.graphicalProperties.solidFill=colr; sx.graphicalProperties.line.solidFill='FFFFFF'
bc.set_categories(Reference(ps,min_col=1,min_row=r1,max_row=r2)); bc.y_axis.numFmt='0'; bc.y_axis.title='pp'
bc.x_axis.scaling.orientation='maxMin'; bc.x_axis.tickLblPos='low'; bc.legend.position='b'
bc.y_axis.majorGridlines.spPr=GraphicalProperties(ln=LineProperties(solidFill='E5E4E0'))
bc.x_axis.delete=False; bc.y_axis.delete=False; bc.width=20; bc.height=9
ps.add_chart(bc,f'{ch_col}{SC0}')
def linechart(title,key,ytitle,glist,greycol):
    lc=LineChart(); lc.title=title
    for n,g in enumerate(glist):
        i=blocks[key]+GR.index(g)
        lc.add_data(Reference(ps,min_col=C0,max_col=C0+W-1,min_row=i),from_rows=True,titles_from_data=False)
        s_=lc.series[-1]; s_.tx=SeriesLabel(v=g); s_.smooth=False
        col=greycol.get(g) or PAL[n%8]
        thick=g in greycol
        s_.graphicalProperties.line.solidFill=col; s_.graphicalProperties.line.width=38100 if thick else 22225
        if thick: s_.graphicalProperties.line.dashStyle='dash'
        s_.marker.symbol='circle'; s_.marker.size=5; s_.marker.graphicalProperties=GraphicalProperties(solidFill=col); s_.marker.graphicalProperties.line.solidFill=col
    lc.set_categories(cats); lc.y_axis.numFmt='0%'; lc.y_axis.title=ytitle; lc.x_axis.title='Week ending (2026)'
    lc.y_axis.majorGridlines.spPr=GraphicalProperties(ln=LineProperties(solidFill='E5E4E0'))
    lc.x_axis.delete=False; lc.y_axis.delete=False; lc.legend.position='b'; lc.width=24; lc.height=11; lc.x_axis.tickLblPos='low'
    return lc
regs=[g for g in RG if g!='Sub-Saharan Africa']
ps.add_chart(linechart('Cumulative vs benchmark average (per C4), by region','yc','vs benchmark (cumulative)',['Rest of World excl. China','China']+regs,{'Rest of World excl. China':'000000','China':'8C8B86'}),f'{ch_col}25')
ps.add_chart(linechart('Cumulative vs 2025, by region','yc25','vs 2025 (cumulative)',['Rest of World excl. China']+regs,{'Rest of World excl. China':'000000'}),f'{ch_col}48')
for n,g in enumerate(CHARTR):
    k=GR.index(g); c=LineChart(); c.title=g
    for key,col,lab,dash,wd in (('c25','A6A6A6','2025','sysDot',19050),('cB','6250D6','Benchmark avg','dash',25400),('c26','2A78D6' if g!='China' else 'EB6834','2026',None,31750)):
        c.add_data(Reference(ps,min_col=C0,max_col=C0+W-1,min_row=blocks[key]+k),from_rows=True,titles_from_data=False)
        s_=c.series[-1]; s_.tx=SeriesLabel(v=lab); s_.smooth=False; s_.graphicalProperties.line.solidFill=col; s_.graphicalProperties.line.width=wd; s_.marker.symbol='none'
        if dash: s_.graphicalProperties.line.dashStyle=dash
    c.set_categories(cats); c.y_axis.numFmt='#,##0,"k"'; c.legend.position='b'
    c.y_axis.majorGridlines.spPr=GraphicalProperties(ln=LineProperties(solidFill='E5E4E0'))
    c.x_axis.delete=False; c.y_axis.delete=False; c.width=12; c.height=7.5
    ps.add_chart(c,f'{L(CHC+(n%2)*8)}{74+(n//2)*16}')
ps.cell(72,CHC,'Cumulative MT: 2026 solid, benchmark avg (per C4) violet dashed, 2025 grey dotted. Own scale per panel.').font=Font(name='Calibri',italic=True,size=9)
