s=summ; CF,CLr=CF_first,CF_last; CS="'Country by Country'"
RG=[x for x in regions if x!='China']
s['A1']='US Soybean Export Inspections — China vs Rest of World, RoW split by region (weekly-report basis)'; s['A1'].font=Font(name='Calibri',bold=True,size=12)
s['A2']='All figures link to the period tabs. Region rows use the region mapping on the Country by Country tab (analyst assumption); they are grouped and can be collapsed with the outline buttons. Soybean meal is not covered: FGIS inspection data include grains/oilseeds only.'
s['A2'].font=Font(name='Calibri',italic=True,size=9)
hdr=['Period','Group / region','MT 2025','MT 2026','Δ MT','Δ %','# Insp 2025','# Insp 2026','Δ Insp','Share of MT 2025','Share of MT 2026','Region tie-out (0 = OK)']
for i,h in enumerate(hdr):
    c=s.cell(4,i+1,h); c.font=HF; c.fill=HFILL; c.alignment=Alignment(horizontal='center',vertical='center',wrap_text=True)
s.row_dimensions[4].height=30
NB=3+len(RG); TOTK=NB-1
labels=['China','Rest of World excl. China']+RG+['Total']
src={'C':'B','D':'C','G':'F','H':'G'}
r=5; blocks=[]
for name,*_ in periods:
    chn,row_,tot=refs[name]; q=f"'{name}'"; pl=tot-1
    s[f'A{r}']=name
    for k,gname in enumerate(labels):
        i=r+k; s[f'B{i}']=gname
        for dc,sc in src.items():
            if k==0: s[f'{dc}{i}']=f'={q}!{sc}{chn}'
            elif k==1: s[f'{dc}{i}']=f'={q}!{sc}{row_}'
            elif k==TOTK: s[f'{dc}{i}']=f'={q}!{sc}{tot}'
            else: s[f'{dc}{i}']=f'=SUMPRODUCT(({CS}!$B${CF}:$B${CLr}=$B{i})*SUMIF({q}!$A$6:$A${pl},{CS}!$A${CF}:$A${CLr},{q}!${sc}$6:${sc}${pl}))'
    blocks.append(r); r+=NB
cum=r; s[f'A{cum}']=SEASON_LABEL
for k,gname in enumerate(labels):
    s[f'B{cum+k}']=gname
    for dc in src: s[f'{dc}{cum+k}']='='+'+'.join(f'{dc}{b+k}' for b in blocks)
blocks.append(cum); blocks_summ=list(blocks)
GREY=Font(name='Calibri',color='FF595959')
for b in blocks:
    t=b+TOTK
    for k in range(NB):
        i=b+k
        s[f'E{i}']=f'=D{i}-C{i}'; s[f'I{i}']=f'=H{i}-G{i}'
        s[f'F{i}']=f'=IF(C{i}=0,IF(D{i}=0,0,"new"),D{i}/C{i}-1)'
        s[f'J{i}']=f'=IF(C${t}=0,0,C{i}/C${t})'; s[f'K{i}']=f'=IF(D${t}=0,0,D{i}/D${t})'
        for c in 'CDEGHI': s[f'{c}{i}'].number_format=NUM
        s[f'F{i}'].number_format=PCT; s[f'F{i}'].alignment=Alignment(horizontal='right')
        for c in 'JK': s[f'{c}{i}'].number_format='0.0%'
        isreg=2<=k<TOTK
        for c in 'ABCDEFGHIJK':
            s[f'{c}{i}'].font=Font(name='Calibri',bold=(k in (1,TOTK) or c=='A'))
            if k==TOTK and c!='A': s[f'{c}{i}'].fill=TFILL
        if isreg:
            s[f'B{i}'].font=GREY; s[f'B{i}'].alignment=Alignment(indent=2); s.row_dimensions[i].outlineLevel=1
    rr=b+1
    s[f'L{rr}']=f'=ABS(SUM(C{b+2}:C{t-1})-C{rr})+ABS(SUM(D{b+2}:D{t-1})-D{rr})+ABS(SUM(G{b+2}:G{t-1})-G{rr})+ABS(SUM(H{b+2}:H{t-1})-H{rr})'
    s[f'L{rr}'].number_format='0;-0;0'; s[f'L{rr}'].font=Font(name='Calibri'); s[f'L{rr}'].alignment=Alignment(horizontal='center')
s.sheet_properties.outlinePr.summaryBelow=False
last=cum+NB-1
cf(s,f'E5:F{last}'); cf(s,f'I5:I{last}')
s.column_dimensions['A'].width=22; s.column_dimensions['B'].width=28
for c in 'CDEFGHIJK': s.column_dimensions[c].width=12
s.column_dimensions['L'].width=13
s.freeze_panes='C5'
# revisions vs previous workbook (auto)
def prev_vals():
    out={}
    try: pw=load_workbook(PREV,data_only=True)['Summary China vs RoW']
    except Exception as e: print('no prev workbook',e); return out
    per=None
    for row in pw.iter_rows(min_row=5,values_only=True):
        if row[0]: per=row[0]
        if row[1] in ('China','Rest of World excl. China','Total') and per: out[(per,row[1])]=(row[3],row[7])
    return out
PV=prev_vals(); CUR={}
for name,short,w25,w26,_ in periods:
    y=x26[x26.Thursday.isin(w26)]; c=y[y.Destination=='CHINA']
    CUR[(name,'China')]=(int(c.MT.sum()),len(c)); CUR[(name,'Total')]=(int(y.MT.sum()),len(y))
    CUR[(name,'Rest of World excl. China')]=(int(y.MT.sum()-c.MT.sum()),len(y)-len(c))
revs=[]
for name,short,*_ in periods:
    if (name,'Total') not in PV: continue
    dT=CUR[(name,'Total')][0]-PV[(name,'Total')][0]; dC=CUR[(name,'China')][0]-PV[(name,'China')][0]; dn=CUR[(name,'Total')][1]-PV[(name,'Total')][1]
    if dT or dn: revs.append(f'{short} {dT:+,} MT ({dn:+} insp; China {dC:+,})')
n=last+2
s[f'A{n}']=f'Revision note (CY2026 file, wk ending {dt.datetime.strptime(weeks[-1][2],"%Y%m%d").strftime("%-d %b %Y")} vs previous version of this workbook):'; s[f'A{n}'].font=B
s[f'A{n+1}']=('Revisions this update: '+'; '.join(revs)+'. Other weeks unchanged.') if revs else 'No revisions to previously reported weeks.'
s[f'A{n+2}']=EARLIER_REVISIONS
for i in (n+1,n+2): s[f'A{i}'].font=Font(name='Calibri',size=9)
print('REVISIONS:',revs)
