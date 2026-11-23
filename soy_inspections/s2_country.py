REG={'CHINA':'China'}
for rg,cs in {'East Asia ex-China':['JAPAN','TAIWAN','KOREA REP','HONG KONG'],
 'Southeast Asia':['INDONESIA','VIETNAM','THAILAND','MALAYSIA','PHILIPPINES','CAMBODIA','SINGAPORE','LAOS','BURMA','COCOS IS','VANUATU'],
 'South Asia':['PAKISTAN','BANGLADESH','NEPAL','SRI LANKA','INDIA','MALDIVES'],
 'Middle East & N. Africa':['EGYPT','ALGERIA','TUNISIA','IRAQ','TURKEY','ISRAEL','SAUDI ARABIA','LEBANON','IRAN','MOROCCO','JORDAN','UN ARAB EM','YEMEN'],
 'Europe':['SPAIN','ITALY','GERMANY','NETHERLANDS','UN KINGDOM','GREECE','PORTUGAL','FRANCE','BELGIUM','ROMANIA','RUSSIA','UKRAINE','GEORGIA','IRELAND'],
 'Americas':['MEXICO','COLOMBIA','COSTA RICA','PERU','VENEZUELA','DOMINICN REP','GRENADA','PANAMA','CANADA','EL SALVADOR','ARGENTINA','CHILE','CUBA','BARBADOS','BRAZIL','NICARAGUA','TRINIDAD','GUATEMALA','VIRGIN ISLAN'],
 'Sub-Saharan Africa':['SENEGAL','REP S AFRICA','NIGERIA','MALAWI']}.items():
    for c in cs: REG[c]=rg
allw25=sum([p[2] for p in periods],[]); allw26=sum([p[3] for p in periods],[])
A=agg(d25,allw25); Bq=agg(d26,allw26); T=A.join(Bq,lsuffix='25',rsuffix='26',how='outer').fillna(0)
T['d']=T.MT26-T.MT25; T['ch']=T.index!='CHINA'; T=T.sort_values(['ch','d'])
missing=[c for c in T.index if c not in REG]
if missing: print('UNMAPPED',missing)
ws=wb.create_sheet('Country by Country',1)
ws['A1']=f'US Soybean Export Inspections by Destination — {COUNTRY_TITLE}, 2025 vs 2026'; ws['A1'].font=Font(name='Calibri',bold=True,size=12)
ws['A2']='Period columns link to the period tabs (SUMIF on destination). Sorted: China first, then by season-to-date Δ MT (largest loss first). Region grouping is an analyst assumption. Bulk buyers take 1–3 vessels/season: read single-week country changes as lumpy, not trend.'
ws['A2'].font=Font(name='Calibri',italic=True,size=9)
ws.merge_cells('A4:A5'); ws.merge_cells('B4:B5'); ws['A4']='Destination'; ws['B4']='Region'
col=3; pcols=[]
for name,short,*_ in periods:
    ws.merge_cells(start_row=4,start_column=col,end_row=4,end_column=col+1)
    ws.cell(4,col,short+' (MT)'); ws.cell(5,col,'2025'); ws.cell(5,col+1,'2026'); pcols.append((name,col)); col+=2
S=col
hdrs=['MT 2025','MT 2026','Δ MT','Δ %','# Insp 2025','# Insp 2026','Δ Insp','Share of RoW Δ']
ws.merge_cells(start_row=4,start_column=S,end_row=4,end_column=S+len(hdrs)-1); ws.cell(4,S,'Season to date')
for k,h in enumerate(hdrs): ws.cell(5,S+k,h)
last_col=S+len(hdrs)-1
for r in (4,5):
    for c in range(1,last_col+1):
        x=ws.cell(r,c); x.font=HF; x.fill=HFILL; x.alignment=Alignment(horizontal='center',vertical='center',wrap_text=True)
c25,c26,cd,cp,ci25,ci26,cdi,csh=[L(S+k) for k in range(8)]
first=6; lastr=first+len(T)-1; tot=lastr+1; chn=tot+1; row_=tot+2
for i,dest in enumerate(T.index):
    r=first+i; ws[f'A{r}']=dest; ws[f'B{r}']=REG.get(dest,'Other')
    p25=[];p26=[];n25=[];n26=[]
    for name,pc in pcols:
        _,_,ptot=refs[name]; pl=ptot-1; q=f"'{name}'"; rng=f"{q}!$A$6:$A${pl}"
        ws.cell(r,pc,f'=SUMIF({rng},$A{r},{q}!$B$6:$B${pl})'); ws.cell(r,pc+1,f'=SUMIF({rng},$A{r},{q}!$C$6:$C${pl})')
        p25.append(L(pc)+str(r)); p26.append(L(pc+1)+str(r))
        n25.append(f'SUMIF({rng},$A{r},{q}!$F$6:$F${pl})'); n26.append(f'SUMIF({rng},$A{r},{q}!$G$6:$G${pl})')
    ws[f'{c25}{r}']='='+'+'.join(p25); ws[f'{c26}{r}']='='+'+'.join(p26)
    ws[f'{ci25}{r}']='='+'+'.join(n25); ws[f'{ci26}{r}']='='+'+'.join(n26)
ws[f'A{tot}']='TOTAL'; ws[f'A{chn}']='China'; ws[f'A{row_}']='Rest of World excl. China'
for c in list(range(3,S+2))+[S+4,S+5]:
    cl=L(c); ws[f'{cl}{tot}']=f'=SUM({cl}{first}:{cl}{lastr})'; ws[f'{cl}{chn}']=f'=SUMIF($A${first}:$A${lastr},"CHINA",{cl}{first}:{cl}{lastr})'; ws[f'{cl}{row_}']=f'={cl}{tot}-{cl}{chn}'
for r in range(first,row_+1):
    ws[f'{cd}{r}']=f'={c26}{r}-{c25}{r}'; ws[f'{cdi}{r}']=f'={ci26}{r}-{ci25}{r}'
    ws[f'{cp}{r}']=f'=IF({c25}{r}=0,IF({c26}{r}=0,0,"new"),{c26}{r}/{c25}{r}-1)'
    if r!=chn and r!=tot: ws[f'{csh}{r}']=f'=IF(OR($A{r}="CHINA",${cd}${row_}=0),"",{cd}{r}/${cd}${row_})'
    for c in range(1,last_col+1):
        x=ws.cell(r,c); x.font=Font(name='Calibri',bold=r>=tot)
        if c>=3: x.number_format=NUM
        if r==tot: x.fill=TFILL
    ws[f'{cp}{r}'].number_format=PCT; ws[f'{cp}{r}'].alignment=Alignment(horizontal='right'); ws[f'{csh}{r}'].number_format='0.0%'
cf(ws,f'{cd}{first}:{cp}{row_}'); cf(ws,f'{cdi}{first}:{cdi}{row_}')
g=row_+3
ws[f'A{g-1}']='By region (season to date)'; ws[f'A{g-1}'].font=B
rh=['Region','# countries','MT 2025','MT 2026','Δ MT','Δ %','Share of RoW Δ']
for k,h in enumerate(rh):
    x=ws.cell(g,1+k,h); x.font=HF; x.fill=HFILL; x.alignment=Alignment(horizontal='center',wrap_text=True)
regions=['China','East Asia ex-China','Southeast Asia','South Asia','Middle East & N. Africa','Europe','Americas','Sub-Saharan Africa']+(['Other'] if missing else [])
bR=f'$B${first}:$B${lastr}'
for k,rg in enumerate(regions):
    r=g+1+k; ws[f'A{r}']=rg; ws[f'B{r}']=f'=COUNTIF({bR},A{r})'
    ws[f'C{r}']=f'=SUMIF({bR},$A{r},{c25}${first}:{c25}${lastr})'; ws[f'D{r}']=f'=SUMIF({bR},$A{r},{c26}${first}:{c26}${lastr})'
    ws[f'E{r}']=f'=D{r}-C{r}'; ws[f'F{r}']=f'=IF(C{r}=0,IF(D{r}=0,0,"new"),D{r}/C{r}-1)'
    ws[f'G{r}']=f'=IF(A{r}="China","",E{r}/${cd}${row_})'
rt=g+1+len(regions); ws[f'A{rt}']='TOTAL'
for cl in 'BCDE': ws[f'{cl}{rt}']=f'=SUM({cl}{g+1}:{cl}{rt-1})'
ws[f'F{rt}']=f'=IF(C{rt}=0,0,D{rt}/C{rt}-1)'
for r in range(g+1,rt+1):
    for c in range(1,8):
        x=ws.cell(r,c); x.font=Font(name='Calibri',bold=r==rt); x.number_format=NUM
        if r==rt: x.fill=TFILL
    ws[f'F{r}'].number_format=PCT; ws[f'F{r}'].alignment=Alignment(horizontal='right'); ws[f'G{r}'].number_format='0.0%'
cf(ws,f'E{g+1}:F{rt}')
ws.column_dimensions['A'].width=26; ws.column_dimensions['B'].width=22
for c in range(3,last_col+1): ws.column_dimensions[L(c)].width=11.5
ws.row_dimensions[5].height=30; ws.freeze_panes='C6'
CF_first,CF_last=first,lastr
