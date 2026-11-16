ws=wb.create_sheet('Weekly China vs RoW',2)
ws['A1']='US Soybean Export Inspections — China vs Rest of World by week, 2025 vs 2026 (MT)'; ws['A1'].font=Font(name='Calibri',bold=True,size=12)
ws['A2']='Weekly MT computed from raw CY2025/CY2026 FGIS files (Grain = SOYBEANS, by "Thursday" week-ending date). Week-ending dates are paired Thursday-to-Thursday (2025 week ends one day later). Check columns must read 0 (tie-out to period tabs).'
ws['A2'].font=Font(name='Calibri',italic=True,size=9)
H=['Week','Label','w/e 2025','w/e 2026','China 2025','RoW 2025','Total 2025','China 2026','RoW 2026','Total 2026','Δ RoW MT','Δ RoW %','Δ Total MT','China share 2026','Period tab']
for k,h in enumerate(H):
    x=ws.cell(4,1+k,h); x.font=HF; x.fill=HFILL; x.alignment=Alignment(horizontal='center',vertical='center',wrap_text=True)
ws.row_dimensions[4].height=30
def cr(x,w): y=x[x.Thursday==w]; c=y[y.Destination=='CHINA'].MT.sum(); return int(c), int(y.MT.sum()-c)
r0=5
for i,(name,a,b) in enumerate(weeks):
    r=r0+i; da=dt.datetime.strptime(a,'%Y%m%d'); db=dt.datetime.strptime(b,'%Y%m%d')
    ws[f'A{r}']=i+1; ws[f'B{r}']=db.strftime('%-d %b'); ws[f'C{r}']=da.date(); ws[f'D{r}']=db.date()
    c5,o5=cr(x25,a); c6,o6=cr(x26,b)
    ws[f'E{r}']=c5; ws[f'F{r}']=o5; ws[f'G{r}']=f'=E{r}+F{r}'; ws[f'H{r}']=c6; ws[f'I{r}']=o6; ws[f'J{r}']=f'=H{r}+I{r}'
    ws[f'K{r}']=f'=I{r}-F{r}'; ws[f'L{r}']=f'=IF(F{r}=0,"new",I{r}/F{r}-1)'; ws[f'M{r}']=f'=J{r}-G{r}'; ws[f'N{r}']=f'=IF(J{r}=0,0,H{r}/J{r})'; ws[f'O{r}']=name
rl=r0+len(weeks)-1; rt=rl+1
ws[f'A{rt}']='TOTAL'
for c in 'EFGHIJKM': ws[f'{c}{rt}']=f'=SUM({c}{r0}:{c}{rl})'
ws[f'L{rt}']=f'=IF(F{rt}=0,"new",I{rt}/F{rt}-1)'; ws[f'N{rt}']=f'=IF(J{rt}=0,0,H{rt}/J{rt})'
for r in range(r0,rt+1):
    for c in range(1,16):
        x=ws.cell(r,c); x.font=Font(name='Calibri',bold=r==rt)
        if r==rt: x.fill=TFILL
    for c in 'EFGHIJKM': ws[f'{c}{r}'].number_format=NUM
    for c in 'CD': ws[f'{c}{r}'].number_format='dd-mmm-yy'
    ws[f'L{r}'].number_format=PCT; ws[f'L{r}'].alignment=Alignment(horizontal='right'); ws[f'N{r}'].number_format='0.0%'
cf(ws,f'K{r0}:M{rt}')
ck=rt+2; ws[f'A{ck}']='Tie-out check vs period tabs (should be 0):'; ws[f'A{ck}'].font=B
ws[f'A{ck+1}']='Period'; ws[f'B{ck+1}']='Δ Total 2025'; ws[f'C{ck+1}']='Δ Total 2026'; ws[f'D{ck+1}']='Δ China 2026'
for c in 'ABCD': ws[f'{c}{ck+1}'].font=B
for k,(name,*_) in enumerate(periods):
    r=ck+2+k; t=refs[name][2]; ch=refs[name][0]; q=f"'{name}'"
    ws[f'A{r}']=name
    ws[f'B{r}']=f'=SUMIF($O${r0}:$O${rl},A{r},$G${r0}:$G${rl})-{q}!B{t}'
    ws[f'C{r}']=f'=SUMIF($O${r0}:$O${rl},A{r},$J${r0}:$J${rl})-{q}!C{t}'
    ws[f'D{r}']=f'=SUMIF($O${r0}:$O${rl},A{r},$H${r0}:$H${rl})-{q}!C{ch}'
    for c in 'BCD': ws[f'{c}{r}'].number_format='0;-0;0'
for c,wd in zip('ABCDEFGHIJKLMNO',[14,13,13,13,11,11,11,11,11,11,11,9,11,10,16]): ws.column_dimensions[c].width=wd
ws.freeze_panes='C5'
bar=BarChart(); bar.type='col'; bar.grouping='stacked'; bar.overlap=100; bar.gapWidth=60
bar.add_data(Reference(ws,min_col=9,min_row=4,max_row=rl),titles_from_data=True)
bar.add_data(Reference(ws,min_col=8,min_row=4,max_row=rl),titles_from_data=True)
cats=Reference(ws,min_col=2,min_row=r0,max_row=rl); bar.set_categories(cats)
for s_,col in zip(bar.series,['2A78D6','EB6834']):
    s_.graphicalProperties.solidFill=col; s_.graphicalProperties.line.solidFill='FFFFFF'; s_.graphicalProperties.line.width=19050
bar.series[0].tx=SeriesLabel(v='RoW 2026'); bar.series[1].tx=SeriesLabel(v='China 2026')
ln=LineChart(); ln.add_data(Reference(ws,min_col=7,min_row=4,max_row=rl),titles_from_data=True); ln.set_categories(cats)
s_=ln.series[0]; s_.graphicalProperties.line.solidFill='52514E'; s_.graphicalProperties.line.width=28575; s_.graphicalProperties.line.dashStyle='dash'
s_.marker.symbol='circle'; s_.marker.size=7; s_.marker.graphicalProperties=GraphicalProperties(solidFill='52514E'); s_.marker.graphicalProperties.line.solidFill='52514E'; s_.smooth=False
s_.tx=SeriesLabel(v='2025 total (all RoW — China 0)')
ln.y_axis.axId=100; ln.y_axis.crosses='autoZero'
bar.title='Soybean inspections by week: China vs RoW 2026 vs 2025 (MT)'
bar.y_axis.title='Metric tons'; bar.x_axis.title='Week ending (2026; 2025 week ends one day later)'
bar.y_axis.numFmt='#,##0'; bar.y_axis.majorGridlines.spPr=GraphicalProperties(ln=LineProperties(solidFill='E5E4E0'))
bar.x_axis.delete=False; bar.y_axis.delete=False; bar.legend.position='b'
bar += ln
bar.width=26; bar.height=13
ws.add_chart(bar,f'A{ck+2+len(periods)+2}')
