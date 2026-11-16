def avgsum(g,ks,a,b):
    cond=f'({GA}="{g}")*({YA}>={SPs}!{a})*({YA}<={SPs}!{b})'
    tot=wkcol(ks[0]) if len(ks)==1 else '('+'+'.join(wkcol(k) for k in ks)+')'
    return f'SUMPRODUCT({cond}*{tot})/SUMPRODUCT({cond})'
for c,h in zip('MNOPQR',['5-yr avg MT','Δ MT 2026 vs 5-yr','2026 vs 5-yr','10-yr avg MT','Δ MT 2026 vs 10-yr','2026 vs 10-yr']):
    x=summ[f'{c}4']; x.value=h; x.font=HF; x.fill=PatternFill('solid',fgColor='FF375623'); x.alignment=Alignment(horizontal='center',vertical='center',wrap_text=True)
pk={name:[j for j,(n,a,b) in enumerate(weeks) if n==name] for name,*_ in periods}
allks=list(range(len(weeks)))
for bi,b0 in enumerate(blocks_summ):
    ks=pk[periods[bi][0]] if bi<len(periods) else allks
    for k,g in enumerate(labels):
        i=b0+k
        for (cm,cd,cp),(a,b) in ((('M','N','O'),('$C$4','$E$4')),(('P','Q','R'),('$H$4','$J$4'))):
            summ[f'{cm}{i}']='='+avgsum(g,ks,a,b); summ[f'{cd}{i}']=f'=D{i}-{cm}{i}'
            summ[f'{cp}{i}']=f'=IF({cm}{i}<1000,"",D{i}/{cm}{i}-1)'
            summ[f'{cm}{i}'].number_format=NUM; summ[f'{cd}{i}'].number_format=NUM
            summ[f'{cp}{i}'].number_format=PCT; summ[f'{cp}{i}'].alignment=Alignment(horizontal='right')
        for c in 'MNOPQR':
            summ[f'{c}{i}'].font=Font(name='Calibri',bold=(k in (1,TOTK)))
            if k==TOTK: summ[f'{c}{i}'].fill=TFILL
cf(summ,f'N5:O{last}'); cf(summ,f'Q5:R{last}')
for c in 'MNOPQR': summ.column_dimensions[c].width=13
summ['A3']=('Green columns (M–R): simple 5-yr (2021–25) and 10-yr (2016–25) averages for the same season weeks, from the History Data tab (windows editable on Seasonal Pace row 4). '
            'Both include trade-war years (2018–19, 2025: China largely absent, RoW unusually high).')
summ['A3'].font=Font(name='Calibri',italic=True,size=9,color='FF375623')
