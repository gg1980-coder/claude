import sys,pandas as pd, matplotlib
matplotlib.use('Agg'); import matplotlib.pyplot as plt
df=pd.read_csv(sys.argv[1],thousands=','); df.columns=[c.strip() for c in df.columns]
df['C']=df.CountryName.str.strip(); df['d']=pd.to_datetime(df.WeekendingDate)
df['Y']=df.ReportingMarketingYearName.str.extract(r'Sep (\d{4})').astype(int)
lo=pd.to_datetime(df.Y.astype(str)+'-08-31'); hi=pd.to_datetime((df.Y+1).astype(str)+'-09-07')
EU=['SPAIN','NETHERLANDS','PORTUGAL','ITALY','GERMANY','GREECE','BELGIUM-LUXEMBOURG','FRANCE','IRELAND','ROMANIA']
ins=df[(df.d>=lo)&(df.d<=hi)&(df.Y.between(2020,2026))&df.C.isin(EU)]
lab=lambda y:f'{y}/{str(y+1)[-2:]}'
w=ins[ins.WeekNumber.between(10,26)].groupby('Y')[['NetSales','WeeklyExports']].sum()
cum=ins.groupby(['Y','WeekNumber']).TotalCommitment.sum().unstack(0)
chk=(cum.loc[26]-cum.loc[9])
fin=ins.sort_values('d').groupby(['Y','C']).tail(1).groupby('Y').AccumulatedExports.sum()
dates=ins[ins.WeekNumber.isin([10,26])].groupby(['Y','WeekNumber']).d.max().unstack()
t=pd.DataFrame({'net_sales_wk10_26':w.NetSales,'commit_change_wk9_26':chk,'shipments_wk10_26':w.WeeklyExports,'full_year_exports':fin})
t['net_sales_share_of_year']=t.net_sales_wk10_26/t.full_year_exports
t=t.join(dates.rename(columns={10:'wk10',26:'wk26'}))
print(t.round(3).to_string())
t=t.loc[2020:2025]
yrs=[lab(y) for y in t.index]
BLUE,ORANGE,INK,INK2,GRID='#2a78d6','#eb6834','#0b0b0b','#52514e','#e4e3df'
plt.rcParams.update({'font.family':'DejaVu Sans','font.size':10,'axes.edgecolor':GRID,'axes.labelcolor':INK2,'xtick.color':INK2,'ytick.color':INK2})
fig,(a,b)=plt.subplots(1,2,figsize=(13,5.2),gridspec_kw={'width_ratios':[1.1,1]},facecolor='#fcfcfb')
for ax in (a,b): ax.set_facecolor('#fcfcfb'); ax.spines[['top','right']].set_visible(False); ax.grid(axis='y',color=GRID,lw=0.8); ax.set_axisbelow(True)
import numpy as np
x=np.arange(len(t)); bw=0.36
b1=a.bar(x-bw/2-0.01,t.net_sales_wk10_26/1e6,bw,color=BLUE,label='Net sales (new bookings)')
b2=a.bar(x+bw/2+0.01,t.shipments_wk10_26/1e6,bw,color=ORANGE,label='Shipments (exports)')
for bars in (b1,b2):
    for r in bars: a.text(r.get_x()+r.get_width()/2,r.get_height()+0.03,f'{r.get_height():.2f}',ha='center',fontsize=8.5,color=INK2)
avg=t.net_sales_wk10_26.loc[2021:2025].mean()/1e6
a.axhline(avg,color=INK2,lw=1,ls=(0,(4,3)),zorder=1)
a.text(len(t)-0.45,avg-0.05,f'5-yr avg net sales\n{avg:.2f} Mt',ha='left',va='top',fontsize=8.5,color=INK2)
a.set_xlim(-0.6,len(t)+0.4); a.set_xticks(x,yrs); a.set_ylabel('Million metric tons'); a.set_ylim(0,3.75); a.legend(frameon=False,loc='upper center',ncol=2,fontsize=9)
a.set_title('EU-27: US soybean net sales and shipments, weeks 10–26 (Nov–Feb)',fontsize=11,color=INK,loc='left')
# right: cumulative net sales from wk10 (commitment change vs wk9)
wk=range(9,27)
base=cum.loc[list(wk)].sub(cum.loc[9],axis=1)/1e6
LBL=[]
for y in range(2020,2025):
    b.plot(list(wk),base[y],color='#c3c2b7',lw=1.2)
    LBL.append([base[y].iloc[-1],lab(y),INK2,'normal'])
b.plot(list(wk),base[2025],color=ORANGE,lw=2,label=lab(2025))
LBL.append([base[2025].iloc[-1],lab(2025),INK,'bold'])
av=base[[2021,2022,2023,2024,2025]].mean(axis=1)
b.plot(list(wk),av,color=BLUE,lw=2.2,label='5-yr avg (2021/22–2025/26)')
LBL.append([av.iloc[-1],'5-yr avg',INK,'bold'])
LBL.sort(key=lambda z:z[0])
for i in range(1,len(LBL)):
    if LBL[i][0]-LBL[i-1][0]<0.09: LBL[i][0]=LBL[i-1][0]+0.09
for yv,t_,c_,w_ in LBL: b.text(26.25,yv,t_,fontsize=8.5,color=c_,va='center',fontweight=w_)
b.set_xlim(9,28.2); b.set_xticks([9,10,14,18,22,26]); b.set_xlabel('Marketing-year week (week 9 = end of Oct, 26 = end of Feb)')
b.set_ylabel('Cumulative net new commitment since week 9 (Mt)')
b.set_title('Cumulative build of EU-27 bookings, weeks 10–26',fontsize=11,color=INK,loc='left')
b.legend(frameon=False,loc='upper left',fontsize=9)
fig.text(0.01,0.01,'Source: USDA FAS Export Sales Reporting (ESR), SOYBEANS. EU-27 = Spain, Netherlands, Germany, Portugal, Italy, Greece, France, Belgium-Lux., Ireland, Romania (as reported). '
         'Net sales = sum of weekly net sales; gray lines = 2020/21–2024/25.',fontsize=7.5,color=INK2)
fig.tight_layout(rect=(0,0.04,1,1))
fig.savefig(sys.argv[2],dpi=200)
