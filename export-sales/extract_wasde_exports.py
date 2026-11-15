import pandas as pd, glob, sys
D = sys.argv[1]
DIRS = sys.argv[1:]
out = []
for f in sorted(f for dd in DIRS for f in glob.glob(dd + '/**/*.csv', recursive=True)):
    d = pd.read_csv(f, encoding='latin1', dtype=str)
    m = d[(d.Region == 'United States') & (d.Attribute == 'Exports') & d.Commodity.isin(['Corn', 'Oilseed, Soybean', 'Soybeans'])
          & d.Unit.str.contains('Bushel', na=False) & ~d.ReportTitle.str.contains('Reliability')]
    out.append(m[['ReportDate', 'ReleaseDate', 'Commodity', 'MarketYear', 'ProjEstFlag', 'Value']])
x = pd.concat(out)
x['Value'] = x.Value.astype(float); x['ReleaseDate'] = pd.to_datetime(x.ReleaseDate, format='mixed')
x['Commodity'] = x.Commodity.replace({'Oilseed, Soybean': 'Soybeans'})
x = x.drop_duplicates(['ReleaseDate', 'Commodity', 'MarketYear'])
x.to_csv(D + '/../us_exports_by_report.csv', index=False)
for c in ('Corn', 'Soybeans'):
    p = x[x.Commodity == c].pivot_table(index='MarketYear', columns='ReleaseDate', values='Value', aggfunc='first')
    p.columns = [d.strftime('%y-%m') for d in p.columns]
    print('==', c); print(p.to_string())
print('reports:', sorted(x.ReleaseDate.dt.strftime('%Y-%m').unique()))
