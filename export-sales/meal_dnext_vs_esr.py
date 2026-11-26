"""US soybean meal: weekly export sales commitments (ESR, from the pace report's country table) vs Dnext monthly forecast by destination.
Usage: python3 meal_dnext_vs_esr.py <dnext_tradeflow.csv> <esr_country_table.csv> <out.xlsx>
Coverage = ESR commitments so far / Dnext Oct-Aug forecast; compared with the normal (10-yr) share of the year committed by this week."""
import sys, pandas as pd, numpy as np
tf, ec, out = sys.argv[1:4]
d = pd.read_csv(tf, low_memory=False); u = d[(d.Product == 'SBM') & (d.Reporter_MappedCode == 'USA')].copy(); u['P'] = u.Partner_Name.str.upper()
cy = u[u.CropYear == '2026/27']; fy = cy.groupby('P').Quantity.sum(); sep = cy[cy.Month == 9].groupby('P').Quantity.sum()
X = pd.DataFrame({'Dnext 2026/27 (Sep-Aug)': fy, 'Dnext Oct-Aug': fy.sub(sep, fill_value=0), 'Actual 2025/26': u[u.CropYear == '2025/26'].groupby('P').Quantity.sum(),
                  '5-yr avg': u[u.CropYear.isin(['2021/22', '2022/23', '2023/24', '2024/25', '2025/26'])].groupby(['P', 'CropYear']).Quantity.sum().unstack().fillna(0).mean(axis=1)}).fillna(0)
X.index = [{'BELGIUM': 'BELGIUM-LUXEMBOURG', 'MYANMAR': 'BURMA'}.get(i, i) for i in X.index]
E = pd.read_csv(ec, index_col=0).rename(columns={'ytd': 'ESR sold so far', 'ly_wk': 'ESR last season same wk', 'a10_wk': 'ESR 10-yr avg same wk', 'tgt10': 'ESR 10-yr avg full year', 'wk': 'ESR this week'})
J = E.join(X, how='outer').fillna(0)
EU = ['NETHERLANDS', 'UNITED KINGDOM', 'IRELAND', 'SPAIN', 'PORTUGAL', 'ITALY', 'GERMANY', 'FRANCE', 'BELGIUM-LUXEMBOURG', 'POLAND', 'DENMARK', 'NORWAY', 'GREECE', 'ROMANIA', 'SWITZERLAND', 'SWEDEN', 'FINLAND', 'AUSTRIA', 'BULGARIA', 'CROATIA', 'SLOVENIA', 'CYPRUS', 'LITHUANIA', 'LATVIA', 'ESTONIA', 'MALTA']
eu = J.loc[[n for n in EU if n in J.index]].sum(); eu.name = 'EUROPE'
J = pd.concat([J.drop(index=[n for n in EU if n in J.index]), eu.to_frame().T])
tot = J.sum(); tot.name = 'TOTAL'; J = pd.concat([J, tot.to_frame().T])
J['Coverage of Dnext'] = J['ESR sold so far'] / J['Dnext Oct-Aug'].replace(0, np.nan)
J['Normal coverage (10-yr)'] = J['ESR 10-yr avg same wk'] / J['ESR 10-yr avg full year'].replace(0, np.nan)
J['Full year at normal coverage (ESR basis)'] = J['ESR sold so far'] / J['Normal coverage (10-yr)']
J['Read'] = np.where(J['Coverage of Dnext'] > J['Normal coverage (10-yr)'] * 1.25, 'Dnext looks low', np.where(J['Coverage of Dnext'] < J['Normal coverage (10-yr)'] * 0.75, 'Dnext looks high', 'in line'))
J = J[(J['Dnext Oct-Aug'] >= 100) | (J['ESR sold so far'] >= 100) | J.index.isin(['EUROPE', 'TOTAL', 'UNKNOWN'])].sort_values('Dnext Oct-Aug', ascending=False)
cols = ['ESR sold so far', 'ESR last season same wk', 'ESR this week', 'Dnext Oct-Aug', 'Dnext 2026/27 (Sep-Aug)', 'Actual 2025/26', '5-yr avg', 'Coverage of Dnext', 'Normal coverage (10-yr)', 'Full year at normal coverage (ESR basis)', 'Read']
J[cols].round(3).to_excel(out, sheet_name='Meal ESR vs Dnext')
print(J[cols].round(2).to_string())
