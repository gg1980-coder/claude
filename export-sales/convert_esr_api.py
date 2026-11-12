"""Convert the ESR API workbook (Raw sheet, 'ESR Raw Weekly - Corn and Soy All Countries.xlsm') into the CSV layout
of the USDA ESR query export used by the build scripts. Soybeans only.
Usage: python convert_esr_api.py <api.xlsm> <old_query_export.csv or -> <out.csv>
- Stub (changeover) week rows are renumbered week 53 (they are the MY's final report, dated in September).
- Country names: taken from the old query export by country code where available (full names), else from EXTRA."""
import sys
import openpyxl
import pandas as pd

api, old, out = sys.argv[1:4]
wb = openpyxl.load_workbook(api, read_only=True, data_only=True)
it = wb['Raw'].iter_rows(values_only=True); hdr = next(it)
d = pd.DataFrame([r for r in it if r[0] == 'Soybeans'], columns=hdr)
EXTRA = {5330: 'INDIA', 7420: 'CAMEROON', 7790: 'KENYA', 6040: 'PAPUA NEW GUINEA', 5200: 'UNITED ARAB EMIRATES'}
names = dict(EXTRA)
if old != '-':
    o = pd.read_csv(old, thousands=','); o.columns = [c.strip() for c in o.columns]
    names.update(o.drop_duplicates('CountryCode2').set_index('CountryCode2').CountryName.str.strip().to_dict())
d['CountryName'] = d.CountryCode.map(names).fillna(d.Country)
d['WeekEnding'] = pd.to_datetime(d.WeekEnding)
d['WeekNumber'] = d.WeekNo.where(d.StubWeek != 1, 53)
d['ReportingMarketingYearName'] = d.MYStart.map(lambda y: f' MY -  Sep {y}/Aug {y+1}')
tot = d.groupby(['MarketYear', 'WeekEnding']).currentMYTotalCommitment.sum()
outdf = pd.DataFrame({
    'Commodity': 'SOYBEANS',
    'WeekendingDate': d.WeekEnding.dt.strftime('%m/%d/%Y'),
    'ReportingMarketingYearName': d.ReportingMarketingYearName,
    'WeekNumber': d.WeekNumber,
    'CountryName': d.CountryName,
    'CountryCode2': d.CountryCode,
    'GrossNewSales': d.grossNewSales, 'NetSales': d.currentMYNetSales, 'WeeklyExports': d.weeklyExports,
    'OutstandingSales': d.outstandingSales, 'AccumulatedExports': d.accumulatedExports, 'TotalCommitment': d.currentMYTotalCommitment,
    'NextMYNetSales': d.nextMYNetSales, 'NextMYOutstandingSales': d.nextMYOutstandingSales,
    'TotalTotalCommitment': [tot[(m, w)] for m, w in zip(d.MarketYear, d.WeekEnding)],
})
outdf.to_csv(out, index=False)
print('rows', len(outdf), 'MYs', sorted(d.MarketYear.unique())[0], '-', sorted(d.MarketYear.unique())[-1])
