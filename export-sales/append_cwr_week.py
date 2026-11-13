"""Append one week of soybean data from the USDA 'CWR Country Commodity Summary' xlsx (country detail, thousands MT)
to the ESR CSV used by all build scripts.  Usage: python3 append_cwr_week.py <cwr.xlsx> <esr.csv> <MM/DD/YYYY>
Only rows for the given week-ending date are appended; country names are taken from the CSV by code where known."""
import sys, openpyxl, pandas as pd
src, csv, date = sys.argv[1:4]
ws = openpyxl.load_workbook(src, data_only=True).active
sec, rows = None, []
for r in ws.iter_rows(min_row=9, values_only=True):
    if isinstance(r[0], str) and ' - ' in r[0] and r[2] is None: sec = r[0].strip(); continue
    if sec == '801 - SOYBEANS' and r[2] == date and r[0] and str(r[0]).strip(): rows.append(r)
c = pd.read_csv(csv, dtype={'CountryCode2': str})
if (c.WeekendingDate == date).any(): sys.exit(f'{date} already in {csv}')
names = c.drop_duplicates('CountryCode2', keep='last').set_index('CountryCode2').CountryName
myname = c[c.WeekendingDate == c.WeekendingDate.iloc[-1]].ReportingMarketingYearName.iloc[0]
k = lambda v: int(round((v or 0) * 1000))
new = [dict(Commodity='SOYBEANS', WeekendingDate=date, ReportingMarketingYearName=myname, WeekNumber=int(r[3]),
            CountryName=names.get(str(r[0]).strip(), str(r[1]).strip()), CountryCode2=str(r[0]).strip(),
            GrossNewSales=k(r[7]), NetSales=k(r[12]), WeeklyExports=k(r[14]), OutstandingSales=k(r[13]),
            AccumulatedExports=k(r[16]), TotalCommitment=k(r[17]), NextMYNetSales=k(r[22]), NextMYOutstandingSales=k(r[21]))
       for r in rows if (r[17] or r[12] or r[14])]
n = pd.DataFrame(new); n['TotalTotalCommitment'] = n.TotalCommitment.sum()
pd.concat([c, n[c.columns]]).to_csv(csv, index=False)
print(f'appended {len(n)} rows, week {n.WeekNumber.iloc[0]}, total commitment {n.TotalCommitment.sum()/1e6:.3f} Mt')
