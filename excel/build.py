from openpyxl import Workbook
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.styles import Font
import sys
months=["Oct","Nov","Dec","Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep"]
years=[f"{y}/{str(y+1)[-2:]}" for y in range(2014,2027)]
crush='''166 169 173 170 154 171 158 156 152 156 145 135 1903
170 166 167 160 155 166 158 161 154 153 141 138 1890
176 171 169 171 151 161 150 158 148 156 152 145 1908
176 173 176 175 165 182 172 172 170 179 170 170 2079
184 178 184 183 163 179 172 165 158 179 178 162 2085
187 175 185 189 175 192 183 180 177 185 175 171 2173
197 191 193 197 164 188 170 174 162 166 168 164 2134
197 191 198 194 174 193 181 181 174 181 175 168 2207
197 190 187 191 177 198 187 189 175 185 169 175 2219
201 200 204 194 193 204 178 192 183 193 168 187 2297
216 210 218 213 190 207 202 204 197 205 198 205 2464
236 221 230 228 214 232 218 213 218 222 210 199 2640
239 233 240 238 217 241 227 228 221 228 224 199 2735'''
crush=[list(map(int,l.split())) for l in crush.splitlines()]
yv=[11.24,11.61,11.59,11.43,11.61,11.46,11.73,11.85,11.82,11.79,11.86,11.58,11.70]
yld=[[v]*12+[v] for v in yv]
sections=[("Crush",crush,"0"),("Yield",yld,"0.00")]
wb=Workbook(); wb.remove(wb.active)
for name,rows,fmt in sections:
    ws=wb.create_sheet(name)
    ws.append(["Season"]+months+["Total"])
    for y,r in zip(years,rows): ws.append([y]+r)
    for row in ws.iter_rows(min_row=2,min_col=2):
        for c in row: c.number_format=fmt
    for c in ws[1]: c.font=Font(bold=True)
    t=Table(displayName=name,ref=f"A1:N{len(rows)+1}")
    t.tableStyleInfo=TableStyleInfo(name="TableStyleMedium2",showRowStripes=True)
    ws.add_table(t); ws.column_dimensions["A"].width=10; ws.freeze_panes="B2"
ws=wb.create_sheet("Notes")
for l in ["Source: DNEXT screenshots, Production section.",
 "Crush: transcribed from a clear image. The Total column is the source's own total; it differs from the sum of months by 0-2 because the source rounds each month.",
 "Yield: constant across months in every season; read from a low-resolution image. Medium confidence.",
 "Carry In, Production, Imports: not included. The images are too low-resolution to read the digits reliably."]:
    ws.append([l])
ws.column_dimensions["A"].width=120
wb.save(sys.argv[1])
