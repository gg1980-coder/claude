"""Build dnext_production.xlsx from transcribed DNEXT screenshots.

Usage: python3 build.py dnext_production.xlsx
To add a section: add its 13x12 grid below and append it to SECTIONS.
"""
import sys
from openpyxl import Workbook
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.formatting.rule import CellIsRule
from openpyxl.comments import Comment
from data_lowres import carry_in, production, imports, biodiesel, exports

MONTHS = ["Oct", "Nov", "Dec", "Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep"]
YEARS = [f"{y}/{str(y + 1)[-2:]}" for y in range(2014, 2027)]
COLS = "BCDEFGHIJKLM"

# ---- Clear source image: Crush (last value = source total) ----
CRUSH = [list(map(int, l.split())) for l in """
166 169 173 170 154 171 158 156 152 156 145 135 1903
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
239 233 240 238 217 241 227 228 221 228 224 199 2735""".strip().splitlines()]

YIELD = [11.24, 11.61, 11.59, 11.43, 11.61, 11.46, 11.73, 11.85, 11.82, 11.79, 11.86, 11.58, 11.70]

# Carry Out(month) = Carry In(next month); Sep = next season's Oct.
# 2026/27 Sep ending stock read directly from the Carry Out image.
carry_out = [carry_in[i][1:] + [carry_in[i + 1][0] if i < 12 else 607] for i in range(13)]

# Cells where readings disagreed or the stock balance looked off: (season, month) -> note
Y = lambda s: YEARS.index(s)
M = lambda m: MONTHS.index(m)
FLAGS = {
    "Carry In": {
        ("2014/15", "Jun"): "2189 or 2180", ("2015/16", "Dec"): "1966 or 1955",
        ("2016/17", "Nov"): "1795 or 1755", ("2016/17", "Mar"): "2213 or 2218",
        ("2017/18", "Nov"): "1826 or 1876", ("2017/18", "Aug"): "2334 or 2384",
        ("2019/20", "May"): "2601 read twice, but April balance suggests ~2501",
        ("2021/22", "Nov"): "2386 or 2384", ("2021/22", "Jun"): "2388 or 2384",
        ("2021/22", "Jul"): "2318 or 2316", ("2022/23", "Jan"): "2356 or 2355",
        ("2023/24", "Nov"): "1502; Oct balance looks high, could be 1602",
        ("2023/24", "Sep"): "1620 or 1629",
    },
    "Production": {
        ("2014/15", "Dec"): "1936 or 1906", ("2015/16", "Dec"): "1929 or 1920",
        ("2015/16", "Jan"): "1855 or 1865", ("2016/17", "Mar"): "1865 (crush x yield rules out 1065)",
        ("2017/18", "Feb"): "1880 or 1800", ("2019/20", "Apr"): "2089 or 2099",
        ("2021/22", "Oct"): "2318 or 2316", ("2021/22", "Dec"): "2324 or 2274",
        ("2021/22", "Jun"): "2068 or 2088", ("2022/23", "Oct"): "2335 or 2338",
        ("2022/23", "Feb"): "2081 or 2061", ("2022/23", "Mar"): "2340; balance for March looks high",
        ("2023/24", "Nov"): "2325 or 2335", ("2023/24", "May"): "2267 or 2207",
        ("2023/24", "Jul"): "2306 or 2308", ("2024/25", "Jan"): "2527 or 2577",
        ("2024/25", "Feb"): "2245 or 2215", ("2025/26", "Nov"): "2532 or 2632",
        ("2025/26", "Dec"): "2667 or 2567", ("2025/26", "Jan"): "2627 or 2527",
    },
    "Imports": {
        **{("2014/15", m): "row sums to 256 vs ~264 total in image" for m in MONTHS},
        **{("2015/16", m): "row sums to 277 vs ~287 total in image" for m in MONTHS},
        **{("2020/21", m): "row sums to 316; image total unreadable (302/308?)" for m in MONTHS},
        ("2023/24", "May"): "89 (row total 622 rules out 69)",
    },
    "Biodiesel": {
        ("2014/15", "Jan"): "386 or 306", ("2016/17", "Jan"): "380 or 350",
        ("2016/17", "Feb"): "399 or 309", ("2017/18", "Mar"): "674 or 624",
        ("2018/19", "Dec"): "835 or 805", ("2018/19", "Mar"): "654 or 684",
        ("2018/19", "Apr"): "699 or 689",
        **{("2019/20", m): "672/714/642/712/664 vs 872/814/842/812/864 - unresolved"
           for m in ["May", "Jun", "Jul", "Aug", "Sep"]},
        ("2021/22", "Oct"): "837 or 832", ("2021/22", "Dec"): "907 or 902",
        ("2021/22", "Apr"): "842 or 847", ("2021/22", "Jun"): "830 or 836",
        ("2022/23", "Mar"): "906 or 986; March balance looks high",
        ("2023/24", "Feb"): "691 or 891; February balance looks high",
        ("2024/25", "Oct"): "1212 or 1232", ("2025/26", "Oct"): "1006 or 1206; October balance looks high",
    },
    "Exports": {
        ("2014/15", "May"): "73 or 75", ("2017/18", "May"): "431 or 481",
        ("2020/21", "Mar"): "157; March balance looks high", ("2025/26", "Oct"): "24 or 224",
    },
}
FLAGS["Carry Out"] = {}
for (s, m), note in FLAGS["Carry In"].items():  # carry-out cell for the month before
    i, j = Y(s), M(m)
    key = (s, MONTHS[j - 1]) if j > 0 else (YEARS[i - 1], "Sep")
    FLAGS["Carry Out"][key] = "mirror of Carry In: " + note

# name, rows, number format, total kind, confidence
SECTIONS = [
    ("Crush", [r[:12] for r in CRUSH], "0", "source", "High - clear image"),
    ("Yield", [[v] * 12 for v in YIELD], "0.00", "avg", "Medium - low-res image, constant per season"),
    ("Carry In", carry_in, "0", "first", "Low - low-res image, cross-checked with Carry Out"),
    ("Production", production, "0", "sum", "Low - low-res image, sanity-checked vs crush x yield"),
    ("Imports", imports, "0", "sum", "Low - low-res image"),
    ("Biodiesel", biodiesel, "0", "sum", "Low - low-res image"),
    ("Exports", exports, "0", "sum", "Low - low-res image"),
    ("Carry Out", carry_out, "0", "last", "Low - derived from Carry In (= next month's Carry In)"),
]

YELLOW = PatternFill("solid", fgColor="FFF2A8")
BOLD = Font(bold=True)


def tname(name):
    return name.replace(" ", "")


wb = Workbook()
wb.remove(wb.active)
for name, rows, fmt, total, conf in SECTIONS:
    ws = wb.create_sheet(name)
    ws.append(["Season"] + MONTHS + ["Total"])
    for i, (y, r) in enumerate(zip(YEARS, rows)):
        n = i + 2
        tot = {"source": CRUSH[i][12] if name == "Crush" else None,
               "avg": f"=AVERAGE(B{n}:M{n})", "sum": f"=SUM(B{n}:M{n})",
               "first": f"=B{n}", "last": f"=M{n}"}[total]
        ws.append([y] + r + [tot])
    for row in ws.iter_rows(min_row=2, min_col=2):
        for c in row:
            c.number_format = fmt
    for c in ws[1]:
        c.font = BOLD
    for (s, m), note in FLAGS.get(name, {}).items():
        c = ws[f"{COLS[M(m)]}{Y(s) + 2}"]
        c.fill = YELLOW
        c.comment = Comment(note, "transcription")
    t = Table(displayName=tname(name), ref=f"A1:N{len(rows) + 1}")
    t.tableStyleInfo = TableStyleInfo(name="TableStyleMedium2", showRowStripes=True)
    ws.add_table(t)
    ws["P1"], ws["P2"] = "Confidence", conf
    ws["P1"].font = BOLD
    if name == "Crush":
        ws["P3"] = "Total = source total (months are rounded, so it can differ from their sum by 0-2)"
    ws.column_dimensions["A"].width = 10
    ws.column_dimensions["P"].width = 50
    ws.freeze_panes = "B2"

# ---- Balance check: implied other domestic use = CI + Prod + Imp - Bio - Exp - CO ----
ws = wb.create_sheet("Balance Check")
ws.append(["Season"] + MONTHS + ["Total"])
for i, y in enumerate(YEARS):
    n = i + 2
    ws.append([y] + [f"='Carry In'!{c}{n}+Production!{c}{n}+Imports!{c}{n}"
                     f"-Biodiesel!{c}{n}-Exports!{c}{n}-'Carry Out'!{c}{n}" for c in COLS]
              + [f"=SUM(B{n}:M{n})"])
for c in ws[1]:
    c.font = BOLD
for row in ws.iter_rows(min_row=2, min_col=2):
    for c in row:
        c.number_format = "0"
rng = f"B2:M{len(YEARS) + 1}"
ws.conditional_formatting.add(rng, CellIsRule(operator="greaterThan", formula=["1400"],
                                              fill=PatternFill("solid", fgColor="F4B6B6")))
ws.conditional_formatting.add(rng, CellIsRule(operator="lessThan", formula=["950"],
                                              fill=PatternFill("solid", fgColor="F4B6B6")))
t = Table(displayName="BalanceCheck", ref=f"A1:N{len(YEARS) + 1}")
t.tableStyleInfo = TableStyleInfo(name="TableStyleLight9", showRowStripes=True)
ws.add_table(t)
ws["P1"] = "What this is"
ws["P1"].font = BOLD
ws["P2"] = ("Implied other domestic use = Carry In + Production + Imports - Biodiesel - Exports - Carry Out. "
            "It should be smooth (~1,000-1,400). Red cells are outside 950-1,400 and usually point to a "
            "misread value in that month. Correct the value on its sheet and this updates.")
ws["P2"].alignment = Alignment(wrap_text=True, vertical="top")
ws.column_dimensions["P"].width = 60
ws.row_dimensions[2].height = 90
ws.column_dimensions["A"].width = 10
ws.freeze_panes = "B2"

# ---- All Data: long format, formulas pointing at each sheet ----
ws = wb.create_sheet("All Data")
ws.append(["Section", "Season", "Month", "Value"])
for name, rows, fmt, *_ in SECTIONS:
    for i, y in enumerate(YEARS):
        for j, m in enumerate(MONTHS):
            ref = f"'{name}'!{COLS[j]}{i + 2}"
            ws.append([name, y, m, f"={ref}"])
            ws.cell(ws.max_row, 4).number_format = fmt
for c in ws[1]:
    c.font = BOLD
t = Table(displayName="AllData", ref=f"A1:D{ws.max_row}")
t.tableStyleInfo = TableStyleInfo(name="TableStyleMedium2", showRowStripes=True)
ws.add_table(t)
ws.column_dimensions["A"].width = 12
ws.freeze_panes = "A2"

# ---- Notes ----
ws = wb.create_sheet("Notes", 0)
notes = [
    ("DNEXT - Production section, monthly, 2014/15 to 2026/27", True),
    ("", False),
    ("Likely US soybean oil: crush in million bushels, yield in lbs/bushel, other lines in million lbs "
     "(inferred from the numbers, not stated in the screenshots).", False),
    ("", False),
    ("Confidence", True),
    ("Crush: high. Clear image; matches its own row totals within rounding.", False),
    ("Yield: medium. Low-res image, but each season is one repeated value.", False),
    ("Carry In / Carry Out: low-medium. Low-res, but the two tables must chain "
     "(Carry Out = next month's Carry In), which resolved most conflicts.", False),
    ("Production, Imports, Biodiesel, Exports: low. Images ~330x150 px; digits 0/6/8/9 and 2/3 are ambiguous.", False),
    ("", False),
    ("Yellow cells = specific doubt; hover for the alternative reading.", False),
    ("Balance Check sheet = automatic error detector (see its note). Red cells there are the first place to look.", False),
    ("Totals on low-res sheets are formulas (the image totals were unreadable).", False),
    ("All Data = every section in long format (Section | Season | Month | Value), linked by formula; "
     "use it for pivots and charts.", False),
    ("", False),
    ("To replace with exact numbers: re-take low-res screenshots at ~740 px wide, or export from DNEXT.", False),
]
for text, bold in notes:
    ws.append([text])
    if bold:
        ws.cell(ws.max_row, 1).font = BOLD
ws.column_dimensions["A"].width = 130

wb.save(sys.argv[1])
print("saved", sys.argv[1])
