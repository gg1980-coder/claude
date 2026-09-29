"""Build index.html from the processor basis sheets.

Usage: drop new "Soybean_Processor_Basis_M-D-YY.pdf" files into sheets/, then run
    python3 build.py
It parses every sheet, writes data/history.json, and inlines the data plus the
US state outlines into index.html (a self-contained page).
"""
import json
import re
from datetime import date
from pathlib import Path

from pypdf import PdfReader

HERE = Path(__file__).parent
ROW = re.compile(r"^(.+?)\s{2,}(Not Posted|[+-]?\d+)\s*(S[A-Z])?\s+(Steady|[+-]\d+)\s*$")


def sheet_date(path):
    m = re.search(r"(\d{1,2})-(\d{1,2})-(\d{2})", path.stem)
    mo, d, y = map(int, m.groups())
    return date(2000 + y, mo, d).isoformat()


def parse(path):
    text = PdfReader(path).pages[0].extract_text(extraction_mode="layout")
    rows = {}
    for line in text.splitlines():
        m = ROW.match(line.strip())
        if not m:
            continue
        name, basis, month, chg = m.groups()
        rows[re.sub(r"\s+", " ", name)] = {
            "basis": None if basis == "Not Posted" else int(basis),
            "chg": 0 if chg == "Steady" else int(chg),
            "fut": month or "SX",
        }
    return rows


def main():
    plants = json.loads((HERE / "data/plants.json").read_text())
    sheets = sorted(((sheet_date(p), p) for p in (HERE / "sheets").glob("*.pdf")))
    dates = [d for d, _ in sheets]
    series = {k: {"basis": [], "chg": [], "fut": []} for k in plants}
    for d, p in sheets:
        rows = parse(p)
        unknown = set(rows) - set(plants)
        if unknown:
            raise SystemExit(f"{p.name}: add these plants to data/plants.json: {sorted(unknown)}")
        for k in plants:
            r = rows.get(k, {"basis": None, "chg": None, "fut": None})
            for f in ("basis", "chg", "fut"):
                series[k][f].append(r[f])
        print(f"{d}: {len(rows)} plants")

    history = {"dates": dates, "plants": [{**plants[k], **series[k]} for k in plants]}
    (HERE / "data/history.json").write_text(json.dumps(history, indent=1))

    topo = (HERE / "data/states-albers-10m.json").read_text()
    page = (HERE / "template.html").read_text()
    page = page.replace("__TOPO__", topo.strip()).replace("__HISTORY__", json.dumps(history, separators=(",", ":")))
    (HERE / "index.html").write_text(page)
    print(f"wrote index.html with {len(dates)} dates")


if __name__ == "__main__":
    main()
