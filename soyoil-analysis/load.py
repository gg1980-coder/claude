"""Load and clean the BO_relationships workbook into tidy daily / weekly / monthly frames."""
from pathlib import Path
import numpy as np
import pandas as pd

XLSX = Path(__file__).parent / "data" / "BO_relationships_v2.xlsx"   # v2 adds RIN4Y0, AVFOT006, AVFOP018

# column index of the value column in 'inputs 1' (date sits one column to the left)
PRICE_COLS = {
    "BO": 4,            # CBOT soybean oil front month, USc/lb
    "HO": 6,            # NYMEX heating oil (ULSD) front month, USc/gal
    "PALM": 8,          # Bursa Malaysia crude palm oil, MYR/t
    "RAPESEED": 10,     # Euronext rapeseed (seed), EUR/t
    "SOYBEANS": 12,     # CBOT soybeans front month, USc/bu
    "SBO_US_FOB": 14,   # US soyoil FOB, USD/t (physical)
    "SBO_BR_FOB": 16,   # Brazil soyoil FOB Paranagua, USD/t
    "SUNOIL_UA": 18,    # Ukraine sunflower oil, USD/t (sparse, starts 2019)
    "RAPEOIL_HAM": 20,  # Rapeseed oil FOB Hamburg, EUR/t
    "RIN_D4": 22,       # RIN4Y0: D4 (biomass-based diesel) RIN, USD/RIN, daily from Nov-2020
}
MONTHLY_COLS = {"TALLOW": 24, "POULTRY_FAT": 26}   # AVFOT006 / AVFOP018: monthly, likely production volume (000 lb) - confirm units


def load_daily() -> pd.DataFrame:
    raw = pd.read_excel(XLSX, sheet_name="inputs 1", header=None)
    out = {}
    for name, c in PRICE_COLS.items():
        d = pd.to_datetime(raw.iloc[3:, c - 1], errors="coerce")
        v = pd.to_numeric(raw.iloc[3:, c], errors="coerce")
        s = pd.Series(v.values, index=d.values).dropna()
        s = s[~s.index.isna()]
        out[name] = s[~s.index.duplicated()]
    df = pd.DataFrame(out).sort_index()
    df.index.name = "date"
    # Bloomberg fill='P' carries previous value on non-trading days -> treat repeats on
    # sparse physical series as missing for return calculations
    df["SUNOIL_UA"] = df["SUNOIL_UA"].where(df["SUNOIL_UA"].diff() != 0)
    # derived spreads / ratios
    df["BOHO"] = df["BO"] * 7.5 - df["HO"]           # c/gal: soyoil (7.5 lb/gal) minus ULSD
    df["OIL_SHARE_PROXY"] = 11.0 * df["BO"] / df["SOYBEANS"]  # oil value / bean value
    df["US_BR_FOB_SPREAD"] = df["SBO_US_FOB"] - df["SBO_BR_FOB"]
    # Biodiesel gross margin proxy, c/gal, EXCLUDING tax credits:
    #   revenue = ULSD + 1.5 D4 RINs per gal  ;  cost = 7.5 lb soyoil per gal
    df["BD_MARGIN_EX_CREDIT"] = df["HO"] + 150 * df["RIN_D4"] - 7.5 * df["BO"]
    return df


def load_avfo() -> pd.DataFrame:
    raw = pd.read_excel(XLSX, sheet_name="inputs 1", header=None)
    out = {}
    for name, c in MONTHLY_COLS.items():
        d = pd.to_datetime(raw.iloc[3:, c - 1], errors="coerce")
        v = pd.to_numeric(raw.iloc[3:, c], errors="coerce")
        s = pd.Series(v.values, index=d.values).dropna()
        s = s[s.diff() != 0]                       # keep only release points (sheet forward-fills)
        # month-end values that fall on a weekend first appear on the next trading day,
        # and late revisions land on the 1st: shift 4 days back before assigning the month
        out[name] = s.groupby((s.index - pd.Timedelta(days=4)).to_period("M")).last()
    m = pd.DataFrame(out)
    m.index = m.index.to_timestamp()
    return m


def load_monthly_fundamentals() -> pd.DataFrame:
    f = pd.read_excel(XLSX, sheet_name="Feedstock", header=None)
    rows = f.iloc[4:121]
    idx = pd.to_datetime(rows[1], errors="coerce")
    num = lambda c: pd.to_numeric(rows[c], errors="coerce").values
    fs = pd.DataFrame({
        "FS_TOTAL": num(23),        # total feedstock into BD+RD, m lbs
        "FS_VEGOIL": num(24),
        "FS_ANIMAL": num(25),
        "FS_SOYOIL": num(13),       # soybean oil into biofuels, m lbs
        "FS_CANOLA": num(9),
        "FS_CORNOIL": num(10),
        "SOY_SHARE": num(28),       # soyoil share of total feedstock
    }, index=idx)
    fs = fs[fs.index.notna()]

    b = pd.read_excel(XLSX, sheet_name="BS", header=None)
    rows = b.iloc[3:106]
    idx = pd.to_datetime(rows[1], errors="coerce")
    num = lambda c: pd.to_numeric(rows[c], errors="coerce").values
    bs = pd.DataFrame({
        "BD_PROD": num(3), "BD_IMP": num(4), "BD_DEMAND": num(6), "BD_STOCKS": num(7),
        "RD_PROD": num(10), "RD_IMP": num(11),
    }, index=idx)
    bs = bs[bs.index.notna()]
    m = fs.join(bs, how="outer")
    m["BIOFUEL_PROD"] = m["BD_PROD"] + m["RD_PROD"]
    return m


def to_weekly(df: pd.DataFrame) -> pd.DataFrame:
    return df.resample("W-FRI").last()


def to_monthly(df: pd.DataFrame) -> pd.DataFrame:
    return df.resample("MS").mean()   # monthly average, aligned to month start like EIA data
