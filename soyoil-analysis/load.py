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


SHIFT_PX_CLOSE_1D = True


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
    if SHIFT_PX_CLOSE_1D:
        # 'inputs 1' pulls field PX_CLOSE_1D = PREVIOUS day's close; re-date each value to the
        # prior row's date so prices sit on the day they were set (verified against PX_LAST
        # in RD_margins.xlsx: BO, HO, RIN match exactly after the shift)
        dates = df.index
        df = df.iloc[1:].set_axis(dates[:-1])
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


RD_XLSX = Path(__file__).parent / "data" / "RD_margins.xlsx"
# (date col, value col) in sheet 'RD Margins' (0-based); PX_LAST, calendar days with fill
RD_COLS = {"HO": (16, 17), "RIN_D4": (19, 20), "LCFS_DEC": (22, 23), "LCFS_WAVG": (24, 25),
           "P_SOY": (27, 28), "P_CANOLA": (32, 33), "P_CORNOIL": (36, 37), "P_TALLOW": (40, 41),
           "P_UCO": (44, 45), "P_YG": (46, 47),
           "M_SOY": (27, 29), "M_CANOLA": (32, 34), "M_CORNOIL": (36, 38), "M_TALLOW": (40, 42),
           "M_UCO": (44, 48), "M_BLEND": (27, 50)}


def load_rd() -> pd.DataFrame:
    """RD margin calculator series. Gross margins are the sheet's own cached results ($/gal)."""
    x = pd.read_excel(RD_XLSX, sheet_name="RD Margins", header=None)
    out = {}
    for k, (dc, vc) in RD_COLS.items():
        d = pd.to_datetime(x.iloc[4:, dc], errors="coerce")
        # some RIN cells (Dec-21..Jul-22) are date-formatted in Excel, e.g. 1.605 shows as
        # 1900-01-01 14:31 -> convert those back to the Excel serial number
        raw = x.iloc[4:, vc].map(lambda z: (pd.Timestamp(z) - pd.Timestamp("1899-12-31")).total_seconds() / 86400
                                 if isinstance(z, (pd.Timestamp, __import__("datetime").datetime)) else z)
        v = pd.to_numeric(raw, errors="coerce")
        s = pd.Series(v.values, index=d.values)
        s = s[s.index.notna()].dropna()
        out[k] = s[~s.index.duplicated()]
    df = pd.DataFrame(out).sort_index()
    df = df[df.index.dayofweek < 5]                       # drop weekend fill rows
    # LCFS as used by the sheet: weekly-average index before Dec-2022, Dec futures after
    df["LCFS"] = df["LCFS_WAVG"].where(df.index < "2022-12-01", df["LCFS_DEC"])
    df["P_UCO_OR_YG"] = df["P_UCO"].fillna(df["P_YG"])   # sheet falls back to yellow grease
    df.index.name = "date"
    return df
