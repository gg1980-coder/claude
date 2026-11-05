"""Renewable-diesel margin calculator (RD_margins.xlsx): audit + what it adds for soybean oil.

Run: python3 rd.py  -> rd_results.json, rd_stats.xlsx, charts/rd_*.png
"""
import itertools
import json
import warnings
from math import factorial
from pathlib import Path

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import numpy as np
import openpyxl
import pandas as pd
import statsmodels.api as sm
from scipy import stats
from statsmodels.tsa.stattools import adfuller

from load import load_daily, load_rd, load_monthly_fundamentals, RD_XLSX

warnings.filterwarnings("ignore")
OUT = Path(__file__).parent
CH = OUT / "charts"
R = {}

d = load_daily()
rd = load_rd()
FEEDS = ["SOY", "CANOLA", "CORNOIL", "TALLOW", "UCO"]

# ---------------------------------------------------------------- 0. audit the sheet's margin formula
ws = openpyxl.load_workbook(RD_XLSX, data_only=True)["RD Margins"]
name_map = {"Soybean": "SOY", "Canola": "CANOLA", "Corn Oil": "CORNOIL", "Tallow": "TALLOW", "UCO": "UCO"}
lbs = {name_map[ws.cell(46, c).value]: float(ws.cell(47, c).value) for c in range(4, 9)}
cred = {name_map[ws.cell(23, c).value]: {"BTC": float(ws.cell(24, c).value), 2025: float(ws.cell(25, c).value),
                                          2026: float(ws.cell(26, c).value)} for c in range(4, 9)}
lcfs_f = {name_map[ws.cell(30, c).value]: {int(ws.cell(r, 3).value): float(ws.cell(r, c).value) for r in range(31, 44)}
          for c in range(4, 9)}


def components(f, df):
    yr = df.index.year
    price = df["P_UCO_OR_YG"] if f == "UCO" else df[f"P_{f}"]
    comp = pd.DataFrame(index=df.index)
    comp["diesel"] = df["HO"] / 100
    comp["feedstock"] = -price / 100 * lbs[f]
    comp["RIN"] = df["RIN_D4"] * np.where(df.index < "2026-01-01", 1.7, 1.6)
    comp["BTC_45Z"] = np.where(df.index < "2025-01-01", cred[f]["BTC"],
                               np.where(df.index < "2026-01-01", cred[f][2025], cred[f][2026]))
    comp["LCFS"] = df["LCFS"] * pd.Series(yr, index=df.index).map(lcfs_f[f]).values
    return comp


audit = {}
for f in FEEDS:
    c = components(f, rd)
    mine = c.sum(axis=1, min_count=c.shape[1])
    j = pd.concat([mine, rd[f"M_{f}"]], axis=1).dropna()
    audit[f] = {"max_abs_diff_$/gal": round(float((j.iloc[:, 0] - j.iloc[:, 1]).abs().max()), 4), "n": len(j)}
R["formula_audit"] = audit
R["assumptions"] = {"lbs_per_gal": lbs, "credits_$/gal": cred,
                    "RIN_equivalence": "1.7 before 2026, 1.6 from 2026", "LCFS_factor_2026": {f: lcfs_f[f][2026] for f in FEEDS}}

today = rd.index[rd["M_SOY"].notna()][-1]
dec = {f: components(f, rd).loc[today].round(3).to_dict() | {"margin": round(float(rd.loc[today, f"M_{f}"]), 3)} for f in FEEDS}
R["margin_decomposition_today"] = {"date": str(today.date()), **dec}
# sensitivity of soy margin to the lb/gal yield assumption (RD yields are typically above biodiesel's 7.5)
R["soy_yield_sensitivity_$/gal"] = {str(y): round(float(rd.loc[today, "M_SOY"] - (y - lbs["SOY"]) * rd.loc[today, "P_SOY"] / 100), 3)
                                    for y in (7.5, 8.0, 8.3, 8.6)}

# ---------------------------------------------------------------- 1. weekly panel
w_old = d.resample("W-FRI").last()
w_rd = rd.resample("W-FRI").last()
W = pd.concat([w_old[["BO", "HO", "PALM", "RAPESEED", "SOYBEANS", "RIN_D4", "OIL_SHARE_PROXY"]],
               w_rd[["LCFS", "P_CANOLA", "P_CORNOIL", "P_TALLOW", "P_UCO", "P_UCO_OR_YG", "P_YG"] + [f"M_{f}" for f in FEEDS] + ["M_BLEND"]]],
              axis=1).loc["2020-11-06":]
W["SOY_ADV"] = W["M_SOY"] - W[["M_TALLOW", "M_UCO", "M_CORNOIL", "M_CANOLA"]].mean(axis=1)   # soy margin vs other feeds
W["BO_TALLOW"] = W["BO"] - W["P_TALLOW"]
W["BO_UCO"] = W["BO"] - W["P_UCO_OR_YG"]
W["BO_CORNOIL"] = W["BO"] - W["P_CORNOIL"]
W["BO_CANOLA"] = W["BO"] - W["P_CANOLA"]
LEVEL = [c for c in W.columns if c.startswith("M_") or c in ("SOY_ADV", "BO_TALLOW", "BO_UCO", "BO_CORNOIL", "BO_CANOLA")]


def chg(c):
    return W[c].diff() if c in LEVEL else np.log(W[c]).diff()


bo = np.log(W["BO"]).diff()
rows = []
for c in ["LCFS", "P_CANOLA", "P_CORNOIL", "P_TALLOW", "P_UCO", "P_YG"] + [f"M_{f}" for f in FEEDS] + ["M_BLEND", "SOY_ADV", "BO_TALLOW", "BO_UCO"]:
    s = W[c].dropna(); r = chg(c).dropna()
    lvl = s if c in LEVEL else np.log(s)
    rows.append({"series": c, "start": str(s.index.min().date()), "weeks": len(s), "last": round(s.iloc[-1], 3),
                 "mean": round(s.mean(), 3), "min": round(s.min(), 3), "max": round(s.max(), 3),
                 "pctile_now_%": round(100 * (s <= s.iloc[-1]).mean()),
                 "weekly_unchanged_%": round(100 * (s.diff() == 0).mean(), 1),
                 "vol": round(r.std() * np.sqrt(52) * (100 if c not in LEVEL else 1), 2),
                 "vol_unit": "% ann." if c not in LEVEL else "$/gal or c/lb ann. sd",
                 "ADF_p_level": round(adfuller(lvl.dropna())[1], 3),
                 "corr_BO_same_week": round(bo.corr(r), 3),
                 "corr_X_lastweek_to_BO": round(bo.corr(r.shift(1)), 3),
                 "corr_BO_lastweek_to_X": round(r.corr(bo.shift(1)), 3)})
stats_tab = pd.DataFrame(rows).set_index("series")
R["series_stats"] = stats_tab.to_dict(orient="index")

# ---------------------------------------------------------------- 2. explanatory power (weekly returns)
X = pd.DataFrame({c: np.log(W[c]).diff() for c in ["BO", "HO", "PALM", "RAPESEED", "SOYBEANS", "RIN_D4", "LCFS",
                                                    "P_TALLOW", "P_UCO_OR_YG", "P_CORNOIL", "P_CANOLA"]})
BASE = ["HO", "PALM", "RAPESEED", "SOYBEANS"]
SETS = {"base": BASE, "base+RIN": BASE + ["RIN_D4"], "base+RIN+LCFS": BASE + ["RIN_D4", "LCFS"],
        "base+RIN+LCFS+fats(tallow,UCO/YG,corn oil)": BASE + ["RIN_D4", "LCFS", "P_TALLOW", "P_UCO_OR_YG", "P_CORNOIL"],
        # canola is largely quoted as BO futures + basis (half its daily moves leave canola-BO unchanged)
        # -> its R2 contribution is mechanical; shown only to make that visible
        "... + canola (MECHANICAL, do not use)": BASE + ["RIN_D4", "LCFS", "P_TALLOW", "P_UCO_OR_YG", "P_CORNOIL", "P_CANOLA"]}
ex = {}
for per, (a, b) in {"2021-2026": ("2021", "2026"), "2024-2026": ("2024", "2026")}.items():
    ex[per] = {}
    for k, cols in SETS.items():
        j = X.loc[a:b, ["BO"] + cols].dropna()
        m = sm.OLS(j["BO"], sm.add_constant(j[cols])).fit(cov_type="HAC", cov_kwds={"maxlags": 4})
        ex[per][k] = {"n": int(m.nobs), "R2": round(m.rsquared, 3), "adjR2": round(m.rsquared_adj, 3),
                      "t": m.tvalues.drop("const").round(2).to_dict()}
# monthly version: physical fat assessments are stale week to week (tallow unchanged in ~60% of weeks)
mo = pd.concat([d[["BO", "HO", "PALM", "RAPESEED", "SOYBEANS", "RIN_D4"]].resample("MS").mean(),
                rd[["LCFS", "P_TALLOW", "P_UCO_OR_YG", "P_CORNOIL"]].resample("MS").mean()], axis=1).loc["2020-12":"2026-10"]
mr = np.log(mo).diff().dropna()
exm = {}
for k, cols in list(SETS.items())[:4]:
    m = sm.OLS(mr["BO"], sm.add_constant(mr[cols])).fit(cov_type="HAC", cov_kwds={"maxlags": 3})
    exm[k] = {"n": int(m.nobs), "R2": round(m.rsquared, 3), "adjR2": round(m.rsquared_adj, 3), "t": m.tvalues.drop("const").round(2).to_dict()}
ex["monthly 2021-2026"] = exm
canola_basis = rd["P_CANOLA"] - rd["P_SOY"]
chg_days = rd["P_CANOLA"].diff() != 0
R["canola_basis_check"] = {"share_of_canola_moves_with_unchanged_basis": round(float((chg_days & (canola_basis.diff().abs() < 1e-9)).sum() / chg_days.sum()), 3)}
R["explanatory"] = ex

# ---------------------------------------------------------------- 3. does the margin calculator explain feedstock choice? (EIA, monthly)
mf = load_monthly_fundamentals()
mm = rd[["M_SOY", "M_TALLOW", "M_UCO", "M_CORNOIL", "M_CANOLA"]].resample("MS").mean()
mm["SOY_ADV"] = mm["M_SOY"] - mm[["M_TALLOW", "M_UCO", "M_CORNOIL", "M_CANOLA"]].mean(axis=1)
mm = mm.join(mf[["SOY_SHARE", "FS_SOYOIL", "FS_TOTAL"]], how="inner").loc["2021-01":"2026-08"].dropna()
excl = ~mm.index.year.isin([2023])            # 2023 EIA shares look interpolated (see REPORT.md)
fc = {}
for lag in (0, 1, 2, 3):
    a = mm["SOY_ADV"].shift(lag)
    fc[f"SOY_ADV_lead_{lag}m"] = {"corr_with_SOY_SHARE": round(a.corr(mm["SOY_SHARE"]), 3),
                                  "corr_ex2023": round(a[excl].corr(mm["SOY_SHARE"][excl]), 3),
                                  "corr_with_dSOY_SHARE_vs_dADV": round(a.diff().corr(mm["SOY_SHARE"].diff()), 3)}
fc["M_SOY_vs_FS_TOTAL_corr"] = round(mm["M_SOY"].corr(np.log(mm["FS_TOTAL"])), 3)
fc["n_months"] = len(mm)
R["feedstock_choice"] = fc

fig, ax1 = plt.subplots(figsize=(10, 4.2))
ax1.plot(mm.index, mm["SOY_ADV"], color="C3", label="Soy RD margin minus avg of other feeds ($/gal)")
ax1.axhline(0, color="grey", lw=.6)
ax2 = ax1.twinx(); ax2.plot(mm.index, mm["SOY_SHARE"] * 100, color="C0", label="Soyoil share of US biofuel feedstock (EIA, %)")
h1, l1 = ax1.get_legend_handles_labels(); h2, l2 = ax2.get_legend_handles_labels()
ax1.legend(h1 + h2, l1 + l2, frameon=False, loc="upper left", fontsize=8)
ax1.set_title("Does the calculator's soy advantage line up with what producers actually used?")
fig.tight_layout(); fig.savefig(CH / "rd_soy_advantage_vs_share.png", dpi=130); plt.close(fig)

# ---------------------------------------------------------------- 4. margin dynamics: who adjusts when margins are high?
dyn = {}
for c in ["M_SOY", "M_BLEND", "SOY_ADV"]:
    s = W[c].dropna()
    ar = sm.OLS(s.diff().dropna(), sm.add_constant(s.shift(1).dropna())).fit()
    dyn[c] = {"ADF_p": round(adfuller(s)[1], 3),
              "halflife_weeks": round(float(-np.log(2) / np.log(1 + ar.params.iloc[1])), 1) if ar.params.iloc[1] < 0 else None}
adj = {}
for h in (4, 13):
    for tgt in ["BO", "RIN_D4", "HO"]:
        fwd = np.log(W[tgt]).shift(-h) - np.log(W[tgt])
        j = pd.concat([W["M_SOY"], fwd], axis=1).dropna()
        m = sm.OLS(j.iloc[:, 1], sm.add_constant(j.iloc[:, 0])).fit(cov_type="HAC", cov_kwds={"maxlags": h})
        adj[f"{tgt}_next_{h}w_on_M_SOY"] = {"coef_per_$1": round(m.params.iloc[1], 4), "t_HAC": round(m.tvalues.iloc[1], 2),
                                           "R2_in_sample": round(m.rsquared, 3)}
dyn["who_adjusts_in_sample"] = adj
R["margin_dynamics"] = dyn

fig, ax = plt.subplots(figsize=(10, 4.5))
for f in FEEDS:
    ax.plot(W[f"M_{f}"], lw=1.2, label=f)
ax.axhline(1.25, color="k", ls="--", lw=.8, label="BI baseline 1.25")
ax.set_title("Gross RD margin by feedstock, $/gal (sheet's own calculation, weekly)")
ax.legend(frameon=False, ncol=6, fontsize=8)
fig.tight_layout(); fig.savefig(CH / "rd_margins.png", dpi=130); plt.close(fig)

dc = components("SOY", rd).resample("W-FRI").last().loc["2020-11-06":]
fig, ax = plt.subplots(figsize=(10, 4.5))
pos = dc.clip(lower=0); neg = dc.clip(upper=0)
cols_ = [f"C{i}" for i in range(dc.shape[1])]
ax.stackplot(dc.index, *[pos[c] for c in dc.columns], labels=dc.columns, colors=cols_, alpha=.8)
ax.stackplot(dc.index, *[neg[c] for c in dc.columns], colors=cols_, alpha=.8)
ax.plot(dc.sum(axis=1), color="k", lw=1.4, label="margin")
ax.set_title("Soy RD gross margin decomposition, $/gal"); ax.legend(frameon=False, ncol=6, fontsize=8)
fig.tight_layout(); fig.savefig(CH / "rd_soy_decomposition.png", dpi=130); plt.close(fig)

# ---------------------------------------------------------------- 5. pre-specified predictive test (same protocol as rin.py)
def z52(x):
    return (x - x.rolling(52).mean()) / x.rolling(52).std()


wf = d.resample("W-FRI").last()
F = pd.DataFrame({
    "M_SOY": W["M_SOY"], "z_M_SOY": z52(W["M_SOY"]), "z_M_BLEND": z52(W["M_BLEND"]),
    "z_SOY_ADV": z52(W["SOY_ADV"]), "z_BO_TALLOW": z52(W["BO_TALLOW"]), "z_BO_UCO": z52(W["BO_UCO"]),
    "z_OILSHARE": z52(wf["OIL_SHARE_PROXY"]).reindex(W.index),
    "z_BDMARGIN": z52(wf["BD_MARGIN_EX_CREDIT"]).reindex(W.index),
})
MODELS = {"hist_mean (benchmark)": [], "M_SOY level": ["M_SOY"], "z_M_SOY": ["z_M_SOY"], "z_M_BLEND": ["z_M_BLEND"],
          "z_SOY_ADV": ["z_SOY_ADV"], "z_BO_TALLOW": ["z_BO_TALLOW"], "z_BO_UCO": ["z_BO_UCO"],
          "z_OILSHARE + z_BDMARGIN (prior candidate)": ["z_OILSHARE", "z_BDMARGIN"],
          "z_OILSHARE + z_M_SOY": ["z_OILSHARE", "z_M_SOY"]}
lbo = np.log(W["BO"])
START, EVAL0 = pd.Timestamp("2021-11-05"), pd.Timestamp("2023-06-01")
pred = {}
for h, n in {"1w": 1, "4w": 4, "13w": 13}.items():
    y = lbo.shift(-n) - lbo
    rows_h = {}
    for name, cols in MODELS.items():
        P, A = [], []
        for t in W.index[W.index >= EVAL0]:
            if np.isnan(y.get(t, np.nan)):
                continue
            tr = W.index[(W.index >= START) & (W.index <= t)][:-n]
            yt = y.loc[tr]
            if cols:
                Xt = F.loc[tr, cols]; ok = Xt.notna().all(1) & yt.notna()
                if ok.sum() < 60 or F.loc[t, cols].isna().any():
                    continue
                b = np.linalg.lstsq(np.column_stack([np.ones(ok.sum()), Xt[ok].values]), yt[ok].values, rcond=None)[0]
                p = b[0] + F.loc[t, cols].values @ b[1:]
            else:
                p = yt.dropna().mean()
            P.append(p); A.append(y.loc[t])
        P, A = np.array(P), np.array(A)
        f_ = A ** 2 - ((A - P) ** 2 - P ** 2)
        cw = sm.OLS(f_, np.ones(len(f_))).fit(cov_type="HAC", cov_kwds={"maxlags": max(1, n - 1)})
        pn, an = P[::n], A[::n]; pnl = np.sign(pn) * an
        rows_h[name] = {"OOS_R2_%": round(100 * (1 - ((A - P) ** 2).sum() / (A ** 2).sum()), 2),
                        "CW_t": round(float(cw.tvalues[0]), 2), "hit_%": round(100 * (np.sign(P) == np.sign(A)).mean(), 1),
                        "sign_sharpe": round(float(pnl.mean() / pnl.std() * np.sqrt(52 / n)), 2), "n": len(A)}
    pred[h] = rows_h
R["predictive_prespecified"] = {"eval_window": f"{EVAL0.date()} to latest", "results": pred}

# ---------------------------------------------------------------- outputs
with pd.ExcelWriter(OUT / "rd_stats.xlsx") as xw:
    stats_tab.to_excel(xw, sheet_name="series_stats")
    pd.DataFrame(R["margin_decomposition_today"] | {}).drop("date", errors="ignore").to_excel(xw, sheet_name="margin_today")
    for per in ex:
        pd.DataFrame({k: {kk: vv for kk, vv in v.items() if kk != "t"} for k, v in ex[per].items()}).T.to_excel(xw, sheet_name=f"explanatory_{per}")
    for h in pred:
        pd.DataFrame(pred[h]).T.to_excel(xw, sheet_name=f"predictive_{h}")
json.dump(R, open(OUT / "rd_results.json", "w"), indent=1, default=str)
print(json.dumps(R, indent=1, default=str))
