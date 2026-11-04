"""What does the D4 RIN (RIN4Y0) add to the soybean oil framework?

Run: python3 rin.py  -> rin_results.json, rin_stats.xlsx, charts/rin_*.png

RIN history starts Nov-2020, so everything here is on 2021-2026 only (~300 weeks).
The predictive test uses a SMALL PRE-SPECIFIED set of models (no search), because a
6-year sample cannot support a data-mined basket.
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
import pandas as pd
import statsmodels.api as sm
from scipy import stats
from statsmodels.tsa.stattools import adfuller, coint, grangercausalitytests
from statsmodels.tsa.vector_ar.vecm import coint_johansen

from load import load_daily, load_avfo, to_weekly, to_monthly

warnings.filterwarnings("ignore")
OUT = Path(__file__).parent
CH = OUT / "charts"
R = {}

d = load_daily()
w = to_weekly(d).loc["2020-11-06":]
DRV = ["HO", "PALM", "RAPESEED", "SOYBEANS"]
lw = np.log(w[["BO", "RIN_D4"] + DRV])
rw = lw.diff().dropna()
PER = {"2021-2026": ("2021", "2026"), "2021-2023": ("2021", "2023"), "2024-2026": ("2024", "2026")}


def shapley(y, X):
    cols = list(X.columns); n = len(cols); r2 = {}
    for k in range(n + 1):
        for s in itertools.combinations(cols, k):
            r2[frozenset(s)] = sm.OLS(y, sm.add_constant(X[list(s)])).fit().rsquared if s else 0.0
    out = {}
    for c in cols:
        o = [x for x in cols if x != c]; t = 0
        for k in range(n):
            for s in itertools.combinations(o, k):
                t += factorial(k) * factorial(n - k - 1) / factorial(n) * (r2[frozenset(s) | {c}] - r2[frozenset(s)])
        out[c] = t
    return out


# ---------------------------------------------------------------- 1. RIN descriptive stats
s = w["RIN_D4"].dropna(); r = rw["RIN_D4"]
R["rin_stats"] = {
    "start": str(s.index.min().date()), "weeks": len(s), "last": s.iloc[-1], "mean": round(s.mean(), 3),
    "min": s.min(), "max": s.max(), "pctile_now_%": round(100 * (s <= s.iloc[-1]).mean()),
    "ann_vol_%": round(100 * r.std() * np.sqrt(52), 1), "skew": round(stats.skew(r), 2),
    "excess_kurt": round(stats.kurtosis(r), 2), "ADF_p_log_level": round(adfuller(np.log(s))[1], 3),
    "stale_days_%": round(100 * (d["RIN_D4"].dropna().diff() == 0).mean(), 1),
    "corr_BO_same_week": round(rw["BO"].corr(r), 3),
    "corr_RIN_lastweek_to_BO": round(rw["BO"].corr(r.shift(1)), 3),
    "corr_BO_lastweek_to_RIN": round(r.corr(rw["BO"].shift(1)), 3),
    "corr_by_period": {k: round(rw.loc[a:b, "BO"].corr(rw.loc[a:b, "RIN_D4"]), 3) for k, (a, b) in PER.items()},
    "Granger_RIN->BO_p": [round(x, 3) for x in [grangercausalitytests(rw[["BO", "RIN_D4"]], maxlag=4)[l][0]["ssr_ftest"][1] for l in range(1, 5)]],
    "Granger_BO->RIN_p": [round(x, 3) for x in [grangercausalitytests(rw[["RIN_D4", "BO"]], maxlag=4)[l][0]["ssr_ftest"][1] for l in range(1, 5)]],
}
rd = np.log(d[["BO", "RIN_D4"]]).diff().dropna()
R["rin_daily_leadlag"] = {f"lag{k}": round(rd["BO"].corr(rd["RIN_D4"].shift(k)), 3) for k in (-2, -1, 0, 1, 2)}

# ---------------------------------------------------------------- 2. explanatory power: does RIN fill the missing half?
ex = {}
for k, (a, b) in PER.items():
    sub = rw.loc[a:b]
    m0 = sm.OLS(sub["BO"], sm.add_constant(sub[DRV])).fit()
    m1 = sm.OLS(sub["BO"], sm.add_constant(sub[DRV + ["RIN_D4"]])).fit(cov_type="HAC", cov_kwds={"maxlags": 4})
    sh = shapley(sub["BO"], sub[DRV + ["RIN_D4"]])
    ex[k] = {"n": int(m1.nobs), "R2_without_RIN": round(m0.rsquared, 3), "R2_with_RIN": round(m1.rsquared, 3),
             "RIN_beta": round(m1.params["RIN_D4"], 3), "RIN_t_HAC": round(m1.tvalues["RIN_D4"], 2),
             "betas": m1.params.drop("const").round(3).to_dict(),
             "shapley_share": {c: round(v / m1.rsquared, 3) for c, v in sh.items()}}
R["explanatory"] = ex

# rolling 52w correlation
fig, ax = plt.subplots(figsize=(10, 4.2))
ax.plot(rw["BO"].rolling(52).corr(rw["RIN_D4"]), label="BO vs D4 RIN", lw=1.6, color="C3")
ax.plot(rw["BO"].rolling(52).corr(rw["HO"]), label="BO vs HO", lw=1.1, color="C0", alpha=.7)
ax.plot(rw["BO"].rolling(52).corr(rw["SOYBEANS"]), label="BO vs soybeans", lw=1.1, color="C2", alpha=.7)
ax.axhline(0, color="grey", lw=.6); ax.legend(frameon=False, ncol=3)
ax.set_title("Rolling 52-week correlation of weekly returns with BO")
fig.tight_layout(); fig.savefig(CH / "rin_rolling_corr.png", dpi=130); plt.close(fig)

# ---------------------------------------------------------------- 3. the big unexplained weeks: does RIN explain them?
m_base = sm.OLS(rw["BO"], sm.add_constant(rw[DRV])).fit()
m_rin = sm.OLS(rw["BO"], sm.add_constant(rw[DRV + ["RIN_D4"]])).fit()
top = m_base.resid.abs().sort_values(ascending=False).head(12).index
R["big_weeks"] = [{"week": str(t.date()), "BO_%": round(100 * rw.loc[t, "BO"], 1),
                   "resid_without_RIN_%": round(100 * m_base.resid[t], 1),
                   "RIN_%": round(100 * rw.loc[t, "RIN_D4"], 1),
                   "resid_with_RIN_%": round(100 * m_rin.resid[t], 1)} for t in sorted(top)]
R["resid_sd_%"] = {"without_RIN": round(100 * m_base.resid.std(), 2), "with_RIN": round(100 * m_rin.resid.std(), 2)}

# ---------------------------------------------------------------- 4. levels: does RIN fix the broken fair value?
lv = lw.dropna()
lev = {}
for lab, cols in {"without_RIN": DRV, "with_RIN": DRV + ["RIN_D4"]}.items():
    eg = sm.OLS(lv["BO"], sm.add_constant(lv[cols])).fit()
    res = eg.resid
    jo = coint_johansen(lv[["BO"] + cols], det_order=0, k_ar_diff=2)
    lev[lab] = {"elasticities": eg.params.drop("const").round(3).to_dict(), "R2": round(eg.rsquared, 3),
                "resid_ADF_p(approx)": round(adfuller(res)[1], 3), "resid_sd_%": round(100 * res.std(), 1),
                "resid_max_abs_%": round(100 * res.abs().max(), 1),
                "resid_2025-26_mean_%": round(100 * res.loc["2025":].mean(), 1),
                "johansen_rank_5pct": int(sum(jo.lr1 > jo.cvt[:, 1]))}
    lev[lab]["_res"] = res
R["levels_2021_2026"] = {k: {kk: vv for kk, vv in v.items() if kk != "_res"} for k, v in lev.items()}
R["pairwise_EG_BO_RIN_p"] = round(coint(lv["BO"], lv["RIN_D4"])[1], 3)

fig, ax = plt.subplots(figsize=(10, 4.2))
ax.plot(100 * lev["without_RIN"]["_res"], label="fair-value gap, HO+palm+rapeseed+beans", lw=1.2)
ax.plot(100 * lev["with_RIN"]["_res"], label="... + D4 RIN", lw=1.4, color="C3")
ax.axhline(0, color="grey", lw=.6); ax.legend(frameon=False); ax.set_ylabel("%")
ax.set_title("BO minus fitted level, 2021-2026: does adding RINs close the gap?")
fig.tight_layout(); fig.savefig(CH / "rin_levels_resid.png", dpi=130); plt.close(fig)

# ---------------------------------------------------------------- 5. biodiesel margin proxy
mg = w["BD_MARGIN_EX_CREDIT"].dropna()
ar = sm.OLS(mg.diff().dropna(), sm.add_constant(mg.shift(1).dropna())).fit()
R["margin"] = {"definition": "HO + 1.5 x D4 RIN x 100 - 7.5 x BO, c/gal, excludes $1 BTC (to 2024) and 45Z (2025+)",
               "last": round(mg.iloc[-1], 1), "mean": round(mg.mean(), 1),
               "mean_by_year": mg.groupby(mg.index.year).mean().round(1).to_dict(),
               "pctile_now_%": round(100 * (mg <= mg.iloc[-1]).mean()),
               "ADF_p": round(adfuller(mg)[1], 3),
               "halflife_weeks": round(float(-np.log(2) / np.log(1 + ar.params.iloc[1])), 1) if ar.params.iloc[1] < 0 else None}
fig, ax = plt.subplots(2, 1, figsize=(10, 6), sharex=True)
ax[0].plot(w["RIN_D4"], color="C3"); ax[0].set_title("D4 RIN, USD/RIN")
ax[1].plot(mg, color="C4", label="BD margin ex credits (HO + 1.5 RIN - 7.5 BO)")
ax[1].plot(-w["BOHO"], color="grey", lw=1, label="-BOHO (HO - 7.5 BO)")
ax[1].axhline(0, color="k", lw=.6); ax[1].legend(frameon=False); ax[1].set_ylabel("c/gal")
fig.tight_layout(); fig.savefig(CH / "rin_margin.png", dpi=130); plt.close(fig)

# ---------------------------------------------------------------- 6. pre-specified predictive test
def z52(x):
    return (x - x.rolling(52).mean()) / x.rolling(52).std()


wf = to_weekly(d)  # full history so z-scores of non-RIN features are warmed up
lbo = np.log(wf["BO"])
F = pd.DataFrame({
    "RIN_r1": np.log(wf["RIN_D4"]).diff(), "RIN_r4": np.log(wf["RIN_D4"]).diff(4),
    "z_RIN": z52(np.log(wf["RIN_D4"])), "z_MARGIN": z52(wf["BD_MARGIN_EX_CREDIT"]),
    "z_OILSHARE": z52(wf["OIL_SHARE_PROXY"]), "z_BOHO": z52(wf["BOHO"]),
})
MODELS = {
    "hist_mean (benchmark)": [],
    "RIN_r1": ["RIN_r1"], "RIN_r4": ["RIN_r4"], "z_RIN": ["z_RIN"], "z_MARGIN": ["z_MARGIN"],
    "z_OILSHARE + z_BOHO (prior best)": ["z_OILSHARE", "z_BOHO"],
    "z_OILSHARE + z_MARGIN": ["z_OILSHARE", "z_MARGIN"],
    "z_OILSHARE + z_BOHO + z_RIN": ["z_OILSHARE", "z_BOHO", "z_RIN"],
}
H = {"1w": 1, "4w": 4, "13w": 13}
START = pd.Timestamp("2021-11-05")     # all RIN features (incl. 52w z) defined from here
EVAL0 = pd.Timestamp("2023-06-01")     # first forecast date: >= ~80 training weeks for every model
pred_tab = {}
for h, n in H.items():
    y = lbo.shift(-n) - lbo
    rows = {}
    for name, cols in MODELS.items():
        P, A = [], []
        for t in wf.index[(wf.index >= EVAL0)]:
            if np.isnan(y.get(t, np.nan)):
                continue
            tr_idx = wf.index[(wf.index >= START) & (wf.index <= t)][: -n] if n > 0 else None
            yt = y.loc[tr_idx]
            if cols:
                Xt = F.loc[tr_idx, cols]; ok = Xt.notna().all(1) & yt.notna()
                if ok.sum() < 60 or F.loc[t, cols].isna().any():
                    continue
                b = np.linalg.lstsq(np.column_stack([np.ones(ok.sum()), Xt[ok].values]), yt[ok].values, rcond=None)[0]
                p = b[0] + F.loc[t, cols].values @ b[1:]
            else:
                p = yt.dropna().mean()
            P.append(p); A.append(y.loc[t])
        P, A = np.array(P), np.array(A)
        r2 = 1 - ((A - P) ** 2).sum() / (A ** 2).sum()
        f = A ** 2 - ((A - P) ** 2 - P ** 2)
        cw = sm.OLS(f, np.ones(len(f))).fit(cov_type="HAC", cov_kwds={"maxlags": max(1, n - 1)})
        pn, an = P[::n], A[::n]; pnl = np.sign(pn) * an
        rows[name] = {"OOS_R2_%": round(100 * r2, 2), "CW_t": round(float(cw.tvalues[0]), 2),
                      "hit_%": round(100 * (np.sign(P) == np.sign(A)).mean(), 1),
                      "sign_sharpe": round(float(pnl.mean() / pnl.std() * np.sqrt(52 / n)), 2), "n": len(A)}
    pred_tab[h] = rows
R["predictive_prespecified"] = {"eval_window": f"{EVAL0.date()} to latest", "train_from": str(START.date()),
                                "results": pred_tab}

# ---------------------------------------------------------------- 7. tallow (AVFOT006) and poultry fat (AVFOP018)
av = load_avfo()
full = pd.period_range(av.index.min(), av.index.max(), freq="M").to_timestamp()
av = av.reindex(full)
bom = np.log(to_monthly(d)["BO"]).reindex(full)
avr = {"months_present": int(av["TALLOW"].notna().sum()), "months_missing": int(av["TALLOW"].isna().sum())}
for c in av.columns:
    yoy = np.log(av[c]).diff(12)
    j = pd.concat([yoy, bom.diff(12), (bom.shift(-3) - bom), yoy.shift(2)], axis=1)
    j.columns = ["yoy", "bo_yoy", "bo_f3", "yoy_known"]
    avr[c] = {"corr_YoY_with_BO_YoY": round(j["yoy"].corr(j["bo_yoy"]), 3),
              "n_yoy_pairs": int(j[["yoy", "bo_yoy"]].dropna().shape[0]),
              "corr_known_YoY_to_BO_next3m": round(j["yoy_known"].corr(j["bo_f3"]), 3),
              "corr_level_with_BO": round(np.log(av[c]).corr(bom), 3)}
R["avfo"] = avr

# ---------------------------------------------------------------- outputs
with pd.ExcelWriter(OUT / "rin_stats.xlsx") as xw:
    pd.Series(R["rin_stats"]).astype(str).to_frame("D4 RIN").to_excel(xw, sheet_name="rin_stats")
    pd.DataFrame({k: {kk: vv for kk, vv in v.items() if not isinstance(vv, dict)} for k, v in ex.items()}).to_excel(xw, sheet_name="explanatory")
    pd.DataFrame({k: v["shapley_share"] for k, v in ex.items()}).to_excel(xw, sheet_name="shapley_share")
    pd.DataFrame(R["big_weeks"]).to_excel(xw, sheet_name="big_weeks", index=False)
    pd.DataFrame(R["levels_2021_2026"]).astype(str).to_excel(xw, sheet_name="levels")
    for h in H:
        pd.DataFrame(pred_tab[h]).T.to_excel(xw, sheet_name=f"predictive_{h}")
json.dump(R, open(OUT / "rin_results.json", "w"), indent=1, default=str)
print(json.dumps(R, indent=1, default=str))
