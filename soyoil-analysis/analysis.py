"""Statistical analysis of CBOT soybean oil (BO1) vs candidate drivers.

Run:  python3 analysis.py      -> prints results, writes charts/ and results.json
"""
import itertools
import json
import warnings
from pathlib import Path

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
import statsmodels.api as sm
from statsmodels.tsa.stattools import adfuller, kpss, coint, grangercausalitytests, ccf
from statsmodels.tsa.vector_ar.vecm import coint_johansen

from load import load_daily, load_monthly_fundamentals, to_weekly, to_monthly

warnings.filterwarnings("ignore")
OUT = Path(__file__).parent
CH = OUT / "charts"
CH.mkdir(exist_ok=True)
R = {}  # results collector

DRIVERS = ["HO", "PALM", "RAPESEED", "SOYBEANS"]          # liquid futures, clean daily data
PHYS = ["SBO_US_FOB", "SBO_BR_FOB", "RAPEOIL_HAM"]         # physical assessments
REGIMES = {"2016-2020 (pre-RD)": ("2016", "2020"),
           "2021-2023 (RD build-out)": ("2021", "2023"),
           "2024-2026 (45Z / RVO era)": ("2024", "2026")}

d = load_daily()
w = to_weekly(d)
mfund = load_monthly_fundamentals()
mp = to_monthly(d)

lw = np.log(w[["BO"] + DRIVERS + PHYS])
rw = lw.diff().dropna(how="all")
rw = rw.dropna(subset=["BO"] + DRIVERS)


def plt_save(fig, name):
    fig.tight_layout()
    fig.savefig(CH / name, dpi=130)
    plt.close(fig)


# ---------------------------------------------------------------- 1. stationarity
st = {}
for c in ["BO"] + DRIVERS + ["BOHO", "OIL_SHARE_PROXY"]:
    s = np.log(w[c]).dropna() if c not in ("BOHO",) else w[c].dropna()
    st[c] = {
        "ADF_p_level": round(adfuller(s, autolag="AIC")[1], 3),
        "KPSS_p_level": round(kpss(s, regression="c", nlags="auto")[1], 3),
        "ADF_p_diff": round(adfuller(s.diff().dropna(), autolag="AIC")[1], 4),
    }
R["stationarity"] = st

# ---------------------------------------------------------------- 2. correlations of weekly returns
R["corr_full"] = rw[["BO"] + DRIVERS + PHYS].corr()["BO"].round(3).drop("BO").to_dict()
R["corr_regime"] = {k: rw.loc[a:b, ["BO"] + DRIVERS + PHYS].corr()["BO"].round(3).drop("BO").to_dict()
                    for k, (a, b) in REGIMES.items()}

# rolling 52w correlation chart
fig, ax = plt.subplots(figsize=(10, 4.5))
for c in DRIVERS:
    ax.plot(rw["BO"].rolling(52).corr(rw[c]), label=c, lw=1.4)
ax.axhline(0, color="grey", lw=.6)
ax.set_title("Rolling 52-week correlation of weekly log returns with soybean oil (BO1)")
ax.legend(ncol=4, frameon=False)
plt_save(fig, "rolling_corr.png")


# ---------------------------------------------------------------- 3. multivariate return regression + Shapley R2
def ols_hac(y, X, lags=4):
    return sm.OLS(y, sm.add_constant(X)).fit(cov_type="HAC", cov_kwds={"maxlags": lags})


def shapley_r2(y, X):
    cols = list(X.columns)
    r2 = {}
    for k in range(len(cols) + 1):
        for sub in itertools.combinations(cols, k):
            r2[frozenset(sub)] = 0.0 if not sub else sm.OLS(y, sm.add_constant(X[list(sub)])).fit().rsquared
    n = len(cols)
    from math import factorial
    out = {}
    for c in cols:
        tot = 0.0
        others = [o for o in cols if o != c]
        for k in range(n):
            for sub in itertools.combinations(others, k):
                wgt = factorial(k) * factorial(n - k - 1) / factorial(n)
                tot += wgt * (r2[frozenset(sub) | {c}] - r2[frozenset(sub)])
        out[c] = tot
    return out


reg = {}
for label, (a, b) in {"Full 2016-2026": ("2016", "2026"), **REGIMES}.items():
    sub = rw.loc[a:b]
    m = ols_hac(sub["BO"], sub[DRIVERS])
    sh = shapley_r2(sub["BO"], sub[DRIVERS])
    reg[label] = {
        "n": int(m.nobs), "R2": round(m.rsquared, 3),
        "beta": m.params.drop("const").round(3).to_dict(),
        "t": m.tvalues.drop("const").round(2).to_dict(),
        "shapley_R2_share": {k: round(v / m.rsquared, 3) for k, v in sh.items()},
    }
R["return_regression"] = reg

# Chow-type test for structural break in betas at 2021-01 and 2024-01
def chow(y, X, split):
    Xc = sm.add_constant(X)
    full = sm.OLS(y, Xc).fit()
    m1 = sm.OLS(y[:split], Xc[:split]).fit()
    m2 = sm.OLS(y[split:], Xc[split:]).fit()
    k = Xc.shape[1]
    ssr_u = m1.ssr + m2.ssr
    F = ((full.ssr - ssr_u) / k) / (ssr_u / (len(y) - 2 * k))
    from scipy.stats import f
    return round(F, 2), float(f.sf(F, k, len(y) - 2 * k))

R["chow"] = {s: chow(rw["BO"], rw[DRIVERS], s) for s in ["2021-01-01", "2024-01-01"]}

# rolling 104w betas
roll_b = {}
for end in rw.index[104:]:
    sub = rw.loc[:end].iloc[-104:]
    roll_b[end] = sm.OLS(sub["BO"], sm.add_constant(sub[DRIVERS])).fit().params.drop("const")
roll_b = pd.DataFrame(roll_b).T
fig, ax = plt.subplots(figsize=(10, 4.5))
for c in DRIVERS:
    ax.plot(roll_b[c], label=c, lw=1.4)
ax.axhline(0, color="grey", lw=.6)
ax.set_title("Rolling 2-year multivariate betas: weekly BO return on driver returns")
ax.legend(ncol=4, frameon=False)
plt_save(fig, "rolling_betas.png")

# ---------------------------------------------------------------- 4. lead-lag
# Daily: palm (Kuala Lumpur) closes before CBOT opens; rapeseed (Paris) closes mid CBOT session.
rd = np.log(d[["BO"] + DRIVERS]).diff().dropna()
ll = {}
for c in DRIVERS:
    ll[c] = {f"lag{k}": round(rd["BO"].corr(rd[c].shift(k)), 3) for k in (-2, -1, 0, 1, 2)}
R["daily_leadlag_corr"] = ll  # lagK>0: driver K days earlier vs BO today

# Granger on weekly returns (does X help predict BO next week, and vice versa)
gr = {}
for c in DRIVERS:
    x_to_bo = grangercausalitytests(rw[["BO", c]].dropna(), maxlag=4)
    bo_to_x = grangercausalitytests(rw[[c, "BO"]].dropna(), maxlag=4)
    gr[c] = {"X->BO p (lags1..4)": [round(x_to_bo[l][0]["ssr_ftest"][1], 3) for l in range(1, 5)],
             "BO->X p (lags1..4)": [round(bo_to_x[l][0]["ssr_ftest"][1], 3) for l in range(1, 5)]}
R["granger_weekly"] = gr

# ---------------------------------------------------------------- 5. cointegration / long-run levels
lvl = np.log(w[["BO"] + DRIVERS]).dropna()
coin = {}
for label, (a, b) in {"Full 2016-2026": ("2016", "2026"), **REGIMES}.items():
    sub = lvl.loc[a:b]
    eg = sm.OLS(sub["BO"], sm.add_constant(sub[DRIVERS])).fit()
    resid = eg.resid
    adf_p = adfuller(resid, autolag="AIC")[1]
    # half-life of the residual (AR1)
    ar = sm.OLS(resid.diff().dropna(), sm.add_constant(resid.shift(1).dropna())).fit()
    hl = -np.log(2) / np.log(1 + ar.params.iloc[1]) if ar.params.iloc[1] < 0 else np.inf
    jo = coint_johansen(sub, det_order=0, k_ar_diff=2)
    n_rank = int(sum(jo.lr1 > jo.cvt[:, 1]))
    coin[label] = {"EG_elasticities": eg.params.drop("const").round(3).to_dict(),
                   "EG_R2_levels": round(eg.rsquared, 3),
                   "EG_resid_ADF_p(approx)": round(adf_p, 3),
                   "resid_halflife_weeks": round(float(hl), 1),
                   "johansen_rank_5pct": n_rank}
R["cointegration"] = coin

# pairwise Engle-Granger (proper MacKinnon critical values)
R["pairwise_EG_p"] = {c: round(coint(lvl["BO"], lvl[c])[1], 3) for c in DRIVERS}
R["pairwise_EG_p_by_regime"] = {k: {c: round(coint(lvl.loc[a:b, "BO"], lvl.loc[a:b, c])[1], 3) for c in DRIVERS}
                                for k, (a, b) in REGIMES.items()}

# chart: actual vs full-sample fitted level
eg_full = sm.OLS(lvl["BO"], sm.add_constant(lvl[DRIVERS])).fit()
fig, ax = plt.subplots(2, 1, figsize=(10, 6), sharex=True, gridspec_kw={"height_ratios": [2, 1]})
ax[0].plot(np.exp(lvl["BO"]), label="BO1 actual (c/lb)", lw=1.3)
ax[0].plot(np.exp(eg_full.fittedvalues), label="Fitted from HO, palm, rapeseed, soybeans", lw=1.1)
ax[0].legend(frameon=False); ax[0].set_title("Soybean oil vs a static 'related markets' fair value")
ax[1].plot(eg_full.resid * 100, color="C3", lw=1); ax[1].axhline(0, color="grey", lw=.6)
ax[1].set_ylabel("residual, %")
plt_save(fig, "levels_fit.png")

# ---------------------------------------------------------------- 6. spreads: BOHO and oil share
fig, ax = plt.subplots(2, 1, figsize=(10, 6), sharex=True)
ax[0].plot(w["BOHO"], lw=1.2); ax[0].set_title("BOHO spread: soyoil (7.5 lb/gal) minus heating oil, c/gal")
ax[0].axhline(0, color="grey", lw=.6)
ax[1].plot(w["OIL_SHARE_PROXY"] * 100, lw=1.2, color="C2")
ax[1].set_title("Oil value share proxy: 11 lb x BO / soybean price, %")
plt_save(fig, "spreads.png")

hl_sp = {}
for c in ["BOHO", "OIL_SHARE_PROXY"]:
    s = w[c].dropna()
    ar = sm.OLS(s.diff().dropna(), sm.add_constant(s.shift(1).dropna())).fit()
    hl_sp[c] = {"ADF_p": round(adfuller(s, autolag="AIC")[1], 3),
                "halflife_weeks": round(float(-np.log(2) / np.log(1 + ar.params.iloc[1])), 1)
                if ar.params.iloc[1] < 0 else None}
R["spreads"] = hl_sp

# ---------------------------------------------------------------- 7. monthly fundamentals
mm = mp[["BO", "HO", "PALM", "SOYBEANS", "BOHO", "OIL_SHARE_PROXY"]].join(mfund, how="inner")
mm = mm.loc[:"2026-08"]               # beyond Aug-26 the feedstock sheet holds projections, not actuals
mm = mm[mm["FS_TOTAL"].notna()]
yoy = np.log(mm[["BO", "HO", "PALM", "SOYBEANS", "FS_SOYOIL", "FS_TOTAL", "BIOFUEL_PROD", "RD_PROD", "BD_PROD"]]).diff(12)
yoy["SOY_SHARE"] = mm["SOY_SHARE"].diff(12)
yoy = yoy.dropna(subset=["BO", "FS_SOYOIL"])
R["monthly_n"] = int(len(yoy))
R["monthly_yoy_corr_with_BO"] = yoy.corr()["BO"].round(3).drop("BO").to_dict()

# lead/lag: price change vs fundamentals change at different lags (fund lag k months)
lagc = {}
for c in ["FS_SOYOIL", "SOY_SHARE", "FS_TOTAL", "BIOFUEL_PROD"]:
    lagc[c] = {f"fund_leads_{k}m": round(yoy["BO"].corr(yoy[c].shift(k)), 3) for k in (-6, -3, 0, 3, 6)}
R["monthly_yoy_leadlag"] = lagc  # positive k: fundamental k months earlier vs BO now; negative k: BO leads

# month-on-month change model: does adding fundamentals explain BO beyond related markets?
mom = np.log(mm[["BO", "HO", "PALM", "SOYBEANS", "FS_SOYOIL", "FS_TOTAL"]]).diff()
mom["SOY_SHARE"] = mm["SOY_SHARE"].diff()
mom = mom.dropna()
base = ols_hac(mom["BO"], mom[["HO", "PALM", "SOYBEANS"]], lags=3)
aug = ols_hac(mom["BO"], mom[["HO", "PALM", "SOYBEANS", "FS_SOYOIL", "SOY_SHARE"]], lags=3)
R["monthly_mom_models"] = {
    "markets_only_R2": round(base.rsquared, 3),
    "markets+fundamentals_R2": round(aug.rsquared, 3),
    "fund_coefs": aug.params[["FS_SOYOIL", "SOY_SHARE"]].round(3).to_dict(),
    "fund_t": aug.tvalues[["FS_SOYOIL", "SOY_SHARE"]].round(2).to_dict(),
    "n": int(aug.nobs),
}
# Data-integrity flags in the feedstock sheet
fsd = mfund.loc["2023-05":"2023-12", ["FS_TOTAL", "SOY_SHARE"]]
R["feedstock_2023_flat_share"] = fsd["SOY_SHARE"].round(4).tolist()
R["feedstock_2024_break"] = {str(k.date()): v for k, v in mfund.loc["2023-11":"2024-02", "FS_ANIMAL"].items()}

fig, ax1 = plt.subplots(figsize=(10, 4.5))
ax1.plot(mm.index, mm["BO"], color="C0", label="BO1 monthly avg (c/lb)")
ax2 = ax1.twinx()
ax2.bar(mm.index, mm["FS_SOYOIL"], width=20, alpha=.35, color="C1", label="Soyoil into biofuels (m lb)")
ax2.plot(mm.index, mm["FS_SOYOIL"].rolling(12).mean(), color="C1", lw=1.4, label="Soyoil into biofuels, 12m avg")
ax1.set_title("Soybean oil price vs soyoil used in US biofuel production (EIA)")
h1, l1 = ax1.get_legend_handles_labels(); h2, l2 = ax2.get_legend_handles_labels()
ax1.legend(h1 + h2, l1 + l2, frameon=False, loc="upper left")
plt_save(fig, "fundamentals.png")

# ---------------------------------------------------------------- 8. out-of-sample: do drivers PREDICT BO?
# Weekly, expanding window from 2019. Models: (a) random walk (zero return),
# (b) lagged driver returns, (c) error-correction term from a rolling 3y level regression.
lv = np.log(w[["BO"] + DRIVERS]).dropna()
dl = lv.diff()
rows = []
dates = lv.loc["2019":].index
for i, t in enumerate(dates[:-1]):
    hist = lv.loc[:t]
    win = hist.iloc[-156:]
    eg = sm.OLS(win["BO"], sm.add_constant(win[DRIVERS])).fit()
    ect = (hist["BO"] - eg.predict(sm.add_constant(hist[DRIVERS]))).iloc[-1]
    # train ECM / lag model on window
    dw = dl.loc[win.index]
    ect_w = (win["BO"] - eg.fittedvalues)
    X = pd.concat([ect_w.shift(1).rename("ECT"), dw[DRIVERS].shift(1), dw["BO"].shift(1).rename("BO_l1")], axis=1).dropna()
    y = dw["BO"].loc[X.index]
    mfit = sm.OLS(y, sm.add_constant(X)).fit()
    xnow = pd.DataFrame([[1, ect, *dl.loc[t, DRIVERS].values, dl.loc[t, "BO"]]], columns=["const"] + list(X.columns))
    pred = float(mfit.predict(xnow)[0])
    actual = dl.loc[dates[i + 1], "BO"]
    rows.append((dates[i + 1], actual, pred))
oos = pd.DataFrame(rows, columns=["date", "actual", "ecm"]).set_index("date")
mse_rw = (oos["actual"] ** 2).mean()
mse_ecm = ((oos["actual"] - oos["ecm"]) ** 2).mean()
R["oos_weekly"] = {"n": len(oos), "R2_oos_vs_randomwalk": round(1 - mse_ecm / mse_rw, 4),
                   "hit_rate_direction": round(float((np.sign(oos["ecm"]) == np.sign(oos["actual"])).mean()), 3)}

# monthly-horizon (4 week) version of the ECM signal alone
lv4 = lv.copy()
f4 = np.log(w["BO"]).shift(-4) - np.log(w["BO"])
ect_series = {}
for t in dates:
    win = lv.loc[:t].iloc[-156:]
    eg = sm.OLS(win["BO"], sm.add_constant(win[DRIVERS])).fit()
    ect_series[t] = win["BO"].iloc[-1] - eg.fittedvalues.iloc[-1]
ect_series = pd.Series(ect_series)
dfe = pd.concat([ect_series.rename("ECT"), f4.rename("fwd4")], axis=1).dropna()
m4 = sm.OLS(dfe["fwd4"], sm.add_constant(dfe["ECT"])).fit(cov_type="HAC", cov_kwds={"maxlags": 4})
R["ect_predicts_4w"] = {"coef": round(m4.params["ECT"], 3), "t_HAC": round(m4.tvalues["ECT"], 2),
                        "R2": round(m4.rsquared, 3), "n": int(m4.nobs),
                        "note": "rolling-3y residual (in-sample per date, no look-ahead); negative coef = mean reversion"}

# ---------------------------------------------------------------- 9. volatility & tail co-movement
vol = rw[["BO"] + DRIVERS].groupby(rw.index.year).std() * np.sqrt(52)
R["ann_vol_by_year"] = vol.round(3).to_dict()
q = rw["BO"].quantile([.05, .95])
R["tail_corr"] = {
    "BO_down_5pct_weeks": rw[rw["BO"] <= q[.05]][DRIVERS].mean().round(4).to_dict(),
    "BO_up_5pct_weeks": rw[rw["BO"] >= q[.95]][DRIVERS].mean().round(4).to_dict(),
}

# idiosyncratic share: residual from full-sample weekly regression
full_m = sm.OLS(rw["BO"], sm.add_constant(rw[DRIVERS])).fit()
res = full_m.resid
big = res.abs().sort_values(ascending=False).head(12)
R["largest_unexplained_weeks"] = {str(k.date()): round(float(res[k]) * 100, 2) for k in big.index}

fig, ax = plt.subplots(figsize=(10, 4))
ax.bar(res.index, res * 100, width=5, color="C3")
ax.set_title("Weekly BO return NOT explained by HO, palm, rapeseed, soybeans (%) - the policy/idiosyncratic part")
plt_save(fig, "idiosyncratic.png")

json.dump(R, open(OUT / "results.json", "w"), indent=1, default=str)
print(json.dumps(R, indent=1, default=str))
