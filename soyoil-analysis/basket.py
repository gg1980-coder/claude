"""Per-input statistics + predictive basket search for CBOT soybean oil (BO1).

Run:  python3 basket.py
Outputs: input_stats.xlsx, basket_results.json, charts/basket_*.png

Adding new series (e.g. RINs): drop a CSV into data/extra/ with two columns
`date,value` (any frequency). The file stem becomes the series name and it is
automatically included in the stats tables and in the basket search.

Selection protocol (to avoid fooling ourselves with data-mining):
  * every forecast is genuinely out-of-sample: the model at week t is fitted only on
    rows whose target was already realised at t (expanding window);
  * baskets are RANKED on the selection period 2019-2022 only;
  * the top baskets are then SCORED on the untouched hold-out 2023-2026.
  Only the hold-out numbers should be believed.
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
from scipy import stats
from statsmodels.tsa.stattools import adfuller, grangercausalitytests

from load import load_daily, load_monthly_fundamentals, to_weekly, to_monthly

warnings.filterwarnings("ignore")
OUT = Path(__file__).parent
CH = OUT / "charts"
EXTRA = OUT / "data" / "extra"
SEL = ("2019-01-01", "2022-12-31")     # ranking period
HOLD = ("2023-01-01", "2026-12-31")    # hold-out period
HORIZONS = {"1w": 1, "4w": 4, "13w": 13}
MAX_K = int(__import__("os").environ.get("MAX_K", 3))  # max basket size searched exhaustively

UNITS = {"BO": "USc/lb", "HO": "USc/gal", "PALM": "MYR/t", "RAPESEED": "EUR/t", "SOYBEANS": "USc/bu",
         "SBO_US_FOB": "USD/t", "SBO_BR_FOB": "USD/t", "SUNOIL_UA": "USD/t", "RAPEOIL_HAM": "EUR/t",
         "BOHO": "USc/gal", "OIL_SHARE_PROXY": "ratio", "US_BR_FOB_SPREAD": "USD/t"}
SPREADS = ["BOHO", "OIL_SHARE_PROXY", "US_BR_FOB_SPREAD"]   # can be <=0 -> no logs


# ------------------------------------------------------------------ data
d = load_daily()
extra_names = []
if EXTRA.exists():
    for f in sorted(EXTRA.glob("*.csv")):
        e = pd.read_csv(f, parse_dates=[0]).iloc[:, :2]
        e.columns = ["date", "value"]
        s = e.set_index("date")["value"].astype(float).sort_index()
        d = d.join(s.rename(f.stem), how="outer")
        extra_names.append(f.stem)
        UNITS.setdefault(f.stem, "as supplied")
        SPREADS.append(f.stem)          # treat extras in levels/diffs (RINs can be ~0)

w = to_weekly(d)
PRICES = ["BO", "HO", "PALM", "RAPESEED", "SOYBEANS", "SBO_US_FOB", "SBO_BR_FOB", "SUNOIL_UA", "RAPEOIL_HAM"]
ALL = PRICES + SPREADS


def wret(c):
    """weekly change: log return for prices, simple difference for spreads/extras"""
    return w[c].diff() if c in SPREADS else np.log(w[c]).diff()


bo_r = np.log(w["BO"]).diff()
fwd = {h: np.log(w["BO"]).shift(-n) - np.log(w["BO"]) for h, n in HORIZONS.items()}


def oos_r2_single(x, y, h, start="2019-01-01"):
    """expanding-window OOS R2 (vs zero forecast) of y ~ a + b*x, honouring target lag h"""
    df = pd.concat([x.rename("x"), y.rename("y")], axis=1)
    idx = df.index
    preds, acts = [], []
    for i, t in enumerate(idx):
        if t < pd.Timestamp(start) or np.isnan(df["y"].iloc[i]) or np.isnan(df["x"].iloc[i]):
            continue
        tr = df.iloc[: max(0, i - h + 1)].dropna()
        if len(tr) < 104:
            continue
        b = np.polyfit(tr["x"].values, tr["y"].values, 1)
        preds.append(np.polyval(b, df["x"].iloc[i])); acts.append(df["y"].iloc[i])
    p, a = np.array(preds), np.array(acts)
    return 1 - ((a - p) ** 2).sum() / (a ** 2).sum() if len(a) else np.nan


# ------------------------------------------------------------------ 1. per-input statistics (weekly)
rows = []
for c in ALL:
    s = w[c].dropna()
    if len(s) < 60:
        continue
    r = wret(c).dropna()
    daily = d[c].dropna()
    lvl = s if c in SPREADS else np.log(s)
    j = pd.concat([bo_r.rename("BO"), r.rename("X")], axis=1).dropna()
    corr0 = j["BO"].corr(j["X"]) if c != "BO" else 1.0
    lead1 = bo_r.corr(r.shift(1))                     # X last week vs BO this week
    lag1 = bo_r.corr(r.shift(-1))                     # BO this week vs X next week (BO leads)
    gp = (grangercausalitytests(j[["BO", "X"]], maxlag=1)[1][0]["ssr_ftest"][1]
          if c != "BO" else np.nan)
    row = {
        "series": c, "units": UNITS.get(c, ""), "start": s.index.min().date(), "end": s.index.max().date(),
        "weeks": len(s), "stale_days_%": round(100 * (daily.diff() == 0).mean(), 1),
        "last": round(s.iloc[-1], 2), "mean": round(s.mean(), 2), "min": round(s.min(), 2), "max": round(s.max(), 2),
        "pctile_now_%": round(100 * (s <= s.iloc[-1]).mean(), 0),
        "ann_vol_%" if c not in SPREADS else "weekly_chg_sd": round(r.std() * (np.sqrt(52) * 100 if c not in SPREADS else 1), 2),
        "skew": round(stats.skew(r), 2), "excess_kurt": round(stats.kurtosis(r), 2),
        "ADF_p_level": round(adfuller(lvl, autolag="AIC")[1], 3),
        "AR1_weekly_chg": round(r.autocorr(1), 3),
        "corr_BO_same_week": round(corr0, 3),
        "R2_single_same_week": round(corr0 ** 2, 3),
        "corr_X_lastweek_to_BO": round(lead1, 3),
        "corr_BO_to_X_nextweek": round(lag1, 3),
        "Granger_X->BO_p(lag1)": round(gp, 3) if not np.isnan(gp) else None,
    }
    for h, n in HORIZONS.items():   # predictive: X's own past change over same horizon -> BO future
        xh = (w[c].diff(n) if c in SPREADS else np.log(w[c]).diff(n))
        row[f"OOS_R2_mom_{h}_%"] = round(100 * oos_r2_single(xh, fwd[h], n), 2)
    rows.append(row)
stats_w = pd.DataFrame(rows).set_index("series")

# monthly fundamentals stats
mf = load_monthly_fundamentals()
mp = to_monthly(d)
mm = mp[["BO"]].join(mf, how="inner").loc[:"2026-08"]
mm = mm[mm["FS_TOTAL"].notna()]
bo_m = np.log(mm["BO"])
bo_f3 = bo_m.shift(-3) - bo_m             # next 3 months BO return
frows = []
for c in ["FS_TOTAL", "FS_VEGOIL", "FS_ANIMAL", "FS_SOYOIL", "FS_CANOLA", "FS_CORNOIL", "SOY_SHARE",
          "BD_PROD", "RD_PROD", "BIOFUEL_PROD", "BD_STOCKS"]:
    s = mm[c].dropna()
    yoy = s.diff(12) if c == "SOY_SHARE" else np.log(s.replace(0, np.nan)).diff(12)
    known = yoy.shift(2)                    # EIA publication lag ~2 months
    frows.append({
        "series": c, "start": s.index.min().date(), "end": s.index.max().date(), "months": len(s),
        "last": round(s.iloc[-1], 2), "mean": round(s.mean(), 2),
        "yoy_sd": round(yoy.std(), 3),
        "corr_YoY_with_BO_YoY": round(yoy.corr(bo_m.diff(12)), 3),
        "corr_YoY_with_BO_YoY_ex2023-24": round((lambda j: j.iloc[:, 0].corr(j.iloc[:, 1]))(
            pd.concat([yoy, bo_m.diff(12)], axis=1).dropna().loc[lambda x: ~x.index.year.isin([2023, 2024])]), 3),
        "corr_known_YoY_to_BO_next3m": round(known.corr(bo_f3), 3),
        "n_pred_obs": int(pd.concat([known, bo_f3], axis=1).dropna().shape[0]),
    })
stats_m = pd.DataFrame(frows).set_index("series")

# ------------------------------------------------------------------ 2. candidate predictive features (all known at t)
F = {}
mom_src = ["BO", "HO", "PALM", "RAPESEED", "SOYBEANS", "SBO_US_FOB", "SBO_BR_FOB"] + extra_names
for c in mom_src:
    base = w[c] if c in SPREADS else np.log(w[c])
    F[f"{c}_r1"] = base.diff(1)
    F[f"{c}_r4"] = base.diff(4)


def z52(s):
    return (s - s.rolling(52).mean()) / s.rolling(52).std()


lb = np.log(w["BO"])
F["z_BO/HO"] = z52(lb - np.log(w["HO"]))
F["z_BO/PALM"] = z52(lb - np.log(w["PALM"]))
F["z_BO/RAPESEED"] = z52(lb - np.log(w["RAPESEED"]))
F["z_OILSHARE"] = z52(w["OIL_SHARE_PROXY"])
F["z_BOHO"] = z52(w["BOHO"])
F["z_US-BR_FOB"] = z52(w["US_BR_FOB_SPREAD"])
for c in extra_names:
    F[f"z_{c}"] = z52(w[c])

# rolling 3y related-markets fair-value residual (no look-ahead)
DRV = ["HO", "PALM", "RAPESEED", "SOYBEANS"]
lv = np.log(w[["BO"] + DRV]).dropna()
fv = {}
for i in range(156, len(lv)):
    win = lv.iloc[i - 156:i + 1]
    X = np.column_stack([np.ones(len(win)), win[DRV].values])
    b = np.linalg.lstsq(X, win["BO"].values, rcond=None)[0]
    fv[lv.index[i]] = win["BO"].values[-1] - X[-1] @ b
F["FV_resid"] = pd.Series(fv)
F = pd.DataFrame(F).reindex(w.index)
FEATS = list(F.columns)

# ------------------------------------------------------------------ 3. fast expanding-window OOS engine
idx = w.index
eval_mask = idx >= pd.Timestamp(SEL[0])
eval_pos = np.where(eval_mask)[0]
Fv = F.values
REFIT = 4  # refit every 4 weeks (forecast every week)


def oos_forecasts(cols, h):
    n = HORIZONS[h]
    y = fwd[h].values
    Xall = Fv[:, [FEATS.index(c) for c in cols]]
    pred = np.full(len(idx), np.nan)
    beta = None
    for k, i in enumerate(eval_pos):
        if np.isnan(y[i]) or np.isnan(Xall[i]).any():
            continue
        if beta is None or k % REFIT == 0:
            tr_end = i - n + 1                       # targets realised by t
            Xt, yt = Xall[:tr_end], y[:tr_end]
            ok = ~np.isnan(Xt).any(1) & ~np.isnan(yt)
            if ok.sum() < 104:
                continue
            Xd = np.column_stack([np.ones(ok.sum()), Xt[ok]])
            beta = np.linalg.lstsq(Xd, yt[ok], rcond=None)[0]
        pred[i] = beta[0] + Xall[i] @ beta[1:]
    return pred


def score(pred, h, period):
    y = fwd[h].values
    m = (idx >= pd.Timestamp(period[0])) & (idx <= pd.Timestamp(period[1])) & ~np.isnan(pred) & ~np.isnan(y)
    p, a = pred[m], y[m]
    if len(a) < 30:
        return None
    r2 = 1 - ((a - p) ** 2).sum() / (a ** 2).sum()
    hit = (np.sign(p) == np.sign(a)).mean()
    # non-overlapping long/short signal stats (sample every n weeks)
    n = HORIZONS[h]
    pn, an = p[::n], a[::n]
    pnl = np.sign(pn) * an
    sharpe = pnl.mean() / pnl.std() * np.sqrt(52 / n) if pnl.std() > 0 else np.nan
    # Clark-West test vs zero forecast (nested models), HAC t with n-1 lags
    f = a ** 2 - ((a - p) ** 2 - p ** 2)
    cw = sm.OLS(f, np.ones(len(f))).fit(cov_type="HAC", cov_kwds={"maxlags": max(1, n - 1)})
    return {"OOS_R2_%": round(100 * r2, 2), "hit_%": round(100 * hit, 1), "sharpe_sign": round(float(sharpe), 2),
            "CW_t": round(float(cw.tvalues[0]), 2), "n": int(len(a))}


results = {}
all_scores = []
for h in HORIZONS:
    combos = [c for k in range(1, MAX_K + 1) for c in itertools.combinations(FEATS, k)]
    for cols in combos:
        p = oos_forecasts(cols, h)
        s_sel = score(p, h, SEL)
        if s_sel is None:
            continue
        all_scores.append({"h": h, "basket": " + ".join(cols), "k": len(cols), **{f"sel_{k}": v for k, v in s_sel.items()},
                           "_pred": p})
    print("done", h, len(combos))

sc = pd.DataFrame(all_scores)
summary = {}
for h in HORIZONS:
    sub = sc[sc.h == h].sort_values("sel_OOS_R2_%", ascending=False)
    top = sub.head(10).copy()
    hold = [score(p, h, HOLD) for p in top["_pred"]]
    for k in ["OOS_R2_%", "hit_%", "sharpe_sign", "CW_t", "n"]:
        top[f"hold_{k}"] = [x[k] if x else None for x in hold]
    # also: how many of ALL baskets beat zero in hold-out (base rate of luck)
    hold_all = [score(p, h, HOLD) for p in sub["_pred"]]
    share_pos = np.mean([x["OOS_R2_%"] > 0 for x in hold_all if x])
    # what the hold-out winner would have been (pure hindsight, NOT a valid selection)
    hsub = sub.assign(hold_r2=[x["OOS_R2_%"] if x else np.nan for x in hold_all]).sort_values("hold_r2", ascending=False)
    # kitchen-sink benchmark: all features at once
    ks = score(oos_forecasts(tuple(FEATS), h), h, HOLD)
    summary[h] = {
        "n_baskets_tested": int(len(sub)),
        "share_of_all_baskets_with_positive_holdout_R2": round(float(share_pos), 3),
        "selection_winner_holdout": top.iloc[0][[c for c in top.columns if c.startswith("hold_")]].to_dict(),
        "hindsight_best_holdout_basket": hsub.iloc[0]["basket"], "hindsight_best_holdout_R2_%": hsub.iloc[0]["hold_r2"],
        "kitchen_sink_all_features_holdout": ks,
        "top10": top.drop(columns=["_pred", "h"]).to_dict(orient="records"),
    }
    results[h] = top

# ------------------------------------------------------------------ 4. single-feature predictive table
single = sc[sc.k == 1].copy()
single["feature"] = single["basket"]
sp = []
for _, r in single.iterrows():
    s_h = score(r["_pred"], r["h"], HOLD)
    sp.append({"feature": r["feature"], "h": r["h"], "sel_OOS_R2_%": r["sel_OOS_R2_%"],
               "hold_OOS_R2_%": s_h["OOS_R2_%"] if s_h else None, "hold_CW_t": s_h["CW_t"] if s_h else None,
               "hold_hit_%": s_h["hit_%"] if s_h else None})
single_tab = pd.DataFrame(sp).pivot(index="feature", columns="h", values=["sel_OOS_R2_%", "hold_OOS_R2_%", "hold_CW_t"])
single_tab.columns = [f"{a}_{b}" for a, b in single_tab.columns]

# ------------------------------------------------------------------ 5. explanatory (same-week) basket for attribution
rw = pd.DataFrame({c: wret(c) for c in ALL if c != "BO"}).join(bo_r.rename("BO"))
cand = ["HO", "PALM", "RAPESEED", "SOYBEANS", "SBO_BR_FOB", "SBO_US_FOB", "RAPEOIL_HAM"] + extra_names
expl = []
for k in range(1, len(cand) + 1):
    for cols in itertools.combinations(cand, k):
        j = rw[["BO", *cols]].loc["2016":].dropna()
        if len(j) < 200:
            continue
        m = sm.OLS(j["BO"], sm.add_constant(j[list(cols)])).fit()
        j2 = j.loc["2024":]
        m2 = sm.OLS(j2["BO"], sm.add_constant(j2[list(cols)])).fit()
        expl.append({"basket": " + ".join(cols), "k": k, "R2_full": round(m.rsquared, 3), "BIC": round(m.bic, 1),
                     "R2_2024-26": round(m2.rsquared, 3)})
expl = pd.DataFrame(expl).sort_values("BIC")

# ------------------------------------------------------------------ outputs
with pd.ExcelWriter(OUT / "input_stats.xlsx") as xw:
    stats_w.to_excel(xw, sheet_name="weekly_inputs")
    stats_m.to_excel(xw, sheet_name="monthly_fundamentals")
    single_tab.round(2).to_excel(xw, sheet_name="single_feature_predictive")
    for h in HORIZONS:
        results[h].drop(columns=["_pred", "h"]).to_excel(xw, sheet_name=f"top_baskets_{h}", index=False)
    expl.head(25).to_excel(xw, sheet_name="explanatory_baskets", index=False)

json.dump({"predictive": summary, "explanatory_top": expl.head(10).to_dict(orient="records"),
           "features": FEATS, "extras": extra_names},
          open(OUT / "basket_results.json", "w"), indent=1, default=str)

# chart: distribution of hold-out OOS R2 across all baskets, per horizon
fig, ax = plt.subplots(1, 3, figsize=(12, 3.8), sharey=False)
for a, h in zip(ax, HORIZONS):
    sub = sc[sc.h == h]
    vals = [score(p, h, HOLD)["OOS_R2_%"] for p in sub["_pred"] if score(p, h, HOLD)]
    a.hist(vals, bins=40, color="C0", alpha=.8)
    a.axvline(0, color="k", lw=1)
    w0 = summary[h]["selection_winner_holdout"]["hold_OOS_R2_%"]
    a.axvline(w0, color="C3", lw=1.5, label="selected basket")
    a.set_title(f"{h} ahead: hold-out OOS R² (%)\nall {len(vals)} baskets")
    a.legend(frameon=False, fontsize=8)
fig.tight_layout(); fig.savefig(CH / "basket_holdout_distribution.png", dpi=130); plt.close(fig)

pd.set_option("display.width", 250); pd.set_option("display.max_columns", 40)
print(stats_w.T.to_string())
print(stats_m.to_string())
print(single_tab.round(2).sort_values("hold_OOS_R2_%_4w", ascending=False).to_string())
print(json.dumps(summary, indent=1, default=str))
print(expl.head(10).to_string())
