# What the D4 RIN adds (rerun on `BO_relationships_v2.xlsx`)

> **Correction, 5 Nov 2026.** The `inputs 1` sheet pulls the `PX_CLOSE_1D` field, which is the *previous* day's close, so every price was dated one trading day late. I checked this against the `PX_LAST` series in `RD_margins.xlsx`, and `load.py` now re-dates the data. The numbers below are the corrected ones (Friday-to-Friday weeks).
> - **The conclusions are unchanged.**
> - **The specific "biggest unexplained weeks" are *not* robust to which weekday is sampled.** That table has been removed, and event effects should be tested on daily windows around each announcement.
> - Day-of-week sampling alone moves weekly R² by about ±0.05.

Reproduce with `python3 rin.py`. Tables are in `rin_stats.xlsx`, raw numbers in `rin_results.json`, charts in `charts/rin_*.png`.

**What changed in v2.** Every series from v1 is cell-for-cell identical, so the results in `REPORT.md` and `BASKET_REPORT.md` still stand. The new series are:
- **RIN4Y0**: the D4 RIN, daily from Nov 2020.
- **AVFOT006 (tallow) and AVFOP018 (poultry fat)**: monthly. These are most likely production volumes in thousand lb, so please confirm the units.

Soybean meal (SM1) is still not pulled.

---

## 1. Headline: RINs are the missing explanatory variable, not a predictor

| Weekly returns, BO explained by… | 2021–26 | 2021–23 | **2024–26** |
|---|---|---|---|
| HO + palm + rapeseed + soybeans | 0.49 | 0.53 | **0.44** |
| … + D4 RIN | 0.55 | 0.56 | **0.57** |
| RIN coefficient (t-stat, HAC) | 0.18 (5.2) | 0.15 (2.4) | 0.21 (5.3) |
| RIN share of explained R² (Shapley) | 19% | 12% | **32%, the largest single driver** |

On every weekday sampling (Mon–Fri), adding the RIN raises the 2024–26 R² by +0.13 to +0.21. On *monthly* averages the effect is stronger still: R² goes from 0.50 to 0.72 (see `RD_REPORT.md`).

- **In 2024–26 RINs recover almost all of the R² that was lost**, from 0.44 back to 0.57. The fall in R² noted in `REPORT.md` was missing policy information.
- **A 10% RIN move is associated with about a 2% soyoil move in the same week**, holding the other markets constant.

**Level fair value, 2021–26** (`charts/rin_levels_resid.png`):

| | R² | Deviation, std dev | Largest deviation | Average deviation 2025–26 |
|---|---|---|---|---|
| Without RIN | 0.75 | 9.9% | 25.2% | +5.5% |
| With RIN | **0.91** | **6.0%** | 16.4% | **−0.1%** |

- **The 2025–26 premium is explained by RINs.** Once RINs are included, the +20–25% gap from the earlier model mostly disappears.
- Fitted long-run elasticities: HO 0.26, RIN 0.24, soybeans 0.26, palm 0.13, rapeseed 0.06.
- **The evidence for a stable long-run relationship is mixed.** The Engle-Granger test on the deviation passes (approximate p < 0.01), and the BO–RIN pair passes on its own (p = 0.03). The Johansen test still finds no stable relationship. Six years is short for this test, so treat the fair value as a useful frame, not a proven anchor.
- **Right now soyoil is about 8% *below* the RIN-inclusive fair value.** That isn't a signal (see section 3); it only says the recent fall in soyoil went further than RINs and related markets did.

## 2. RIN statistics and lead-lag

| | |
|---|---|
| Last | $1.91, the 89th percentile since Nov 2020 (range $0.39–2.54) |
| Annual volatility | 49.5%, almost twice soyoil's 26.9% |
| Fat tails | excess kurtosis 4.2 |
| Stale days | **29%**. This is an OTC assessment, so its weekly moves are smoothed. |
| Correlation with BO, same week | 0.42 (0.37 in 2021–23, 0.50 in 2024–26) |
| RIN last week → BO this week | 0.00 |
| BO last week → RIN this week | 0.08 (Granger p = 0.06) |
| Daily: BO today → RIN tomorrow | **0.19** |
| Daily: RIN yesterday → BO today | 0.05 |

**Soyoil leads RINs, not the reverse.** Futures price policy news first, and the RIN assessment follows a day later. Partly this is the stale-assessment effect. Don't expect RIN moves to front-run soyoil.

**Unexplained weeks.** The earlier list of large unexplained weeks changed completely after the date fix, so I've removed it. One finding still stands: in some EPA announcement weeks, soyoil moved sharply while RINs barely moved. A daily event study around each announcement date is needed to show this properly. **An event calendar is still needed alongside RINs.**

## 3. Biodiesel margin proxy, and the predictive test

**Margin proxy (excluding tax credits) = HO + 1.5 × D4 RIN − 7.5 lb × BO, in c/gal** (`charts/rin_margin.png`).

- It is now **250 c/gal, the 97th percentile**. The 2026 average is 167, against 13 in 2025 and −1 in 2024.
- It is non-stationary (ADF p = 0.95). It excludes the $1/gal biodiesel tax credit (through 2024) and the 45Z credit (2025 onward), and the regime changed in 2025.
- **What it says:** RINs are currently paying far above what the diesel/soyoil spread alone requires. Soyoil is cheap against diesel (BOHO at the 16th percentile), while the policy credit is rich. The market is pricing in a lot of policy support; that is the risk to watch.

**Predictive test.** I tested 7 pre-chosen models, out of sample from Jun 2023 onward, trained from Nov 2021. There was no search, because a 6-year sample can't support one.

| Model | 1 week | 4 weeks | 13 weeks |
|---|---|---|---|
| Historical mean (benchmark) | −0.8% | −3.5% | −9.3% |
| RIN last-week move | −1.7% | −5.5% | −12.6% |
| RIN vs its 52-week average | −1.5% | −22.8% | −92.3% |
| Margin vs its 52-week average | −2.4% | −7.2% | −3.5% (CW t = 1.48) |
| Oil share + BOHO (prior best) | −5.9% | −18.3% | −32.1% |
| **Oil share + margin** | −2.1% | −6.2% | −7.4% (**CW t = 1.76**) |

OOS R², 2023–26, compared with a "no change" forecast.

- **Still nothing passes convincingly.** Oil share + margin at 13 weeks now has a Clark-West t-stat of 1.76, just over the 1.65 bar, but it was 1.62 before the date fix. A result that flips with one day of re-dating, out of 24 tests, is not a robust edge.
- **The prior best basket (oil share + BOHO) failed in this window.** Replacing BOHO with the margin makes it consistently better. That makes sense, because the margin is BOHO with the policy revenue added back in. **Keep oil share + margin as the candidate signal**, re-test it as more data accrues, and add a proper 45Z credit value to the margin.

## 4. Tallow and poultry fat (monthly)

- **37 of 130 months are missing.** When a month-end falls on a weekend, the trading-day pull drops that value. Re-pull with monthly periodicity (`per = M`) in BDH.
- Tallow: YoY correlation with BO YoY 0.06, predictive correlation 0.04. **No information.**
- Poultry fat: YoY correlation −0.27, predictive −0.16, level correlation −0.50. Weak, and the level correlation is likely a shared-trend effect.
- **The wrong measure for the question.** Animal-fat *production* is slow and price-inelastic, because it's a by-product of meat. What competes with soyoil for biofuel demand is the *price* of the alternative feedstocks: bleachable fancy tallow, choice white grease, poultry fat, distillers corn oil and UCO, in c/lb. **The soyoil-minus-tallow spread is the variable to get.**

## 5. Updated watch list, in order of evidence

1. **D4 RIN**: the largest single same-week driver in 2024–26 (32% of explained variance). Add **D6 RIN** and **LCFS** credit prices next.
2. **Soybeans**, and add **meal (SM1)** for the crush and oil share.
3. **Heating oil / ULSD**.
4. **Policy event calendar** (EPA RVO and small-refinery-exemption decisions, 45Z guidance): covers the shocks RINs miss.
5. Palm (convert to USD) and rapeseed: lower weight since 2024.
6. **Margin proxy + oil share**: the only near-predictive pair, so monitor it but don't trade on it yet.
7. **Fat prices** (tallow, UCO, corn oil spreads) instead of fat production volumes.
