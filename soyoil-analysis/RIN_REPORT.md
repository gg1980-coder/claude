# What the D4 RIN adds (rerun on `BO_relationships_v2.xlsx`)

Reproduce with `python3 rin.py`. Tables are in `rin_stats.xlsx`, raw numbers in `rin_results.json`, charts in `charts/rin_*.png`.

**What changed in v2.** Every series from v1 is cell-for-cell identical, so the results in `REPORT.md` and `BASKET_REPORT.md` still stand. The new series are:
- **RIN4Y0**: the D4 RIN, daily from Nov 2020.
- **AVFOT006 (tallow) and AVFOP018 (poultry fat)**: monthly. These are most likely production volumes in thousand lb, so please confirm the units.

Soybean meal (SM1) is still not pulled.

---

## 1. Headline: RINs are the missing explanatory variable, not a predictor

| Weekly returns, BO explained by… | 2021–26 | 2021–23 | **2024–26** |
|---|---|---|---|
| HO + palm + rapeseed + soybeans | 0.53 | 0.58 | **0.44** |
| … + D4 RIN | 0.60 | 0.62 | **0.59** |
| RIN coefficient (t-stat, HAC) | 0.19 (6.4) | 0.18 (3.3) | 0.20 (6.0) |
| RIN share of explained R² (Shapley) | 24% | 18% | **36%, the largest single driver** |

- **In 2024–26 RINs recover almost all of the R² that was lost**, from 0.44 back to 0.59. The fall in R² noted in `REPORT.md` was missing policy information.
- **A 10% RIN move is associated with about a 2% soyoil move in the same week**, holding the other markets constant.

**Level fair value, 2021–26** (`charts/rin_levels_resid.png`):

| | R² | Deviation, std dev | Largest deviation | Average deviation 2025–26 |
|---|---|---|---|---|
| Without RIN | 0.75 | 9.9% | 24.5% | +5.4% |
| With RIN | **0.91** | **6.0%** | 17.0% | **−0.1%** |

- **The 2025–26 premium is explained by RINs.** Once RINs are included, the +20–25% gap from the earlier model mostly disappears.
- Fitted long-run elasticities: HO 0.25, RIN 0.24, soybeans 0.24, palm 0.14, rapeseed 0.06.
- **The evidence for a stable long-run relationship is mixed.** The Engle-Granger test on the deviation passes (approximate p < 0.01), and the BO–RIN pair passes on its own (p = 0.03). The Johansen test still finds no stable relationship. Six years is short for this test, so treat the fair value as a useful frame, not a proven anchor.
- **Right now soyoil is about 8% *below* the RIN-inclusive fair value.** That isn't a signal (see section 3); it only says the recent fall in soyoil went further than RINs and related markets did.

## 2. RIN statistics and lead-lag

| | |
|---|---|
| Last | $1.91, the 89th percentile since Nov 2020 (range $0.39–2.54) |
| Annual volatility | 49.5%, almost twice soyoil's 26.9% |
| Fat tails | excess kurtosis 4.2 |
| Stale days | **29%**. This is an OTC assessment, so its weekly moves are smoothed. |
| Correlation with BO, same week | 0.51 (0.50 in 2021–23, 0.56 in 2024–26) |
| RIN last week → BO this week | 0.00 |
| BO last week → RIN this week | 0.08 (Granger p = 0.06) |
| Daily: BO today → RIN tomorrow | **0.19** |
| Daily: RIN yesterday → BO today | 0.05 |

**Soyoil leads RINs, not the reverse.** Futures price policy news first, and the RIN assessment follows a day later. Partly this is the stale-assessment effect. Don't expect RIN moves to front-run soyoil.

**Big unexplained weeks.** RINs explain part of several of them:

| Week | Unexplained before RINs | RIN move | Unexplained after RINs |
|---|---|---|---|
| Jun 2021 | −7.6% | −34% | −3.4% |
| Oct 2023 | −7.3% | −15% | −4.4% |
| Nov 2024 | −7.0% | −9% | −5.5% |
| Jun 2025 | +7.6% | +19% | +4.9% |
| Jan 2026 | +6.9% | +11% | +5.1% |

**They do not explain the Dec 2022 and Jun 2023 RVO weeks.** Soyoil fell 5–9% and RINs barely moved. In those weeks the policy shock went straight into soyoil futures, so **an event calendar of EPA and tax-credit announcements is still needed alongside RINs**.

## 3. Biodiesel margin proxy, and the predictive test

**Margin proxy (excluding tax credits) = HO + 1.5 × D4 RIN − 7.5 lb × BO, in c/gal** (`charts/rin_margin.png`).

- It is now **250 c/gal, the 97th percentile**. The 2026 average is 167, against 13 in 2025 and −1 in 2024.
- It is non-stationary (ADF p = 0.95). It excludes the $1/gal biodiesel tax credit (through 2024) and the 45Z credit (2025 onward), and the regime changed in 2025.
- **What it says:** RINs are currently paying far above what the diesel/soyoil spread alone requires. Soyoil is cheap against diesel (BOHO at the 16th percentile), while the policy credit is rich. The market is pricing in a lot of policy support; that is the risk to watch.

**Predictive test.** I tested 7 pre-chosen models, out of sample from Jun 2023 onward, trained from Nov 2021. There was no search, because a 6-year sample can't support one.

| Model | 1 week | 4 weeks | 13 weeks |
|---|---|---|---|
| Historical mean (benchmark) | −0.9% | −3.8% | −9.0% |
| RIN last-week move | −0.8% | −5.7% | −13.4% |
| RIN vs its 52-week average | −2.4% | −22.1% | −95.0% |
| Margin vs its 52-week average | −1.7% | −8.8% | −4.5% (CW t = 1.13) |
| Oil share + BOHO (prior best) | −3.8% | −20.5% | −34.7% |
| **Oil share + margin** | −1.5% | −8.7% (58% hit rate, sign Sharpe 0.66) | −9.8% (**CW t = 1.62**) |

OOS R², 2023–26, compared with a "no change" forecast.

- **Still nothing passes.** The best result is oil share + margin at 13 weeks, with a Clark-West t-stat of 1.62, just short of the 1.65 bar. Across 24 tests, one near-miss is what chance alone would produce.
- **The prior best basket (oil share + BOHO) failed in this window.** Replacing BOHO with the margin makes it consistently better. That makes sense, because the margin is BOHO with the policy revenue added back in. **Keep oil share + margin as the candidate signal**, re-test it as more data accrues, and add a proper 45Z credit value to the margin.

## 4. Tallow and poultry fat (monthly)

- **37 of 130 months are missing.** When a month-end falls on a weekend, the trading-day pull drops that value. Re-pull with monthly periodicity (`per = M`) in BDH.
- Tallow: YoY correlation with BO YoY 0.06, predictive correlation 0.04. **No information.**
- Poultry fat: YoY correlation −0.27, predictive −0.16, level correlation −0.50. Weak, and the level correlation is likely a shared-trend effect.
- **The wrong measure for the question.** Animal-fat *production* is slow and price-inelastic, because it's a by-product of meat. What competes with soyoil for biofuel demand is the *price* of the alternative feedstocks: bleachable fancy tallow, choice white grease, poultry fat, distillers corn oil and UCO, in c/lb. **The soyoil-minus-tallow spread is the variable to get.**

## 5. Updated watch list, in order of evidence

1. **D4 RIN**: the largest single same-week driver in 2024–26 (36% of explained variance). Add **D6 RIN** and **LCFS** credit prices next.
2. **Soybeans**, and add **meal (SM1)** for the crush and oil share.
3. **Heating oil / ULSD**.
4. **Policy event calendar** (EPA RVO and small-refinery-exemption decisions, 45Z guidance): covers the shocks RINs miss.
5. Palm (convert to USD) and rapeseed: lower weight since 2024.
6. **Margin proxy + oil share**: the only near-predictive pair, so monitor it but don't trade on it yet.
7. **Fat prices** (tallow, UCO, corn oil spreads) instead of fat production volumes.
