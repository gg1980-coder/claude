# Input statistics and predictive basket search: soybean oil (BO1)

Reproduce with `python3 basket.py`. The full tables are in `input_stats.xlsx`, with one sheet per table, and the raw numbers in `basket_results.json`.

**To add RINs or any other series:** put a CSV with two columns, `date,value`, in `data/extra/` (for example `data/extra/RIN_D4.csv`) and rerun. The series is picked up automatically, both in the statistics and in the basket search.


> **Correction, 5 Nov 2026.** The `inputs 1` sheet pulls the `PX_CLOSE_1D` field, which is the *previous* day's close, so all prices were dated one trading day late. `load.py` now re-dates them; this was verified against the `PX_LAST` data in `RD_margins.xlsx`. After rerunning on the corrected dates:
> - The full-sample weekly R² is 0.50, not 0.53. The 2024–26 R² is unchanged at 0.44.
> - Every weekday sampling gives R² in the 0.45–0.54 range.
> - The out-of-sample predictive results still fail: R² −10.0% against a random walk. The basket search still finds no basket that beats "no change" in 2023–26 at any horizon (best: −4.0% at 13 weeks).
> - **The specific "largest unexplained weeks" list is not robust to which weekday is sampled. Treat the policy-week attributions as hypotheses to test on daily data, not as findings.**
> - The tables below are the original Thursday-dated run. The rerun numbers are in `results.json` and `basket_results.json`.

---

## 1. Individual input statistics (weekly, Dec 2015 – Nov 2026)

| Input | Vol (ann.) | Stale days | Now vs history (percentile) | Corr. with BO, same week | Last week's move → BO this week | Out-of-sample R², 4 weeks ahead |
|---|---|---|---|---|---|---|
| BO (target) | 26.9% | 1.4% | 83rd | – | −0.03 | −2.7% |
| Heating oil | 40.1% | 0.2% | 98th | 0.38 | −0.03 | −0.9% |
| Palm (MYR) | 29.1% | 8.2% | 78th | 0.50 | 0.01 | −1.2% |
| Rapeseed (EUR) | 23.6% | 3.9% | 83rd | 0.48 | −0.01 | −1.8% |
| Soybeans | 20.1% | 1.1% | 68th | 0.57 | −0.07 | −2.0% |
| US soyoil FOB | 29.0% | 9.2% | 81st | 0.67 | 0.02 | −2.5% |
| Brazil soyoil FOB | 30.4% | 5.5% | 80th | 0.71 | −0.01 | −2.3% |
| Sunflower oil (Ukraine) | 51.1% | – (2019+) | 68th | 0.15 | 0.02 | −5.9% |
| Rapeseed oil (Hamburg) | 25.8% | 6.9% | 90th | 0.16 | 0.01 | −1.1% |
| BOHO spread | – | – | **16th** | 0.54 | 0.00 | −3.4% |
| Oil-share proxy | – | – | **95th** | 0.69 | 0.02 | −2.2% |
| US − Brazil FOB spread | – | – | 77th | −0.12 | 0.02 | −1.5% |

How to read this:

- **Same-week correlation is high and next-week correlation is about zero, for every input.** They move *with* BO, never ahead of it. Every single-input forecast loses to "no change" at 1, 4 and 13 weeks (`single_feature_predictive` sheet).
- **Where things stand now** is more useful than any forecast:
  - The oil share of soybean value is at the 95th percentile, so oil is carrying the crush.
  - Heating oil is at the 98th percentile.
  - BOHO is only at the 16th percentile. Soyoil is cheap relative to diesel by historical standards, which is consistent with weak credit value or policy support. **This is exactly the gap that RIN data would explain.**
- **Distribution shape:**
  - Fat tails everywhere. Rapeseed and US FOB have excess kurtosis of 9–11, so normal-distribution risk bands understate tail moves.
  - US FOB and the US–Brazil spread show strong negative autocorrelation (−0.12 and −0.33): their weekly moves partly reverse. That is a sign of stale or noisy assessments, not information.

**Monthly EIA fundamentals** (`monthly_fundamentals` sheet). Here I lag each series 2 months to reflect EIA publication, then ask whether its year-on-year change tells you anything about BO over the next 3 months. Every correlation is between −0.25 and +0.14: **none is predictive**. Same-period correlations are higher once the corrupted 2023–24 data is excluded: total feedstock 0.66, animal fats 0.69, corn oil 0.74, soyoil share −0.64. They describe the same demand cycle, after the fact.

---

## 2. Basket search: method

- **21 candidate features**, all known at the time of each forecast:
  - last-week and last-4-week moves of BO, HO, palm, rapeseed, soybeans and the two FOB series;
  - relative-value scores, each measured as the distance from its 52-week average (z-scores): BO/HO, BO/palm, BO/rapeseed, oil share, BOHO and the US–Brazil FOB spread;
  - the rolling 3-year fair-value deviation.
- **Every basket of 1–3 features: 1,561 baskets × 3 horizons (1, 4 and 13 weeks).**
- **Forecasting:** linear regression on an expanding window. The model at each week only uses data whose outcome was already known.
- **Ranking and scoring:** baskets are ranked on 2019–22 only, then scored on the untouched 2023–26 period. I also report how many of *all* baskets succeeded in 2023–26, as the base rate you'd get by luck.
- **Scoring:** R² compared with a "no change" forecast. I also report the Clark-West test: a t-stat above about 1.65 means a genuine edge.

## 3. Results

| Horizon | Best basket picked on 2019–22 | Its 2019–22 R² | **Its 2023–26 R²** | 2023–26 direction hit rate | Share of all 1,561 baskets positive in 2023–26 | All 21 features together, 2023–26 |
|---|---|---|---|---|---|---|
| 1 week | Soybeans, last week | +0.9% | **−1.5%** | 49% | **0%** | −16.9% |
| 4 weeks | Palm + BO/palm + fair-value deviation | +2.5% | **−12.6%** | 46% | 5.9% | −29.4% |
| 13 weeks | Oil share + BOHO (both vs their 52-week average) | +0.4% | **−3.5%** | 56% | **0%** | −53.7% |

Benchmark: simply assuming the historical average trend continues scores −0.4% (1 week), −1.4% (4 weeks) and −3.1% (13 weeks) in 2023–26.

**Conclusion: no basket built from the current inputs predicts BO.**

- **The 2019–22 winners all collapse in 2023–26.** The 4-week winner relies on the fair-value deviation, which reverted toward fair value before 2023 and then trended for two years. This is the same structural break found in `REPORT.md`.
- **The 6% of baskets that look positive at 4 weeks are what luck produces.** The best of them in hindsight (HO + palm moves + BO/HO, +4.3%) could not have been chosen in advance.
- **Adding more features makes it worse.** The all-21-features model loses 17–54%. The signal-to-noise ratio is too low for a bigger model.
- **One combination is worth keeping on watch:** oil share + BOHO at 13 weeks.
  - In 2023–26 it called direction 56% of the time, with a sign-trading Sharpe of 0.5 and a positive Clark-West t-stat of 0.69 (not significant).
  - It roughly matches the historical-trend benchmark (−3.5% vs −3.1%), so it's an edge in direction, not in size.
  - It makes economic sense: crush share and the biofuel margin tend to revert once crushers and blenders respond.
  - It is the natural basket to strengthen with RINs.

**Explanatory basket (same week, for attributing moves, not forecasting):** HO + palm + rapeseed + soybeans gives R² 0.53. Adding the US and Brazil FOB series takes it to 0.69 (best by BIC). That gain comes from including the same commodity, not from better drivers (`explanatory_baskets` sheet).

## 4. What to do with the RIN data

1. Weekly D4 and D6 RIN prices, plus LCFS credits and the 45Z value if you can get them.
2. Drop each into `data/extra/` and rerun. The script builds last-week and last-4-week moves and a 52-week z-score for each, and adds them to the search.
3. **Success test:** a basket with RINs must beat "no change" in 2023–26 *and* have a Clark-West t-stat above about 1.65. The share-of-luck number should also stay low. Anything less is noise.
4. Also test RINs in the *explanatory* regression. If same-week R² rises from 0.44 toward 0.6 in 2024–26, the missing half is found, even if prediction stays hard.
5. **A more promising route than price momentum is policy-event timing.** EPA RVO and small-refinery-exemption dates, and the 45Z guidance calendar, can be modelled as scheduled risk events rather than forecast from prices.
