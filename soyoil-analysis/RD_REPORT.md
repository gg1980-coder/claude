# Renewable-diesel margin calculator: audit and what it adds

Source: `data/RD_margins.xlsx`, sheet *RD Margins*. Daily `PX_LAST` from Jan 2020, with margins from Nov 2020 (when the RIN starts).
Reproduce with `python3 rd.py`. Outputs: `rd_results.json`, `rd_stats.xlsx`, `charts/rd_*.png`.

## 1. What the file contains

| Series | Ticker | Weeks unchanged | Note |
|---|---|---|---|
| ULSD | HO1 | – | same as before |
| D4 RIN | RIN4Y0 STON | – | 208 cells are date-formatted in Excel (see section 2) |
| CA LCFS credit | CLCFACPW (to Nov 2022), then LCICM DEC1 | 2% | |
| Canola oil | TTNSCOVU | 6% | **largely quoted as BO futures + basis** |
| Distillers corn oil | ETDICOIA | 30% | |
| Tallow | TTNSPBGU | **62%** | stale weekly assessment |
| UCO | GUCOUSSH | 26% | only from Aug 2023; the sheet uses yellow grease before that |
| Yellow grease | TTNSYGIL | **80%** | |
| Gross RD margins by feedstock, plus a blended margin | calculated | – | $/gal |

**Margin formula** (verified):

> margin = ULSD − feedstock price × lb/gal + D4 RIN × (1.7, or 1.6 from 2026) + BTC/45Z + LCFS credit × CI factor

My independent recalculation matches the sheet on **1,567 of 1,568 days**. The one exception is 2022-03-22.

## 2. Data problems found (and fixed in `load.py`)

1. **The original `inputs 1` sheet is dated one trading day late.** It pulls the `PX_CLOSE_1D` field, which is the previous day's close. After re-dating, BO matches this file's `PX_LAST` on 99.9% of days and HO on 100%.
   - The earlier reports now carry a correction note. Conclusions are unchanged, but the "largest unexplained weeks" list was not robust and has been withdrawn.
   - Going forward, use `PX_LAST` in the main sheet.
2. **208 D4 RIN cells (Dec 2021 – Jul 2022) are formatted as dates.** For example, 1.605 is stored as 1900-01-01 14:31. A normal read turns them into blanks; `load.py` converts them back.
3. **Canola oil is not an independent price.** On 51% of the days canola moves, the canola-minus-BO gap is unchanged: it is BO futures plus a basis. Adding it lifts weekly R² from 0.57 to 0.72, but **that is mechanical**, so I excluded it from the driver set.
4. **Tallow and yellow grease are stale weekly assessments** (unchanged in 62% and 80% of weeks). Weekly correlations understate their link, so I test them on monthly data as well.

## 3. Critique of the calculator's assumptions (they matter for the *level* of the margin)

**Today's soy RD gross margin** (5 Nov 2026):

| Component | $/gal |
|---|---|
| ULSD | +4.20 |
| Soyoil cost (62.8 c/lb × 7.5 lb) | −4.71 |
| RIN (1.96 × 1.6) | +3.14 |
| 45Z | +0.50 |
| LCFS | +0.22 |
| **Gross margin** | **3.35** |

That is the 97th percentile since 2020. The RIN is the second-largest revenue line, after ULSD.

Where the assumptions are questionable:

| Assumption | Issue | Effect |
|---|---|---|
| **7.5 lb soyoil per gallon** | That is the biodiesel conversion. RD hydrotreating yields are lower, and rules of thumb are roughly 8.0–8.5 lb/gal. | At 8.0 lb the margin is $3.04, at 8.3 it is $2.85, at 8.6 it is $2.66. |
| LCFS applied to every gallon | Only fuel sold in California earns it. | Overstates by up to the LCFS line ($0.22 soy, $0.62 UCO). |
| RIN equivalence 1.6 from 2026 (1.7 before) | Policy-dependent; worth confirming. | ±0.1 × RIN, about ±$0.20 |
| 45Z table (soy 0.2 in 2025, 0.5 in 2026; fats 0.6–0.7) | These are carbon-intensity-dependent estimates, not market prices. | Shifts feedstock rankings |
| No opex, hydrogen or freight | It's a *gross* margin; the sheet's own 1.25 "baseline" breakeven isn't subtracted. | The level isn't comparable to a cash margin. |
| Blend weights of 20% each | Arbitrary. EIA shows soy at about 40% and fats at about 30–35%. | Distorts the blended margin |

**Practical conclusion: use the margin's *changes* and *relative ranking across feedstocks*. Don't read the $/gal level literally.**

## 4. What the new series explain about BO

Same-week explanatory power:

| R² | Weekly 2021–26 | Weekly 2024–26 | **Monthly 2021–26** |
|---|---|---|---|
| HO + palm + rapeseed + soybeans | 0.49 | 0.44 | 0.50 |
| + D4 RIN | 0.55 | 0.57 | **0.72** |
| + LCFS | 0.55 | 0.57 | 0.73 |
| + tallow, UCO/yellow grease, corn oil | 0.57 | 0.57 | **0.78** |

- **RIN is confirmed as the key policy variable.** On monthly averages it lifts R² from 0.50 to 0.72 (t-stat 6.3). The weekly gain is smaller because the RIN is a stale OTC assessment: 29% of days are unchanged.
- **The LCFS adds nothing.** Its correlation with BO is −0.02 and its t-stat is below 1 in every specification. California credits matter for which fats go to California, not for the CBOT soyoil price.
- **Fats add a little at monthly frequency** (+0.06 R², mostly corn oil with a t-stat of 2.7). Week to week they *follow* BO: BO last week correlates with corn oil this week at 0.32, and with UCO at 0.28. The physical fat assessments lag futures; they don't lead.

## 5. Margins: who adjusts, and do margins explain feedstock choice?

- **The soy margin is persistent** (half-life about 16 weeks, not stationary). The soy *advantage* over other feedstocks reverts faster (half-life about 8 weeks, stationary, p = 0.004).
- **When the soy margin is high, it is the RIN that falls afterwards, not soyoil that rises** (in-sample, HAC):

| Next move, per $1/gal of soy margin | 4 weeks | 13 weeks |
|---|---|---|
| RIN | −2.9% (t = −2.2) | −8.8% (t = −2.2) |
| BO | −0.5% (t = −0.6) | −3.2% (t = −1.1) |
| HO | −1.6% (t = −1.1) | −3.7% (t = −0.8) |

  This fits the RIN being the balancing variable: rich margins attract production, which generates more RINs and pushes their price down. It is moderate in-sample evidence only, from about 300 overlapping weeks. It suggests the RIN is more forecastable than soyoil, which is worth testing properly if you trade RINs.
- **The calculator's soy advantage does not explain what producers actually used.**
  - Its correlation with the EIA soyoil share of feedstock is *negative*: −0.25 in the same month, and −0.41 when the soy advantage leads by 3 months.
  - Month-to-month changes are uncorrelated (≤ 0.17). There are 68 months in the sample (`charts/rd_soy_advantage_vs_share.png`).
  - Two explanations fit: (a) causality runs the other way, so more soy use bids BO up and squeezes the soy advantage; or (b) the 45Z and CI assumptions don't reflect real plant economics, contracts or capacity.
  - Either way, **the calculator in its current form is not a validated model of feedstock demand.**

## 6. Predictive test (pre-specified, out of sample from Jun 2023, same protocol as `rin.py`)

| Model | 1 week | 4 weeks | 13 weeks |
|---|---|---|---|
| Historical mean (benchmark) | −0.8% | −3.5% | −9.3% |
| Soy margin level | −3.1% | −7.6% | −25.2% |
| Soy margin vs its 52-week average | −4.1% | −9.1% | −13.6% |
| Blended margin vs its 52-week average | −1.7% | −9.3% | −32.3% |
| Soy advantage vs its 52-week average | −5.2% | −24.2% | −35.9% |
| BO − tallow vs its 52-week average | −3.5% | −21.4% | −45.7% |
| BO − UCO vs its 52-week average | −3.1% | −20.8% | −30.3% |
| Oil share + soy RD margin | −4.2% | −10.2% | −27.0% |
| Oil share + BD margin proxy (prior candidate) | −2.1% | −6.2% | −7.4% (CW t = 1.76) |

OOS R² compared with a "no change" forecast.

**None of the calculator's series predicts BO.** The best performer is still the simpler biodiesel margin proxy from `rin.py`, and its significance depends on a one-day date shift: Clark-West t was 1.62 before the date fix and 1.76 after. That isn't robust.

## 7. Where things stand now (5 Nov 2026)

| Indicator | Now | Percentile since 2020 |
|---|---|---|
| Soy RD gross margin (calculator) | $3.35/gal | 97th |
| Soy advantage vs other feedstocks | +$0.90/gal | 97th |
| BO − tallow | −11.7 c/lb | **9th**: soyoil is cheap against tallow |
| BO − UCO | −6.2 c/lb | 17th |
| D4 RIN | $1.96 | about 89th |
| LCFS | $90.5 | 67th |

**Reading:** soyoil is cheap against competing fats, and credit-rich margins make it the most profitable RD feedstock on these assumptions. Historically, a rich margin has been followed by **RINs falling** rather than soyoil rising. The risk to soyoil is therefore a RIN or policy correction, more than a feedstock-switching rally.

## 8. Updated watch list

1. **D4 RIN**, preferably the `PX_LAST` series. It is the dominant policy variable (monthly R² rises from 0.50 to 0.72).
2. Soybeans, and add **meal (SM1)**.
3. ULSD/HO.
4. **Soy margin and soy advantage**: use them as *state* indicators, and for RINs (the balancing variable). Don't use them as a BO forecast.
5. Fat prices at **monthly** frequency, mainly corn oil. Drop canola as a driver, since it is BO plus a basis.
6. LCFS: low priority for the CBOT soyoil price.
7. Fixes to make in the calculator: soy RD yield of 8.0–8.5 lb/gal, a California share for LCFS, opex, and blend weights taken from EIA shares.
