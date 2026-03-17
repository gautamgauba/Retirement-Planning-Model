# Retirement Planning Model — Public Spec v1.4

*Sanitized for GitHub. No real financial values.*

## Overview

Interactive Excel retirement planning model for a couple (Spouse 1 ~early 60s, Spouse 2 ~mid 50s).
iPad and MacBook compatible. No VBA. Manual calculation mode (F9 / Calculate Now).

## Sheets (10 total)

|# |Sheet              |Purpose                                                             |
|--|-------------------|--------------------------------------------------------------------|
|1 |Benchmarks         |US + 2 HCOL city comparisons. Current income = blank user input only|
|2 |Accounts           |25+ accounts, ESPP, legacy assets (529/residence separate)          |
|3 |Parameters         |All inputs. Yellow=input, Green=from Accounts, Blue=calculated      |
|4 |30-Year Projection |Audited. Start→Withdraw→Grow→Tax→IRMAA→End. CF Check=0              |
|5 |Monte Carlo        |500 scenarios, 4%/10%, column-by-column (iPad safe)                 |
|6 |Scenario Comparison|6 strategies, Scenario 2 live-linked                                |
|7 |Recommendations    |Narrative + action timeline                                         |
|8 |Chart: Assets      |30-year stacked area: IRA / Roth / Taxable / Residence / 529        |
|9 |Chart: Income      |30-year bar: SS / Pension / DefComp / PT / Withdrawals              |
|10|Chart: MC          |Fan chart: 10th/25th/50th/75th/90th percentile paths                |

## Key Fixes in v1.3

- **FIX 1:** Growth on post-withdrawal balance only (not full starting balance)
- **FIX 2:** Taxable withdrawal = Spend + Tax − FixInc − IRA_W (covers full cash need)
- **FIX 3:** CF Check = End − Start − Growth + Tax + Spend − FixInc − Windfall = 0

## New in v1.4

### IRMAA Analysis

- 2026 brackets applied (based on T−2 MAGI)
- Joint filer thresholds: $218K / $274K / $342K / $410K
- Part B surcharge: $0 / $81.20×2 / $202.90×2 / $324.60×2 per year
- Part D surcharge: $0 / $14.50×2 / $37.50×2 / $60.40×2 per year
- MAGI computed from projection: IRA_W + FixInc + OrdGrowth + 85% SS
- **T−2 lookback:** IRMAA in year T uses MAGI from year T−2
- Annual IRMAA cost shown as separate column in Projection
- Lifetime IRMAA cost in Summary box

### IRMAA-Optimized IRA Withdrawal Strategy

- Parameters: choose optimization target
  - Option A: Fill 22% bracket ($211,400 taxable income) — ignores IRMAA
  - Option B: Stay below IRMAA Tier 1 ($218K MAGI) — lowest Medicare cost
  - Option C: Stay below IRMAA Tier 2 ($274K MAGI) — balanced
  - Option D: Stay below IRMAA Tier 3 ($342K MAGI) — aggressive depletion
- IRA withdrawal capped at whichever target is selected
- Roth conversion overlay: surplus bracket capacity → Roth conversion
- Trade-off visible: higher IRA withdrawal = lower RMDs but higher IRMAA

### Legacy Target — Fully Parameter-Driven

- All legacy references pull from `Parameters![J] Legacy Target`
- Scenario Comparison legacy per child = live formula
- Summary rows = live formula
- No hardcoded $6M anywhere in model

### Benchmarks

- Current income columns removed from benchmark tables
- Replaced with blank yellow input cells (user fills manually)
- Retirement income column retained (links to Projection Year 0)

## Projection Column Order (v1.4)

```
Start → Income → RMD/IRA_W → IRA_post/Roth_post → Growth →
Tax → IRMAA(T-2) → Tax_W → Tax_post → End → Legacy → CF=0
```

## Parameters Sections

```
[A] Current Situation
[B] Assets (from Accounts)
[B2] Legacy Assets (reference)
[C] Work Timeline
[D] Income Sources
[D1] Part-Time Income
[D2] Windfall Events
[E] Spending
[F] Tax Parameters (Federal MFJ 2025)
[F2] IRMAA Parameters (2026 brackets — NEW)
[G] Portfolio Income Mix
[H] Withdrawal Strategy (IRA bracket target + IRMAA optimization — UPDATED)
[I] Monte Carlo Parameters
[J] Legacy Target (drives all legacy calculations — UPDATED)
[K] RMD Factors
```

## IRMAA Bracket Table (Parameters [F2])

|MAGI Floor (Joint)|Part B Surcharge/person/yr|Part D Surcharge/person/yr|
|------------------|--------------------------|--------------------------|
|$0                |$0                        |$0                        |
|$218,001          |$81.20                    |$14.50                    |
|$274,001          |$202.90                   |$37.50                    |
|$342,001          |$324.60                   |$60.40                    |
|$410,001          |$446.30                   |$83.30                    |

## Withdrawal Strategy Options (Parameters [H])

|Option     |IRA Target      |IRMAA Tier|Annual IRMAA Cost (est.)     |
|-----------|----------------|----------|-----------------------------|
|A          |$211,400 taxable|Tier 2–3  |~$970–$1,540/yr              |
|B          |$218K MAGI      |Tier 0    |$0                           |
|C          |$274K MAGI      |Tier 1    |~$192/yr                     |
|D          |$342K MAGI      |Tier 2    |~$481/yr                     |
|Recommended|C or D          |Tier 1–2  |Small cost, large RMD savings|

## CF Check Formula (no change from v1.3)

`End - Start - Growth + Tax + IRMAA + Spend - FixInc - Windfall = 0`

## Build Constraints

- No VBA (iPad incompatible)
- Manual calculation mode
- No circular references
- Zero formula errors on delivery
- NORMINV() not NORM.INV()
- MC column-by-column (no nested chains)