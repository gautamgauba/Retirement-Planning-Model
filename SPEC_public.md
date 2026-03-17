# Retirement Planning Model — Product Specification

**Version:** 1.2
**Date:** March 2026
**Status:** Approved for build
**Note:** All financial values are placeholders. See VALUES.md (private, stored locally only).

-----

## Change Log

|Version|Date      |Change                                                                                                              |
|-------|----------|--------------------------------------------------------------------------------------------------------------------|
|1.0    |March 2026|Initial spec — full model build                                                                                     |
|1.1    |March 2026|Accounts sheet; Target Yield; ESPP; Spouse 1/2 terminology; Work Event 1/2; Part-time income; Windfall events       |
|1.2    |March 2026|Added Benchmarks sheet (Sheet 1); sheet order updated; current + retirement position indicators; Boston HCOL overlay|

-----

## 1. Overview

### What to Build

Interactive Excel retirement planning model with Monte Carlo simulation. Works on **iPad (Excel iOS) and MacBook (Excel for Mac)**. No VBA — iPad compatible throughout.

### Spouse 1 (Primary User)

Professional nearing retirement. Planning retirement at age [RETIRE_AGE_A], [RETIRE_AGE_B], or [RETIRE_AGE_C] depending on Work Event Outcome expected end of March 2026. Understands financial concepts deeply. Wants to audit all formulas.

### Spouse 2 (Key Stakeholder)

Requires Monte Carlo simulation for confidence and approval. Uses iPad. Needs clear probability of success — not formula detail.

### Decision Being Modeled

Retire at [RETIRE_AGE_A] (Work Event 2) or [RETIRE_AGE_B]-[RETIRE_AGE_C] (Work Event 1). Optimize Social Security claiming, tax strategy, and withdrawal sequencing.

### Critical Constraints

- NO VBA — must work on iPad Excel
- Monte Carlo via NORMINV(RAND(),…) — no Data Tables required
- Manual calculation mode (iPad: Formulas → Calculate Now / Mac: F9)
- Monte Carlo sheet XML must stay under 7MB — use net spending helper table pattern
- Must be shareable via email
- Must be auditable (all formulas visible)

-----

## 2. File Structure — 7 Sheets

|#|Sheet Name             |Purpose                                                                    |
|-|-----------------------|---------------------------------------------------------------------------|
|1|**Benchmarks**         |US national + Boston HCOL percentile reference with your position indicator|
|2|**Accounts**           |Individual account entry — source of truth for all asset values            |
|3|**Parameters**         |Central control panel — all assumptions and inputs                         |
|4|**30-Year Projection** |Deterministic year-by-year model                                           |
|5|**Monte Carlo**        |2,000-scenario stochastic simulation                                       |
|6|**Scenario Comparison**|6 pre-built strategies side-by-side                                        |
|7|**Recommendations**    |Executive summary and action timeline                                      |

-----

## 3. Sheet 1 — Benchmarks

### Purpose

Reference context showing where your household stands relative to US national and Boston-area percentiles across 5 categories. Both current (working) and retirement numbers shown side by side. Static benchmark data hardcoded from most recent available sources (IRS / Federal Reserve / Census, 2023-2024).

### Header Section

- Title: “HOUSEHOLD BENCHMARKS · US National + Boston HCOL · [DATA_YEAR]”
- Note: “Benchmark data sourced from IRS SOI, Federal Reserve SCF, BLS CEX. Top 0.1% excluded from scale.”
- Note: “Your numbers pull live from Parameters and Projection sheets. Press Calculate Now / F9 to refresh.”

### Layout — 5 Benchmark Tables

Each table follows this structure:

|Percentile|US National   |Boston HCOL   |Your Current|Your Retirement|Your Position|
|----------|--------------|--------------|------------|---------------|-------------|
|Top 10%   |$X            |$X            |[formula]   |[formula]      |[formula]    |
|Top 5%    |$X            |$X            |            |               |             |
|Top 2%    |$X            |$X            |            |               |             |
|Top 1%    |$X            |$X            |            |               |             |
|Note      |excl. top 0.1%|excl. top 0.1%|            |               |             |

**Position indicator logic:**

- TEXT formula comparing your number to each threshold
- Labels: “Above Top 1%” / “Top 1%-2%” / “Top 2%-5%” / “Top 5%-10%” / “Top 10%+”
- Cell color: Green = Top 2%+ · Blue = Top 5%-2% · Light blue = Top 10%-5% · Gray = below Top 10%

### Table 1 — Household Income (Pre-Tax)

|Percentile|US National|Boston HCOL|
|----------|-----------|-----------|
|Top 10%   |$153,000   |$210,000   |
|Top 5%    |$220,000   |$295,000   |
|Top 2%    |$350,000   |$465,000   |
|Top 1%    |$540,000   |$720,000   |

- Your Current: pulls [GROSS_INCOME] from Parameters
- Your Retirement: =Spouse 1 SS + Spouse 2 SS + Pension + Deferred Comp + Part-Time Income (Yr 0 of Projection)

### Table 2 — Household Income (Post-Tax / Take-Home)

|Percentile|US National|Boston HCOL|
|----------|-----------|-----------|
|Top 10%   |$110,000   |$148,000   |
|Top 5%    |$155,000   |$205,000   |
|Top 2%    |$238,000   |$312,000   |
|Top 1%    |$352,000   |$463,000   |

- Your Current: pulls [TAKE_HOME] from Parameters
- Your Retirement: =Total Income (Yr 0 Projection) minus Total Tax (Yr 0 Projection)

### Table 3 — Net Worth / Wealth

|Percentile|US National|Boston HCOL|
|----------|-----------|-----------|
|Top 10%   |$1,200,000 |$1,750,000 |
|Top 5%    |$2,300,000 |$3,200,000 |
|Top 2%    |$4,800,000 |$6,500,000 |
|Top 1%    |$11,100,000|$14,800,000|

- Your Current: =Total Assets from Accounts sheet
- Your Retirement: =Total Assets (Yr 0 of Projection — assets at retirement date)

### Table 4 — Annual Spending / Consumption

|Percentile|US National|Boston HCOL|
|----------|-----------|-----------|
|Top 10%   |$98,000    |$138,000   |
|Top 5%    |$142,000   |$196,000   |
|Top 2%    |$215,000   |$292,000   |
|Top 1%    |$320,000   |$430,000   |

- Your Current: pulls [CURRENT_SPENDING] from Parameters
- Your Retirement: pulls [SPENDING_TARGET] from Parameters

### Table 5 — Retirement Income (Age 60-70 Cohort)

|Percentile|US National|Boston HCOL|
|----------|-----------|-----------|
|Top 10%   |$95,000    |$128,000   |
|Top 5%    |$148,000   |$195,000   |
|Top 2%    |$248,000   |$320,000   |
|Top 1%    |$412,000   |$535,000   |

- Your Current: “N/A — currently working” (gray label)
- Your Retirement: =Total Income Sources Yr 0 (SS + Pension + Deferred Comp + Part-Time)

### Summary Row (below all tables)

Single summary line per life stage:

|                                        |Your Current|Your Retirement|
|----------------------------------------|------------|---------------|
|Overall Position (avg across categories)|[formula]   |[formula]      |
|Highest ranking category                |[formula]   |[formula]      |
|Lowest ranking category                 |[formula]   |[formula]      |

### Data Source Notes (below summary)

- Income data: IRS Statistics of Income 2023
- Wealth data: Federal Reserve Survey of Consumer Finances 2022
- Spending data: BLS Consumer Expenditure Survey 2023
- Boston HCOL overlay: 35% premium applied to national thresholds (BLS metro data)
- Top 0.1% excluded — thresholds above $3M income / $43M net worth distort scale
- Data updated manually — not connected to live sources

-----

## 4. Sheet 2 — Accounts

### Purpose

Single source of truth for all asset balances. Parameters sheet pulls via SUMIF — no manual entry of totals required.

### Header Section

- Title: “ACCOUNTS & BALANCES · As of March 2026”
- ESPP Stock Price: single shared yellow input cell — all ESPP rows reference this
- Last Updated: yellow date input field

### Main Table Columns

|Col|Field        |Type           |Notes                                                |
|---|-------------|---------------|-----------------------------------------------------|
|A  |Account Name |Yellow text    |Free text                                            |
|B  |Account Type |Yellow dropdown|8 options — see below                                |
|C  |Owner        |Yellow dropdown|Spouse 1 · Spouse 2 · Joint                          |
|D  |Stock Symbol |Yellow text    |ESPP rows only — grayed out for others               |
|E  |Shares Held  |Yellow number  |ESPP rows only — grayed out for others               |
|F  |Current Price|Blue formula   |ESPP: =$ESPP_price_cell · Others: blank              |
|G  |Balance      |Formula / Input|ESPP: =Shares×Price (blue) · All others: yellow input|
|H  |Notes        |Text           |Optional                                             |

### Account Type Dropdown — 8 Options

1. Traditional IRA
1. 401k
1. Roth IRA
1. Roth 401k
1. FHSA
1. Taxable
1. Cash
1. ESPP

### Model Bucket Mapping

|Account Type   |Model Bucket        |
|---------------|--------------------|
|Traditional IRA|Traditional IRA/401k|
|401k           |Traditional IRA/401k|
|Roth 401k      |Traditional IRA/401k|
|Roth IRA       |Roth IRA            |
|FHSA           |Roth IRA            |
|Taxable        |Taxable             |
|Cash           |Taxable             |
|ESPP           |Taxable             |

### Pre-Populated Accounts — 21 Rows

|# |Account Name        |Type           |Owner   |Balance                                  |
|--|--------------------|---------------|--------|-----------------------------------------|
|1 |401k                |401k           |Spouse 1|[401K_BAL]                               |
|2 |Traditional IRA #1  |Traditional IRA|Spouse 1|[SP1_TRAD_IRA_1]                         |
|3 |Traditional IRA #2  |Traditional IRA|Spouse 1|[SP1_TRAD_IRA_2]                         |
|4 |Traditional IRA #3  |Traditional IRA|Spouse 1|[SP1_TRAD_IRA_3]                         |
|5 |Traditional IRA #1  |Traditional IRA|Spouse 2|[SP2_TRAD_IRA_1]                         |
|6 |Traditional IRA #2  |Traditional IRA|Spouse 2|[SP2_TRAD_IRA_2]                         |
|7 |Traditional IRA #3  |Traditional IRA|Spouse 2|[SP2_TRAD_IRA_3]                         |
|8 |Roth IRA            |Roth IRA       |Spouse 1|[SP1_ROTH_IRA]                           |
|9 |Roth IRA            |Roth IRA       |Spouse 2|[SP2_ROTH_IRA]                           |
|10|FHSA                |FHSA           |Spouse 1|[FHSA_BAL]                               |
|11|Taxable Brokerage #1|Taxable        |Joint   |[TAX_BROK_1]                             |
|12|Taxable Brokerage #2|Taxable        |Joint   |[TAX_BROK_2]                             |
|13|Taxable Brokerage #3|Taxable        |Joint   |[TAX_BROK_3]                             |
|14|Taxable Brokerage #4|Taxable        |Joint   |[TAX_BROK_4]                             |
|15|Taxable Brokerage #5|Taxable        |Joint   |[TAX_BROK_5]                             |
|16|Bank Account #1     |Cash           |Joint   |[BANK_1]                                 |
|17|Bank Account #2     |Cash           |Joint   |[BANK_2]                                 |
|18|Bank Account #3     |Cash           |Joint   |[BANK_3]                                 |
|19|ESPP Lot #1         |ESPP           |Spouse 2|[SP2_ESPP_1_SHARES] shares @ [ESPP_PRICE]|
|20|ESPP Lot #2         |ESPP           |Spouse 2|[SP2_ESPP_2_SHARES] shares @ [ESPP_PRICE]|
|21|ESPP Lot #3         |ESPP           |Spouse 2|[SP2_ESPP_3_SHARES] shares @ [ESPP_PRICE]|
|  |**TOTAL**           |               |        |**[TOTAL_ASSETS]**                       |

5 blank rows below for future accounts. SUMIFs cover full range.

### Summary Box (below table) — Blue calculated cells

|Bucket              |SUMIF Includes                    |Target            |
|--------------------|----------------------------------|------------------|
|Traditional IRA/401k|Traditional IRA + 401k + Roth 401k|[TRAD_TOTAL]      |
|Roth IRA            |Roth IRA + FHSA                   |[ROTH_TOTAL]      |
|Taxable             |Taxable + Cash + ESPP             |[TAXABLE_TOTAL]   |
|**Total**           |Sum of above                      |**[TOTAL_ASSETS]**|

Breakdown by Owner: Spouse 1 / Spouse 2 / Joint subtotals.

-----

## 5. Sheet 3 — Parameters

### Color Coding

- **Yellow:** User inputs (editable)
- **Green:** Linked from Accounts sheet (do not edit here)
- **Blue:** Calculated formulas
- **Gray:** Reference data (fixed)

### Section [A] — Current Situation (Gray)

- Current Date: March 2026
- Spouse 1 Age: [SP1_AGE] (turns [RETIRE_AGE_A] in [SP1_BIRTH_MONTH] 2026)
- Spouse 2 Age: [SP2_AGE]

### Section [B] — Assets & Allocation

- Traditional IRA/401k → SUMIF from Accounts — **Green**
- Roth IRA → SUMIF from Accounts — **Green**
- Taxable → SUMIF from Accounts — **Green**
- Total Assets → sum of above — **Blue**
- Target Yield (Portfolio Growth Rate): [TARGET_YIELD] default — **Yellow**

### Section [C] — Work Timeline

- Work Event Outcome: Work Event 1 / Work Event 2 (dropdown, yellow)
- Additional Years Working: 0, 1, or 2 (yellow)
- Retirement Age: =62 + additional years (blue)
- Retirement Calendar Year: =2026 + additional years (blue)
- Estimated Assets at Retirement: formula (blue)

### Section [D] — Income Sources

**Social Security:**

- Spouse 1 SS Claim Age: [SP1_SS_CLAIM_AGE] default (yellow) → calculated benefit shown
- Spouse 2 SS Claim Age: [SP2_SS_CLAIM_AGE] default (yellow) → calculated benefit shown
- Interpolation: Age 62-67 linear ramp / Age 67-70 linear ramp (see VALUES.md)
- Spouse 2 SS starts when Spouse 2 reaches claim age

**Fixed Income:**

- Pension: [PENSION_ANNUAL]/year, starts at retirement (gray)
- Deferred Comp: [DEFCOMP_YR1] Year 1 → [DEFCOMP_YR15] Year 15, linear ramp, then $0 (gray)

**Current Income (for Benchmarks sheet):**

- Gross Household Income (working): [GROSS_INCOME] (yellow)
- Take-Home Pay (working): [TAKE_HOME] (yellow)
- Current Annual Spending: [CURRENT_SPENDING] (yellow)

**Part-Time Income (default $0 — inactive until filled):**

|Field                             |Type    |Default |
|----------------------------------|--------|--------|
|Spouse 1 Part-Time Income (annual)|Yellow  |$0      |
|Spouse 1 Start Age                |Yellow  |0       |
|Spouse 1 End Age                  |Yellow  |0       |
|Spouse 1 Income Type              |Dropdown|Ordinary|
|Spouse 2 Part-Time Income (annual)|Yellow  |$0      |
|Spouse 2 Start Age                |Yellow  |0       |
|Spouse 2 End Age                  |Yellow  |0       |
|Spouse 2 Income Type              |Dropdown|Ordinary|

Income Type: **Ordinary** / **Self-Employment**

- Ordinary: taxed at standard brackets
- Self-Employment: ordinary income + 15.3% SE tax; one-half SE tax deductible per IRS rules
- Active when: Amount > $0 AND Spouse Age >= Start Age AND Spouse Age <= End Age

### Section [D2] — Windfall Events (default $0 — inactive until filled)

3 independent windfall slots. Each activates in the year Spouse 1 reaches the specified age.

|Field                    |Type       |Default|
|-------------------------|-----------|-------|
|Windfall 1 Name          |Yellow text|blank  |
|Windfall 1 Amount        |Yellow     |$0     |
|Windfall 1 Age (Spouse 1)|Yellow     |0      |
|Windfall 1 Tax Treatment |Dropdown   |Taxable|
|Windfall 2 Name          |Yellow text|blank  |
|Windfall 2 Amount        |Yellow     |$0     |
|Windfall 2 Age (Spouse 1)|Yellow     |0      |
|Windfall 2 Tax Treatment |Dropdown   |Taxable|
|Windfall 3 Name          |Yellow text|blank  |
|Windfall 3 Amount        |Yellow     |$0     |
|Windfall 3 Age (Spouse 1)|Yellow     |0      |
|Windfall 3 Tax Treatment |Dropdown   |Taxable|

Tax Treatment: **Taxable** / **Capital Gains** / **Tax-Free**

- Active when: Amount > $0 AND Age > 0

### Section [E] — Spending

- Annual Target Spending: [SPENDING_TARGET] (yellow)
- Mortgage: [MORTGAGE_ANNUAL]/year, [MORTGAGE_YEARS] years from retirement (gray)
- Healthcare: $0 → [HEALTHCARE_MIXED] (2030-2035) → [HEALTHCARE_MEDICARE] (2035+)
- All healthcare included in spending target

### Section [F] — Tax Parameters (Federal MFJ 2025)

- Standard Deduction: [STD_DEDUCTION] (yellow)
- Brackets: 10%/12%/22%/24%/32% (see VALUES.md)
- Capital Gains: 0%/15%/20% (see VALUES.md)
- SS Taxable: 85%
- RMD Start Age: 73 (SECURE 2.0)

### Section [G] — Portfolio Income Mix

- Capital Gains %: [CG_PCT] (yellow)
- Qualified Dividends %: [DIV_PCT] (yellow)
- Interest/Ordinary %: [ORD_PCT] (yellow)
- Validation: must sum to 100%

### Section [H] — Withdrawal Strategy

- IRA Bracket Target: [IRA_BRACKET_TARGET] top of 22% bracket (yellow)
- Strategy: IRA First → fill bracket → Taxable → Roth (emergency only)

### Section [I] — Monte Carlo Parameters

- Expected Annual Return: [MC_RETURN] (yellow)
- Standard Deviation: [MC_STDEV] (yellow)
- Scenarios: 2,000 (reference)
- Planning Horizon: 30 years (reference)

### Section [J] — Legacy Target

- Legacy Target Total: [LEGACY_TARGET] (yellow)
- Number of Children: [NUM_CHILDREN] (gray)
- Per Child: =Legacy / children (blue)

### Section [K] — RMD Factors

IRS Uniform Lifetime Table, ages 73-92. Gray reference cells.

-----

## 6. Sheet 4 — 30-Year Projection

### Purpose

Deterministic year-by-year model. 30 rows (years 0-29). ~38 columns.

### Column Groups

- **Timeline:** Year · Spouse 1 Age · Spouse 2 Age · Calendar Year · Phase
- **Starting Balances:** IRA · Roth · Taxable · Total
- **Portfolio Growth:** Total · CG · Dividend · Ordinary portions
- **Income Sources:** Spouse 1 SS · Spouse 2 SS · Pension · Deferred Comp · Spouse 1 Part-Time · Spouse 2 Part-Time · Windfall · Fixed Income Total
- **Withdrawals:** IRA Withdrawal · Taxable Withdrawal · RMD Required
- **Tax Calculation:** Ordinary Income · CG Income · SE Tax · Std Deduction · Taxable Ordinary · Ordinary Tax · CG Tax · Total Tax
- **Spending:** Target spending
- **Ending Balances:** IRA End · Roth End · Taxable End · Total End
- **Cash Flow Check:** Net (should be ≥ 0)

### Key Logic — No Circular References

1. Fixed income (SS + pension + deferred comp)
1. Part-time income if age within range
1. Windfall if Spouse 1 age matches
1. RMD requirement (age 73+)
1. IRA withdrawal = MAX(RMD, bracket fill)
1. Taxable withdrawal = remaining gap
1. Full ordinary income including IRA withdrawal
1. SE tax if applicable
1. Total tax (ordinary + CG + SE)
1. Ending balances

### Summary Box

Final portfolio · Legacy per child · vs targets · Total taxes · Avg annual tax · Total RMDs · Total IRA withdrawn · IRA balance at 73 · First RMD at 73

-----

## 7. Sheet 5 — Monte Carlo

### Purpose

2,000-scenario stochastic simulation. iPad-compatible. No VBA.

### Performance Architecture (Critical for iPad)

- Net spending helper table: 30 rows pre-computing net spend per year
- Includes part-time income deduction and windfall boost per year
- Each MC row references helper cells — NOT embedded full formulas
- Sheet XML must stay under 7MB

### Structure

- Rows 4-2003: 2,000 scenarios
- Cols B-AE: 30 random returns via NORMINV(RAND(), [MC_RETURN], [MC_STDEV])
- Col AF: Final portfolio (30-step chain)
- Col AG: Success flag
- Cols AN-AO: Net spending helper table

### Calculation Mode

Manual. Never set to automatic.

### Results Panel

- Success Rate — large and prominent (key metric for Spouse 2)
- Percentiles: 10th · 25th · 50th · 75th · 90th
- Legacy probability at [LEGACY_TARGET_A] and [LEGACY_TARGET_B]
- Confidence intervals: 50% and 90%
- Risk metrics: mean · std dev · worst · best
- Spending sensitivity reference

-----

## 8. Sheet 6 — Scenario Comparison

### 6 Pre-Built Scenarios

|#  |Name                    |Retire Age    |SP1 SS        |SP2 SS        |Strategy               |
|---|------------------------|--------------|--------------|--------------|-----------------------|
|1  |Baseline (Taxable First)|[RETIRE_AGE_A]|[RETIRE_AGE_A]|[RETIRE_AGE_A]|Taxable First          |
|2 ★|IRA to 22% (RECOMMENDED)|[RETIRE_AGE_A]|[RETIRE_AGE_A]|[RETIRE_AGE_A]|IRA → 22% Bracket      |
|3  |IRA to 12% Bracket      |[RETIRE_AGE_A]|[RETIRE_AGE_A]|[RETIRE_AGE_A]|IRA → 12% Bracket      |
|4  |SS Delay Age 67         |[RETIRE_AGE_A]|67            |[RETIRE_AGE_A]|IRA 22% + SS@67        |
|5  |SS Delay Age 70         |[RETIRE_AGE_A]|70            |[RETIRE_AGE_A]|IRA 22% + SS@70        |
|6  |HELOC + Roth Conv.      |[RETIRE_AGE_A]|70            |70            |Roth Conversion + HELOC|

### Winner Box

Scenario 2 — lowest lifetime cost · strongest simple-strategy legacy · zero complexity.

-----

## 9. Sheet 7 — Recommendations

### Sections

1. Current Situation summary
1. Optimal Strategy (work event dependent)
1. Execute in Either Scenario
1. Expected Outcomes table
1. Action Timeline
1. Three-Sentence Summary

### Key Outcomes (Base Case)

- Final portfolio at age 92: [PROJ_FINAL_PORT]
- After-tax legacy: [PROJ_LEGACY] ([PROJ_LEGACY_PER_CHILD] per child)
- Total lifetime taxes: [PROJ_TAXES] (vs [BASELINE_TAXES] baseline)
- RMD at age 73: [OPTIMIZED_RMD] (vs [UNOPTIMIZED_RMD] unoptimized)
- Monte Carlo success rate: [MC_SUCCESS_RATE]
- 90% confidence range: [MC_P10] – [MC_P90]

-----

## 10. Financial Reference (Placeholders)

### Assets

- Traditional IRA/401k: [TRAD_PCT] · Roth IRA: [ROTH_PCT] · Taxable: [TAXABLE_PCT]
- **Total: [TOTAL_ASSETS]**

### Retirement Timeline

- Age [RETIRE_AGE_A]: ~[ASSETS_AT_62]
- Age [RETIRE_AGE_B]: ~[ASSETS_AT_63]
- Age [RETIRE_AGE_C]: ~[ASSETS_AT_64]
- Each additional year adds ~[ANNUAL_ACCUM]

### IRA Depletion Strategy

- Annual withdrawal target ages 62-72: [IRA_ANNUAL_WITHDRAWAL]
- Depletion goal before age 73: [IRA_DEPLETION_TARGET]
- RMD result: [OPTIMIZED_RMD] vs [UNOPTIMIZED_RMD]
- Lifetime tax savings: [TAX_SAVINGS_LIFETIME]

-----

## 11. Validation Tests

### Unit Tests

1. SS interpolation: Age 64 = [SS_AGE_64_BENEFIT]
1. RMD: Age 73 with [RMD_TEST_BALANCE] = [RMD_TEST_RESULT] (3.77%)
1. Tax: [TAX_TEST_INCOME] ordinary income = [TAX_TEST_RESULT]
1. Deferred comp: Year 8 = [DEFCOMP_YR8]
1. Spouse 2 SS starts at Spouse 2 claim age, not Spouse 1 retirement
1. Part-time income inactive outside start/end range or when $0
1. Windfall inactive when amount = $0 or age = 0
1. SE tax: [SE_TAX_TEST_INCOME] = [SE_TAX_TEST_RESULT]
1. Benchmarks: position indicator updates when Parameters values change
1. Benchmark position: [TOTAL_ASSETS] correctly classified vs net worth thresholds

### Integration Tests

1. Retire [RETIRE_AGE_A] vs [RETIRE_AGE_B]: assets differ by ~[ANNUAL_ACCUM]
1. IRA depletion: RMD at 73 drops from [UNOPTIMIZED_RMD] to ~[OPTIMIZED_RMD]
1. Monte Carlo: success rate improves with part-time income active
1. Windfall: portfolio increases in correct year with correct tax
1. Accounts → Parameters: balance change flows through to Benchmarks position
1. Benchmarks: retirement column updates when Projection Yr 0 changes

### Error Checks (Red Cells)

- Portfolio goes negative
- Withdrawal exceeds available balance
- Portfolio mix ≠ 100%
- RMD not taken when required (age 73+)
- Accounts total ≠ Parameters total
- Part-time end age < start age
- Windfall age outside planning horizon

-----

## 12. Priority Hierarchy

1. **Live well first** — [SPENDING_TARGET] spending non-negotiable
1. **Tax minimization** — reduce where possible without sacrificing #1
1. **Legacy** — nice to have, not a must

-----

## 13. Build Workflow

**PM:** Spouse 1 — reviews spec, approves changes, provides GO
**Stakeholder:** Spouse 2 — Monte Carlo success rate is key metric
**Tool:** Claude builds via Python/openpyxl, delivers .xlsx
**Spec location:** SPEC_public.md in GitHub repo (sanitized)
**Values location:** VALUES.md — stored locally only, never committed to GitHub
**Process:** Spec change → alignment → GO → build → validate → deliver

-----

*Version: 1.2 | Status: Approved | Numbers sanitized — see private VALUES.md*
