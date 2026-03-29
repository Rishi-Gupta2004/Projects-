

# Model RG SNDK — Handoff Document

> **Purpose:** This document is a comprehensive technical and visual reference for the `Model RG SNDK.xlsm` workbook — a **company-agnostic equity valuation framework** currently demonstrated with **Sandisk Corporation (SNDK)**. The model is designed to be re-pointed at any public company by updating the Assumptions sheet. It is intended as a handoff to a portfolio/GitHub agent to build a professional project page. The VBA code has been sent separately; this document covers **what the end user sees**, the named-range architecture, sheet-by-sheet structure, visual design language, and data flow.

> **Version:** v6.4 (25/03/26)
> **File type:** `.xlsm` (macro-enabled workbook)
> **Total sheets:** 9

---

## Table of Contents

1. [Project Overview](#project-overview)
2. [Company-Agnostic Design](#company-agnostic-design)
3. [Architecture & Data Flow](#architecture--data-flow)
4. [Visual Design Language](#visual-design-language)
5. [Sheet-by-Sheet Reference](#sheet-by-sheet-reference)
   - [Dashboard](#1-dashboard)
   - [Assumptions](#2-assumptions)
   - [Financials](#3-financials)
   - [Valuation](#4-valuation--star-sheet)
   - [MC_Output](#5-mc_output--star-sheet)
   - [MC_Histogram](#6-mc_histogram)
   - [Scenarios](#7-scenarios)
   - [Formulas](#8-formulas)
   - [Qualitative](#9-qualitative)
6. [Named Ranges — Complete Registry](#named-ranges--complete-registry)
7. [Key Modeling Choices & Methodology](#key-modeling-choices--methodology)
8. [Emphasis Points for Portfolio](#emphasis-points-for-portfolio)

---

## Project Overview

This is a **rigorous, from-scratch equity valuation framework** built to value **any public company**. It is not a template fill-in — every formula, macro button, named range, and sheet was custom-built. The model is currently configured for **Sandisk Corporation (SNDK)** as a demonstration, but the entire architecture is company-agnostic. Swapping to a new company requires only updating the Assumptions sheet inputs (company identity, base-year financials, segment configuration, and valuation parameters) — all downstream sheets (Financials, Valuation, Scenarios, Monte Carlo) recompute automatically via named ranges.

The framework combines:

- **Cash Proxy DCF** (NI − CapEx + D&A) as the primary valuation basis, with UFCF as a reference cross-check
- **Monte Carlo Simulation** (1,000 iterations) with correlated PERT/Normal distributions, AR(1) autocorrelation, and mean-reversion dynamics
- **Three-scenario analysis** (Bull/Base/Bear) with probability-weighted IVPS
- **Penman-style reformulated financial statements** (Reformulated Balance Sheet, RNOA/ROCE decomposition, Residual Income)
- **Sensitivity tables** (Discount Rate vs. Terminal Growth Rate)
- **Asset-heavy cost modeling** (fixed + variable cost structure with PP&E schedule) — toggleable for asset-light companies
- **Segment-level revenue and gross margin modeling** (up to 4 segments with year-by-year growth/GM curves)
- **Qualitative analysis framework** (SWOT + VRIS + scenario probability mapping)
- **VBA-driven dashboard** with macro buttons for running individual modules or the full model

### Current Valuation Output (Sandisk Demonstration)

*Note: All values below reflect the Sandisk demonstration configuration. These will change entirely when the model is re-pointed at a different company.*

| Metric | Value |
|--------|-------|
| Current Price | $590.00 |
| **Cash Proxy DCF IVPS** | **$447.43** |
| UFCF DCF IVPS (ref) | $404.55 |
| Upside/Downside (Cash Proxy) | −24.2% |
| MC Mean IVPS | $10,382.35 |
| MC Median IVPS (P50) | $992.15 |
| MC % > Current Price | 54.6% |
| Probability-Weighted IVPS (Scenarios) | $571.75 |

---

## Company-Agnostic Design

The model is architected so that **no company-specific data exists outside the Assumptions sheet** (and the Formulas sheet's reformulated balance sheet section). Every downstream sheet reads exclusively from named ranges that resolve to Assumptions cells. To re-point the model at a different company:

### What the User Changes (Assumptions Sheet Only)

| Block | What to Update | Example |
|-------|---------------|---------|
| Company Identity | Name, price, shares, cash, debt | Any public equity |
| Base Year Financials | Revenue, COGS, GP, OpEx breakdown, D&A, SBC, CapEx, current assets/liabilities | From the target company's 10-K or annual report |
| Valuation Parameters | Discount rate, TGR, beta, cost of debt, ROIC, risk-free rate, ERP | Recalibrate for target's risk profile |
| Margins & Expenses | GM%, tax rate, R&D%, S&M%, G&A%, CapEx%, D&A%, SBC% | From target's historical financials |
| Segment Configuration | Segment count (1–4), segment names, base revenues, year-by-year growth and GM curves | Match the target's business segments |
| Asset-Heavy Mode | Fixed/variable cost split, PP&E schedule, fixed cost inflation | Calibrate to target's cost structure (toggle off for asset-light companies) |
| Working Capital | DSO, DIO, DPO or simple NWC%, base receivables/inventory/payables | From target's balance sheet |
| Monte Carlo Config | Distribution parameters (PERT/Normal bounds), correlations, autocorrelation | Adjust uncertainty ranges to match the target's volatility profile |
| Scenario Adjustments | Bull/Base/Bear deltas for revenue, margin, tax, DR, TGR, CapEx | Thesis-specific adjustments |
| Reformulated Balance Sheet | FY current and prior year line items on the Formulas sheet (rows 100–150) | From target's 10-K (Penman reformulation) |

### What Automatically Recomputes (Touch Nothing)

- **Financials** — Full 7-year projection rebuilds from new inputs
- **Valuation** — NPV, IVPS, sensitivity table, terminal value cross-checks all recalculate
- **Scenarios** — Bull/Base/Bear re-run with new base + adjustments
- **MC Simulation** — VBA re-draws from new distribution parameters
- **MC_Output / MC_Histogram** — Statistics and chart refresh after MC run
- **Formulas** — All 69 formulas, ratios, and Penman decomposition recompute

### Design Patterns That Enable This

1. **Named ranges as an abstraction layer** — The 275 named ranges act as a clean API. VBA macros and formulas never hardcode cell addresses; they reference `BaseRevenue`, `DiscountRate`, `SegName_1`, etc. This means the Assumptions sheet layout could even be rearranged without breaking downstream logic.
2. **Toggle-driven flexibility** — Binary toggles (`UseWACC`, `UseAssetHeavy`, `UseDetailedNWC`, `UseSegmentGM`, `UseYoYGM`, `DCFBasis`, `TVMethod`) let the user adapt the model's methodology to fit different company types:
   - Asset-heavy manufacturer → `UseAssetHeavy=1`, detailed PP&E schedule
   - Asset-light SaaS company → `UseAssetHeavy=0`, simple % of revenue costs
   - Multi-segment conglomerate → `SegmentCount=4`, independent growth/GM curves
   - Single-product company → `SegmentCount=1`
   - WACC-based discount → `UseWACC=1` with CAPM inputs
   - Flat rate discount → `UseWACC=0` with manual rate
3. **Segment architecture (1–4 segments)** — Each segment has its own named range set (`SegName_N`, `BaseRev_N`, `SegGM_N`, `Growth_N_Y1`–`Growth_N_Y11`, `SegGM_N_Y1`–`SegGM_N_Y11`). Unused segments default to zero and don't affect the model.
4. **Custom MC inputs** — Two "Custom" variable slots (`MC_Custom1`, `MC_Custom2`) with user-definable labels (`CustomInput1Name`, `CustomInput2Name`) allow adding company-specific stochastic factors without modifying VBA.

---

## Architecture & Data Flow

```
┌──────────────┐
│  DASHBOARD   │  VBA macro buttons: Run Forecast, Run All, Financial,
│  (control)   │  Valuation, Scenario, MC Sim, Clear, Export
│              │  Navigation links: Assumptions -->, Qualitative -->
└──────┬───────┘
       │ triggers VBA
       ▼
┌──────────────┐     reads all inputs from
│  ASSUMPTIONS │◄──────────────────────────────────────────────┐
│  (all inputs)│     named ranges                              │
└──────┬───────┘                                               │
       │ named range references                                │
       ▼                                                       │
┌──────────────┐                                               │
│  FINANCIALS  │  7-year forecast (base yr through end yr)     │
│  (projections│  Revenue → IS → Cash Flow → WC → PP&E        │
│   engine)    │                                               │
└──────┬───────┘                                               │
       │ cash flows feed into                                  │
       ▼                                                       │
┌──────────────┐     ┌──────────────┐                          │
│  VALUATION   │────►│  SCENARIOS   │  Bull/Base/Bear          │
│  (DCF +      │     │  (adj + prob │  adjustments re-run      │
│  sensitivity)│     │   weighted)  │  the valuation           │
└──────┬───────┘     └──────────────┘                          │
       │                                                       │
       │ VBA Monte Carlo loop                                  │
       ▼                                                       │
┌──────────────┐     ┌──────────────┐                          │
│  MC_OUTPUT   │────►│ MC_HISTOGRAM │  Distribution chart      │
│  (statistics)│     │  (50 bins)   │  Column clustered        │
└──────────────┘     └──────────────┘                          │
                                                               │
┌──────────────┐     ┌──────────────┐                          │
│  FORMULAS    │     │ QUALITATIVE  │  SWOT + VRIS framework   │
│  (reference  │     │  (analysis)  │  Bear/Bull case boxes    │
│   + reformat │     │              │  Probability of          │
│   bal sheet) │     │              │  occurrence diagram      │
└──────────────┘     └──────────────┘                          │
```

### Data Flow Summary

1. **All inputs** live on `Assumptions` — every business assumption is a named range
2. **Financials** reads from Assumptions via named ranges, projects a full forecast horizon of IS + Cash Flow + WC + PP&E
3. **Valuation** reads cash flows from Financials, computes NPV, IVPS, sensitivity table, terminal value cross-checks
4. **Scenarios** re-runs the valuation with Bull/Base/Bear adjustments (revenue, margin, tax, DR, TGR, CapEx)
5. **MC Simulation** (VBA) draws random parameters from configured distributions, re-runs the DCF N times (configurable, default 1,000), writes stats to MC_Output and histogram bins to MC_Histogram
6. **Formulas** is a read-only reference sheet documenting every formula, the Penman reformulated balance sheet, and computed ratios
7. **Qualitative** is a standalone analysis sheet (SWOT, VRIS) with navigation links
8. **Dashboard** is the VBA control center with macro buttons

---

## Visual Design Language

The workbook uses a **consistent, professional design system** throughout:

### Color Palette

| Element | Color | Hex (approx) | Usage |
|---------|-------|------|-------|
| Section headers | Dark navy blue | `#1F3864` | All sheet section banners (e.g., "Income Statement", "Valuation Summary") |
| Header text | White | `#FFFFFF` | Text inside navy section headers |
| Row label background | Light steel blue | `#D6E4F0` | Alternating row shading for major line items |
| Input cells | Yellow highlight | `#FFFF00` | Hardcoded assumption inputs that the user can change |
| Positive % / Bull | Green text | `#008000` | Bull case upside percentages |
| Negative % / Bear | Red text | `#FF0000` | Bear case downside percentages, overvalued signals |
| Base case highlight | Yellow background cell | `#FFFF00` | Center cell of sensitivity tables |
| Computed/formula text | Black | `#000000` | All formula-driven outputs |
| Body background | White | `#FFFFFF` | Default cell background |
| Subtle green tint | Light green | `#E2EFDA` | Used on some computed output cells in Formulas sheet |

### Typography & Layout

- **Sheet titles:** Large bold text at row 1 — dynamically reference `CompanyName` (e.g., "[Company] - Financial Model (Cash Proxy DCF)")
- **Section headers:** Navy blue banner bars spanning the full width, white bold text, all-caps or title case
- **Sub-sections:** Bold row labels with light blue background
- **Data rows:** Clean, unshaded, with currency formatting ($#,##0) and percentage formatting (0.0%)
- **Frozen panes:** Most sheets freeze row 1 (title/header) and sometimes columns (Financials freezes row 13 + column 8 for scrollable data)
- **Column structure:** Consistent year headers (base year through forecast end) across Financials, Valuation, and Scenarios

### UX Design

- **Dashboard-first workflow:** User starts at Dashboard, clicks macro buttons in sequence or "Run All"
- **Macro buttons:** Styled with icons, organized in two columns — left column for model steps (Run Forecast, Run All, Financial, Valuation), right column for analysis steps (Scenario, MC Sim, Clear, Export)
- **Navigation links:** "Assumptions -->" and "Qualitative -->" buttons on Dashboard for quick sheet navigation; reciprocal links exist on Qualitative ("Assumptions -->", "Dashboard -->")
- **Version stamp:** "v6.4 25/03/26" displayed in the top-right of Dashboard

---

## Sheet-by-Sheet Reference

### 1. Dashboard

**Purpose:** VBA control center and navigation hub.
**Dimensions:** 19 rows × 18 columns
**Frozen panes:** None

**Layout:**
- Row 1–2: Dark navy banner with "=== DASHBOARD ===" title (left) and version "v6.4 25/03/26" (right)
- Rows 3–7: Empty spacer
- Rows 8–18: Two button panels side by side + navigation panel to the right

**Left button panel (column B–C):**

| Button | VBA Action |
|--------|-----------|
| Run Forecast | Runs the financial projection engine |
| Run All | Runs all modules sequentially |
| Financial | Generates/refreshes the Financials sheet |
| Valuation | Runs the DCF valuation |

**Right button panel (column H–I):**

| Button | VBA Action |
|--------|-----------|
| Scenario | Runs Bull/Base/Bear scenario analysis |
| MC Sim | Runs Monte Carlo simulation (configurable iterations) |
| Clear | Clears output sheets |
| Export | Exports results |

**Navigation panel (columns N–O):**
- "Assumptions -->" with button icon
- "Qualitative -->" with button icon

**Visual note:** Buttons are displayed with a small form-control icon beside each label. The overall look is a clean, dark-themed control panel with light blue button areas against the dark navy header.

---

### 2. Assumptions

**Purpose:** Central input hub and the **only sheet a user needs to modify** when switching companies. Every business assumption, valuation parameter, MC configuration, and segment definition lives here. All downstream sheets read from this sheet exclusively via named ranges.
**Dimensions:** 58 rows × 21 columns (A–U)
**Frozen panes:** Row 1 frozen, all columns frozen

**Layout — organized into ~13 distinct blocks:**

#### Block 1: Company Identity (B3:C8)

| Field | Named Range | Demo Value (Sandisk) |
|-------|------------|---------------------|
| Company Name | `CompanyName` | Sandisk Corporation |
| Current Price ($) | `CurrentPrice` | 590 |
| Shares Outstanding (M) | `SharesOut` | 147.6 |
| Cash & Equivalents ($M) | `CashBalance` | 2,838 |
| Total Debt ($M) | `DebtBalance` | 2,583 |

#### Block 2: Time Horizon (B10:C13)

| Field | Named Range | Demo Value |
|-------|------------|-----------|
| Base Year | `BaseYear` | 2025 |
| Forecast End Year | `ForecastEndYear` | 2032 |
| Perpetuity Years | `PerpetuityYears` | 52 |

#### Block 3: Valuation Parameters (B15:C25)

| Field | Named Range | Demo Value |
|-------|------------|-----------|
| Discount Rate (flat) | `DiscountRate` | 9.00% |
| Terminal Growth Rate | `TGR` | 2.50% |
| TV Method (1=Perp, 2=Gordon, 3=Exit) | `TVMethod` | 2 |
| Use WACC? (1=Yes) | `UseWACC` | 0 |
| Risk-Free Rate | `RiskFreeRate` | 3.80% |
| Equity Risk Premium | `ERP` | 5.50% |
| Beta | `Beta` | 2.74 |
| Cost of Debt (pre-tax) | `CostOfDebt` | 6.00% |
| ROIC | `ROIC` | 12.00% |
| Max (ROIC Cap) | `Max` | 40.00% |

#### Block 4: Margins & Expenses (B26:C34)

| Field | Named Range | Demo Value |
|-------|------------|-----------|
| Gross Margin | `GrossMargin` | 34.00% |
| Tax Rate | `TaxRate` | 15.00% |
| R&D % | `RDPct` | 6.00% |
| S&M % | `SMPct` | 3.00% |
| G&A % | `GAPct` | 4.00% |
| CapEx % | `CapExPct` | 12.00% |
| D&A % | `DAPct` | 8.00% |
| SBC % | `SBCPct` | 4.00% |

#### Block 5: Share Dynamics (B36:C38)

| Field | Named Range | Demo Value |
|-------|------------|-----------|
| Annual Dilution % | `AnnualDilutionPct` | 0.50% |
| Buyback Yield % | `BuybackYield` | 0.00% |

#### Block 6: Working Capital (B40:C58)

| Field | Named Range | Demo Value |
|-------|------------|-----------|
| Use Working Capital? (1/0) | `UseWorkingCap` | 0 |
| NWC % of Revenue | `NWCPctRev` | 6.00% |
| Use Detailed NWC? (1/0) | `UseDetailedNWC` | 1 |
| DSO | `DSO` | 55 |
| DIO | `DIO` | 90 |
| DPO | `DPO` | 45 |
| Base Receivables ($M) | `BaseReceivables` | 1,350 |
| Base Inventory ($M) | `BaseInventory` | 2,200 |
| Base Payables ($M) | `BasePayables` | 730 |

#### Block 7: DCF & Capital Structure (B44:C49)

| Field | Named Range | Demo Value |
|-------|------------|-----------|
| DCF Basis (1=Cash, 2=UFCF) | `DCFBasis` | 1 |
| Exit Multiple (EV/UFCF) | `ExitMultiple` | 12 |
| Use Target Cap Struct? (1/0) | `UseTargetCapStruct` | 0 |
| Target Debt/Equity | `TargetDebtToEquity` | 0.10 |
| Debt Amort % p.a. | `DebtAmortPct` | 3.00% |

#### Block 8: Base Year Financials — $M (E3:F19)

Full base-year income statement sourced from the target company's 10-K: Revenue through Net Income, plus D&A, SBC, CapEx, Current Assets, Current Liabilities. Each line has a corresponding `Base*` named range (e.g., `BaseRevenue`, `BaseCOGS`, `BaseGrossProfit`, etc.).

#### Block 9: Monte Carlo Configuration (H3:L23)

| Variable | Named Range Pattern | Distribution | Param 1 | Param 2 | Param 3 |
|----------|-------------------|-------------|---------|---------|---------|
| Seg 1 Growth | `MC_Seg1Growth_*` | PERT | 10% | 50% | 80% |
| Seg 2–4 Growth | `MC_Seg[2-4]Growth_*` | PERT | 0 | 0 | 0 |
| Gross Margin | `MC_GM_*` | PERT | 10% | 30% | 52% |
| Tax Rate | `MC_Tax_*` | PERT | 10% | 15% | 22% |
| Discount Rate | `MC_DR_*` | Normal | 7% | 8% | 9% |
| TGR | `MC_TGR_*` | PERT | 2% | 2.5% | 4% |
| CapEx % | `MC_CapEx_*` | PERT | 1% | 3% | 0% |
| D&A % | (in MC config area) | Normal | 3% | 6% | 8% |
| SBC % | (in MC config area) | PERT | 2% | 4% | 6% |
| OpEx % | (in MC config area) | PERT | 8% | 13% | 18% |
| Custom 1 & 2 | `MC_Custom[1-2]_*` | PERT | 0 | 0 | 0 |

**MC Correlations & Dynamics (K3:L7):**

| Parameter | Named Range | Demo Value |
|-----------|------------|-----------|
| Margin↔Growth Corr | `MC_MarginGrowthCorr` | 0.50 |
| Growth↔DR Corr | `MC_GrowthDRCorr` | −0.60 |
| AR(1) Autocorrelation | `MC_AutoCorr` | 0.50 |
| Mean-Reversion Speed | `MC_MeanRevSpeed` | 0.30 |

#### Block 10: Scenario Adjustments (E26:H34)

Pattern: `Scen_{Bull|Base|Bear}_{Adjustment}`

| Adjustment | Bull | Base | Bear |
|-----------|------|------|------|
| Revenue Adj | +10% | 0% | −12% |
| Gross Margin Adj | +5% | 0% | −8% |
| Tax Rate Adj | 0% | 0% | +5% |
| Discount Rate Adj | −1% | 0% | +2% |
| Terminal Growth Adj | +0.5% | 0% | −1% |
| CapEx Adj | −2% | 0% | +3% |
| Probability | 25% | 50% | 25% |
| Decay Rate | 0% | 0% | 0% |

#### Block 11: Asset-Heavy Mode (E36:F50)

Toggle: `UseAssetHeavy = 1` (On) — turn off for asset-light companies

| Cost Component | Named Range (Fixed) | Named Range (Variable) | Demo Fixed ($M) | Demo Variable (%) |
|---------------|-------------------|---------------------|-----------------|------------------|
| COGS | `FixedCOGS` | `VarCOGSPct` | 3,000 | 4% |
| R&D | `FixedRD` | `VarRDPct` | 400 | 1.5% |
| S&M | `FixedSM` | `VarSMPct` | 180 | 1% |
| G&A | `FixedGA` | `VarGAPct` | 270 | 1% |

Fixed Cost Inflation: `FixedCostInflation` = 2%

#### Block 12: PP&E Schedule (E52:F56)

| Field | Named Range | Demo Value |
|-------|------------|-----------|
| Base PP&E ($M) | `BasePPE` | 2,400 |
| PP&E Useful Life (Years) | `PPEUsefulLife` | 10 |
| Maint CapEx (% of Net PP&E) | `MaintCapExPct` | 3% |
| Growth CapEx Factor ($/$ Rev) | `GrowthCapExFactor` | 0.04 |

#### Block 13: Segment Configuration (J27:U50)

| Field | Named Range | Demo Value |
|-------|------------|-----------|
| Segment Count | `SegmentCount` | 1 |
| Use Segment-Level GM? | `UseSegmentGM` | 1 (Yes) |
| Use YoY GM? | `UseYoYGM` | 1 (Yes) |

**Segment 1 (Demo — NAND Flash):**
- Name: `SegName_1` = "NAND Flash"
- Base Revenue: `BaseRev_1` = $8,937M
- Segment GM: `SegGM_1` = 40%
- Year-by-year growth (`Growth_1_Y1`–`Growth_1_Y11`): 100%, 40%, 20%, 10%, 10%, 8%, 6%, 5%, 4%, 4%, 3.5%
- Year-by-year GM (`SegGM_1_Y1`–`SegGM_1_Y11`): 55%, 45%, 30%, 30%, 30%, 30%, 30%, 30%, 30%, 30%, 30%

Segments 2–4: Inactive (NA, $0 base rev) — available for multi-segment companies

**Additional VBA buttons on Assumptions sheet (visible in top-right area):**
- **Delete** button (dark teal circular)
- **Restore** button (dark teal circular)
- **Asset** button (dark teal circular with star icon)

---

### 3. Financials

**Purpose:** Multi-year financial projection engine (base year through forecast end year — currently 2025–2032, adjustable via `BaseYear` and `ForecastEndYear`). Entirely formula-driven from Assumptions named ranges; contains no hardcoded company data.
**Dimensions:** 54 rows × 10 columns (A–J)
**Frozen panes:** Row 13 frozen, column 8 frozen (allows scrolling through years while keeping line item labels visible)

**Title:** "[Company Name] - Financial Model (Cash Proxy DCF)" — dynamically references `CompanyName`

**Sections (top to bottom):**

1. **Revenue Breakdown** — Segment-level revenue rolling up to Total Revenue, with Revenue Growth % row
2. **Income Statement** — COGS → Gross Profit → Gross Margin % → R&D, S&M, G&A → Total OpEx → Operating Income → Interest Income/Expense → Pre-Tax Income → Tax → Net Income → EPS → Shares Outstanding → NOL Carryforward
3. **Cash Flow Proxy (NI − CapEx + D&A)** — The primary DCF cash flow: Net Income, (−) CapEx, (+) D&A = Cash Proxy
4. **UFCF (Reference — Not Used in DCF)** — Unlevered FCF computed for cross-reference
5. **Working Capital** — Receivables, Inventory, Payables, Net Working Capital, Change in NWC (detailed method using DSO/DIO/DPO)
6. **PP&E Schedule** — Net PP&E, D&A (asset-based), Maintenance CapEx, Growth CapEx, Total CapEx
7. **Additional Items** — SBC, Debt Balance, Net Cash

**Key data points (Sandisk demo — Year 1 / Terminal):**

| Line | Base Year | Terminal Year |
|------|-----------|--------------|
| Total Revenue | $13,000M | $41,595M |
| Gross Margin % | 61.5% | 21.7% |
| Operating Income | $6,167M | $6,600M |
| Net Income | $5,140M | $6,236M |
| Cash Proxy | $4,783M | $6,310M |
| EPS | $34.82 | $40.80 |
| Net Cash | $255M | $24,923M |

---

### 4. Valuation ⭐ (Star Sheet)

**Purpose:** DCF valuation engine with terminal value cross-checks and sensitivity analysis.
**Dimensions:** 59 rows × 10 columns
**Frozen panes:** Row 1 frozen

**Title:** "[Company Name] - DCF Valuation (Cash Proxy DCF)"

**Sections:**

#### Section 1: Cash Flow Streams ($M) — Rows 3–11

Year-by-year display of:
- Net Income
- **Cash Proxy (primary)** — bolded, this is the DCF basis
- UFCF (ref) — shown as reference only
- Interest Expense, Delta NWC, NOL Carryforward, Debt Balance

#### Section 2: Valuation Summary — Rows 13–20

| Metric | Cash Proxy DCF | UFCF DCF (ref) |
|--------|---------------|-----------------|
| NPV of Cash Flows | $68,132M | $61,578M |
| (+) Net Cash | $255M | $255M |
| **Equity Value** | **$68,387M** | **$61,833M** |
| (/) Diluted Shares | 153 | — |
| **Intrinsic Value Per Share** | **$447.43** | **$404.55** |
| Upside / Downside | −24.2% (red) | −31.4% (red) |

The IVPS is displayed in large, bold format. The upside/downside is color-coded (red = overvalued, green = undervalued).

#### Section 3: Terminal Value Cross-Check — Rows 22–26

| Method | NPV | IVPS |
|--------|-----|------|
| 52-yr Perpetuity | $66,106M | $434.18 |
| **Gordon Growth (selected)** | **$68,132M** | **$447.43** |
| Exit Multiple 12.0x | $55,897M | $367.38 |

#### Section 4: Sensitivity Table — DR vs TGR — Rows 28–36

An 8×7 matrix (Discount Rate 6%–13% vs. Terminal Growth Rate 1%–4%).

- Center cell (base DR / base TGR) = **base case IVPS** — highlighted in **yellow**
- Values range from lowest (high DR / low TGR) to highest (low DR / high TGR)
- Color gradient: green (undervalued) to neutral, with the base case in yellow

#### Section 5: Key Assumptions Summary — Rows 38–59

A clean reference panel listing all key assumptions used in this valuation run:
- Company, DCF Basis, Discount Rate/WACC, WACC Method, TGR, TV Method, Tax Rate
- Mode (Asset-Heavy or simple), Fixed COGS, Variable COGS Fallback, Fixed Cost Inflation
- PP&E parameters, Shares, Net Cash, Current Price, Debt Amort Rate, Working Capital method

---

### 5. MC_Output ⭐ (Star Sheet)

**Purpose:** Monte Carlo simulation results — statistical summary of N DCF iterations.
**Dimensions:** 20 rows × 3 columns
**Frozen panes:** Row 1 frozen

**Title:** "[Company Name] - Monte Carlo Results (N iter, NI DCF)"

#### Statistics Table

| Statistic | NI DCF IVPS |
|-----------|-------------|
| Mean | $10,382.35 |
| Std Dev | $17,732.51 |
| **Median (P50)** | **$992.15** |
| P5 (Bear) | −$451.34 |
| P10 | −$191.31 |
| P25 | $13.71 |
| P75 | $14,925.99 |
| P90 | $35,923.42 |
| P95 (Bull) | $48,542.82 |
| Min | −$12,548.22 |
| Max | $103,973.96 |

#### Price Comparison

| Metric | Value |
|--------|-------|
| Current Price | $590.00 |
| % Iterations > Current Price | 54.6% |
| **Mean Upside/Downside** | **1659.7%** |

**Note on distribution shape:** The wide spread reflects the heavy-tailed nature of the PERT distributions and the high growth uncertainty in the demo configuration. The median ($992) vs. mean ($10,382) divergence indicates significant right-skew — a few extreme upside draws pull the mean up dramatically. This distribution shape will vary significantly by company depending on the configured uncertainty ranges.

---

### 6. MC_Histogram

**Purpose:** Distribution visualization for Monte Carlo results.
**Dimensions:** 54 rows × 4 columns
**Frozen panes:** Row 1 frozen

**Title:** "[Company Name] - NI DCF IVPS Distribution"

**Data:** 50 histogram bins with Bin Mid, Count, and Frequency % columns. Bin ranges are dynamically determined by the MC simulation results.

**Chart:** Column clustered chart titled "[Company Name] - NI DCF IVPS (N iter)"
- Chart ID: `{505AC8E3-B9AF-128F-1985-896DBD15157C}`
- Series: Frequency % (D5:D54) vs. Bin Mid (B5:B54)
- Positioned at top=20pt, left=320pt (to the right of the data table)

**Distribution shape (Sandisk demo):** Highly concentrated around the $557 bin (45.2% of iterations), with a long right tail extending to ~$74K. This "spike + tail" pattern is characteristic of the PERT distribution interaction with the DCF's sensitivity to terminal value assumptions.

---

### 7. Scenarios

**Purpose:** Bull/Base/Bear scenario analysis with probability-weighted IVPS.
**Dimensions:** 19 rows × 13 columns
**Frozen panes:** Row 1 frozen

**Title:** "[Company Name] - Scenario Analysis (Cash Proxy DCF)"

#### Scenario Comparison Table (Sandisk demo values)

| Scenario | Probability | Rev Growth Adj | Margin Adj | Tax Adj | DR Adj | TGR Adj | MaintCapEx Adj | Decay Rate | NPV | IVPS | Upside |
|----------|-------------|---------------|-----------|---------|--------|---------|---------------|-----------|-----|------|--------|
| Bull | 25% | +10% | +5% | 0% | −1% | +0.5% | −2% | 0% | $212,081M | $1,389.23 | +135.5% (green) |
| Base | 50% | 0% | 0% | 0% | 0% | 0% | 0% | 0% | $68,132M | $447.43 | −24.2% (red) |
| Bear | 25% | −12% | −8% | +5% | +2% | −1% | +3% | 0% | $191M | $2.92 | −99.5% (red) |
| **Probability-Weighted IVPS** | | | | | | | | | | **$571.75** | **−3.1%** |

#### Sensitivity Table

DR vs TGR, 8×7 grid, identical to the one on Valuation. Base case highlighted in yellow at center.

---

### 8. Formulas

**Purpose:** Complete formula reference and Penman reformulated balance sheet — this is a documentation/audit sheet, not an input sheet. The reformulated balance sheet section (rows 100–150) does contain company-specific data that needs updating when switching companies.
**Dimensions:** 150 rows × 16 columns
**Frozen panes:** Row 3 frozen

**Title:** "FORMULA REFERENCE" (left) + "COMPUTED VALUES — Linked to Model" (right)

**Left panel (columns A–E): Formula Catalog**

69 numbered formulas organized into sections:
1. **Income Statement** (#1–13): Revenue through CapEx
2. **Reformulated Balance Sheet FY** (#14–25): Total Assets, Liabilities, CSE, FA, FO, NFO, OA, OL, NOA, averages
3. **Free Cash Flow** (#26–31): OI After Tax, ΔNOA, FCF, FCF Proxy, NFE
4. **First Level — Leverage Decomposition** (#32–42): RNOA, FLEV, NBC, SPREAD, ROCE, OLLEV, Implicit Interest, ROOA, OLSPREAD
5. **Second Level — Profitability Decomposition** (#43–47): PM, ATO, DuPont check, Sales PM
6. **Turnover Ratios** (#48–50): Days Receivable, Inventory Turnover, Payable Turnover
7. **Residual Income & Growth** (#51–55): ReOI, NOA growth
8. **Valuation — No Growth** (#56–58): V₀(NOA), V₀(E), IVPS
9. **Valuation — With Growth** (#59–61)
10. **Valuation — FCF** (#62–63)
11. **Market Comparison** (#64–69): Price, Market Cap, EV, Upside

**Right panel (columns H–P): Computed Values**

- **Base Year Inputs from Assumptions** — linked values with source cell references (e.g., "Assumptions!F4")
- **Key Parameters from Assumptions** — discount rate, TGR, tax rate, margins, etc.
- **Reformulated Values (Computed)** — ρ_F, ρ_D, NFO, OI After Tax, PM, FCF Proxy, NWC, Market Cap, EV, etc.
- **Key Ratios & Valuations** — RNOA, FLEV, NBC, SPREAD, ROCE, ATO, ReOI (all auto-computed)

**Bottom section (rows 100–150): Reformulated Balance Sheet**

Full Penman-style reformulation with current FY and prior FY data:
- Current Assets → Non-Current Assets → Total Assets
- Current Liabilities → Non-Current Liabilities → Total Liabilities
- Total Shareholders' Equity
- **Penman Reformulation:** FA, FO, NFO, OA, OL, NOA + balance checks (NOA = CSE + NFO)
- Average NOA, Average CSE, Average NFO

Each line item has source notes referencing the target company's SEC filings. Balance check cells confirm 0 (balanced).

---

### 9. Qualitative

**Purpose:** Qualitative analysis framework for the investment thesis — company-specific content to be filled in by the analyst.
**Dimensions:** 27 rows × 12 columns
**Frozen panes:** Row 1 frozen

**Title:** "Qualitative Analysis"

**Layout (three distinct areas):**

**Left side: SWOT + VRIS Analysis**
- **About** block: Founded, CEO, CFO, CBO (template fields)
- **Strengths** | **Weaknesses** — side-by-side blocks with navy headers
- **Opportunities** | **Threats** — side-by-side blocks with navy headers
- **Valuable** | **Imitable** — VRIS framework top row
- **Rare** | **Substitutable** — VRIS framework bottom row

All SWOT/VRIS cells are template placeholders with navy blue header bars, to be filled per company.

**Right side: Scenario Probability Framework**
- **Bear Case** — large navy block for bear thesis narrative
- **Bull Case** — large navy block for bull thesis narrative
- **Probability of Occurrence diagram** — a horizontal flow diagram showing:
  - "Impossible" → Low probability zone
  - **"IT IS POSSIBLE"** box: "This could happen, but you are not sure what 'this' is, when it will happen and what it will look like when it does."
  - → "Gauge market potential & test products"
  - **"IT IS PLAUSIBLE"** box: "This is something that you can make a reasoned argument could happen, though you have no tangible evidence for it happening (yet)."
  - → "Product success & Financial results"
  - **"POSSIBLE SCENARIOS"** and **"PLAUSIBLE SCENARIOS"** boxes with definitions

**Navigation:**
- "Assumptions -->" link with button
- "Dashboard -->" link with button

---

## Named Ranges — Complete Registry

The workbook contains **275 named ranges** (some are legacy #REF! errors from past iterations). Below is the complete registry of **active, functional named ranges**, grouped by category. All cells reference the Assumptions sheet unless otherwise noted.

### Company Identity & Market Data

| Named Range | Cell Reference | Description |
|-------------|---------------|-------------|
| `CompanyName` | $C$4 | Company name string |
| `CurrentPrice` | $C$5 | Current market price per share |
| `SharesOut` | $C$6 | Shares outstanding (millions) |
| `CashBalance` | $C$7 | Cash & equivalents ($M) |
| `DebtBalance` | $C$8 | Total debt ($M) |

### Time Horizon

| Named Range | Cell Reference | Description |
|-------------|---------------|-------------|
| `BaseYear` | $C$11 | Forecast start year |
| `ForecastEnd` / `ForecastEndYear` | $C$12 | Last forecast year |
| `PerpetuityYears` | $C$13 | Years for perpetuity TV method |

### Valuation Parameters

| Named Range | Cell Reference | Description |
|-------------|---------------|-------------|
| `DiscountRate` | $C$16 | Flat discount rate |
| `TGR` | $C$17 | Terminal growth rate |
| `TVMethod` | $C$18 | TV method (1=Perpetuity, 2=Gordon, 3=Exit Multiple) |
| `UseWACC` | $C$19 | Toggle: 1=use WACC, 0=flat rate |
| `RiskFreeRate` | $C$20 | Risk-free rate for CAPM |
| `EquityRiskPrem` / `ERP` | $C$21 | Equity risk premium |
| `Beta` / `BetaValue` | $C$22 | Equity beta |
| `CostOfDebt` | $C$23 | Pre-tax cost of debt |
| `ROIC` | $C$24 | Return on invested capital |
| `Max` | $C$25 | ROIC cap (max) |
| `ExitMultiple` | $C$46 | EV/UFCF exit multiple (for TV method 3) |

### Margins & Expenses

| Named Range | Cell Reference | Description |
|-------------|---------------|-------------|
| `GrossMargin` | $C$27 | Gross margin % |
| `TaxRate` | $C$28 | Effective tax rate |
| `RDPct` | $C$29 | R&D as % of revenue |
| `SMPct` | $C$30 | S&M as % of revenue |
| `GAPct` | $C$31 | G&A as % of revenue |
| `CapExPct` | $C$32 | CapEx as % of revenue |
| `DAPct` | $C$33 | D&A as % of revenue |
| `SBCPct` | $C$34 | SBC as % of revenue |

### Share Dynamics

| Named Range | Cell Reference | Description |
|-------------|---------------|-------------|
| `AnnualDilutionPct` | $C$37 | Annual share dilution % |
| `BuybackYield` | $C$38 | Buyback yield % |

### Working Capital

| Named Range | Cell Reference | Description |
|-------------|---------------|-------------|
| `UseWorkingCap` | $C$41 | Toggle: 1=include WC in model |
| `NWCPctRev` | $C$42 | Simple NWC as % of revenue |
| `UseDetailedNWC` | $C$52 | Toggle: 1=use DSO/DIO/DPO method |
| `DSO` | $C$53 | Days Sales Outstanding |
| `DIO` | $C$54 | Days Inventory Outstanding |
| `DPO` | $C$55 | Days Payable Outstanding |
| `BaseReceivables` | $C$56 | Base year A/R ($M) |
| `BaseInventory` | $C$57 | Base year inventory ($M) |
| `BasePayables` | $C$58 | Base year A/P ($M) |

### DCF & Capital Structure

| Named Range | Cell Reference | Description |
|-------------|---------------|-------------|
| `DCFBasis` | $C$45 | 1=Cash Proxy (NI−CapEx+D&A), 2=UFCF |
| `UseTargetCapStruct` | $C$47 | Toggle: 1=use target D/E |
| `TargetDebtToEquity` | $C$48 | Target D/E ratio |
| `DebtAmortPct` | $C$49 | Annual debt amortization % |

### Base Year Financials

| Named Range | Cell Reference | Description |
|-------------|---------------|-------------|
| `BaseRevenue` | $F$4 | Base year total revenue |
| `BaseCOGS` | $F$5 | Base year COGS |
| `BaseGrossProfit` | $F$6 | Base year gross profit |
| `BaseRD` | $F$7 | Base year R&D expense |
| `BaseSM` | $F$8 | Base year S&M expense |
| `BaseGA` | $F$9 | Base year G&A expense |
| `BaseOpEx` | $F$10 | Base year total OpEx |
| `BaseOpInc` | $F$11 | Base year operating income |
| `BaseOtherIncome` | $F$12 | Base year other income |
| `BasePreTax` | $F$13 | Base year pre-tax income |
| `BaseNetIncome` | $F$14 | Base year net income |
| `BaseDA` | $F$15 | Base year D&A |
| `BaseSBC` | $F$16 | Base year SBC |
| `BaseCapEx` | $F$17 | Base year CapEx |
| `BaseCurrentAssets` | $F$18 | Base year current assets |
| `BaseCurrentLiab` | $F$19 | Base year current liabilities |

### Asset-Heavy Mode / Fixed+Variable Cost Structure

| Named Range | Cell Reference | Description |
|-------------|---------------|-------------|
| `UseAssetHeavy` | $F$39 | Toggle: 1=on (use for manufacturers; off for asset-light) |
| `FixedCostInflation` | $F$42 | Annual inflation on fixed costs |
| `FixedCOGS` | $F$43 | Fixed COGS component ($M) |
| `VarCOGSPct` | $F$44 | Variable COGS as % of revenue |
| `FixedRD` | $F$45 | Fixed R&D ($M) |
| `VarRDPct` | $F$46 | Variable R&D % of revenue |
| `FixedSM` | $F$47 | Fixed S&M ($M) |
| `VarSMPct` | $F$48 | Variable S&M % of revenue |
| `FixedGA` | $F$49 | Fixed G&A ($M) |
| `VarGAPct` | $F$50 | Variable G&A % of revenue |

### PP&E Schedule

| Named Range | Cell Reference | Description |
|-------------|---------------|-------------|
| `BasePPE` | $F$53 | Base year net PP&E ($M) |
| `PPEUsefulLife` | $F$54 | Useful life for depreciation |
| `MaintCapExPct` | $F$55 | Maintenance CapEx as % of Net PP&E |
| `GrowthCapExFactor` | $F$56 | Growth CapEx per $ of revenue growth |

### Monte Carlo Parameters

| Named Range | Cell Reference | Description |
|-------------|---------------|-------------|
| `MCIterations` | $I$4 | Number of MC iterations |
| `MC_MarginGrowthCorr` | $L$4 | Correlation: GM ↔ revenue growth |
| `MC_GrowthDRCorr` | $L$5 | Correlation: growth ↔ discount rate |
| `MC_AutoCorr` | $L$6 | AR(1) autocorrelation coefficient |
| `MC_MeanRevSpeed` | $L$7 | Mean-reversion speed for growth in years 11+ |
| `CustomInput1Name` | $I$5 | Label for Custom Input 1 |
| `CustomInput2Name` | $I$6 | Label for Custom Input 2 |

**MC Distribution Parameters** (pattern: `MC_{Variable}_Dist`, `MC_{Variable}_P1`, `MC_{Variable}_P2`, `MC_{Variable}_P3`):

| Variable Prefix | Dist Cell | P1 | P2 | P3 |
|-----------------|-----------|----|----|-----|
| `MC_Seg1Growth` | I10 | J10 | K10 | L10 |
| `MC_Seg2Growth` | I11 | J11 | K11 | L11 |
| `MC_Seg3Growth` | I12 | J12 | K12 | L12 |
| `MC_Seg4Growth` | I13 | J13 | K13 | L13 |
| `MC_GM` | I14 | J14 | K14 | L14 |
| `MC_Tax` | I15 | J15 | K15 | L15 |
| `MC_DR` | I16 | J16 | K16 | L16 |
| `MC_TGR` | I17 | J17 | K17 | L17 |
| `MC_CapEx` | I18 | J18 | K18 | L18 |
| `MC_Custom1` | I22 | J22 | K22 | L22 |
| `MC_Custom2` | I23 | J23 | K23 | L23 |

### Scenario Adjustments

Pattern: `Scen_{Scenario}_{Adjustment}`

| Adjustment | Bull Cell | Base Cell | Bear Cell |
|-----------|----------|----------|----------|
| `RevAdj` | F27 | G27 | H27 |
| `GMAdj` | F28 | G28 | H28 |
| `TaxAdj` | F29 | G29 | H29 |
| `DRAdj` | F30 | G30 | H30 |
| `TGRAdj` | F31 | G31 | H31 |
| `CapExAdj` | F32 | G32 | H32 |
| `Prob` | F33 | G33 | H33 |
| `Decay` | F34 | G34 | H34 |

### Segment Configuration

| Named Range | Cell Reference | Description |
|-------------|---------------|-------------|
| `SegmentCount` | $K$27 | Number of active segments (1–4) |
| `UseSegmentGM` | $K$28 | Toggle: use segment-level gross margins |
| `UseYoYGM` / `YoYGM` | $K$29 | Toggle: use year-over-year GM curves |
| `SegName_1` through `SegName_4` | K32, K37, K42, K47 | Segment names |
| `BaseRev_1` through `BaseRev_4` | M32, M37, M42, M47 | Segment base revenues |
| `SegGM_1` through `SegGM_4` | O32, O37, O42, O47 | Segment gross margins |
| `Growth_N_YM` | K–U rows 34,39,44,49 | Year-by-year growth rates per segment (N=1–4, M=1–11) |
| `SegGM_N_YM` | K–U rows 35,40,45,50 | Year-by-year gross margins per segment |

### Legacy / Broken Named Ranges (#REF!)

The following named ranges exist but resolve to `#REF!` (from past model iterations — safe to ignore or clean up):

`Alphabet_Inc.`, `BuybackPct`, `CapExPctRevFinal`, `CapExPctRevY1`, `DAPctRevFinal`, `DAPctRevY1`, `DebtWeight`, `DistSelectors`, `EquityWeight`, `ForecastYearHeaders`, `GALeverage`, `GrowthArray`, `Poo`, `RDLeverage`, `SBCToggle`, `SMLeverage`, `Total_Debt___M`, `WACCOverride`, `WACCOverride\`

Also present: `_xleta.COUNT`, `_xleta.SORT` — system-generated named ranges (Excel Lambda helpers), resolve to `#NAME?`.

---

## Key Modeling Choices & Methodology

### 1. Cash Proxy DCF as Primary Basis
The model uses **NI − CapEx + D&A** (Cash Proxy) rather than traditional UFCF as the primary valuation basis. UFCF is computed alongside for cross-reference. This is a deliberate choice — the Cash Proxy approach provides a more conservative, equity-holder-focused view. The user can switch to UFCF as primary via `DCFBasis=2`.

### 2. Asset-Heavy Cost Structure
When `UseAssetHeavy=1`, costs are modeled as **Fixed ($M) + Variable (% of revenue)** rather than a simple percentage-of-revenue approach. This captures operating leverage inherent in capital-intensive businesses (e.g., semiconductor fabs, manufacturing). For asset-light companies (SaaS, services), toggle off and use the simpler margin-based approach.

### 3. Segment-Level Modeling
Revenue is projected at the segment level (up to 4 segments) with independent year-by-year growth and gross margin curves. This allows modeling of conglomerates, multi-product companies, or businesses with distinct growth trajectories across divisions. Single-segment companies simply use `SegmentCount=1`.

### 4. Monte Carlo with Correlated Draws
The MC simulation isn't naive — it incorporates:
- **PERT and Normal distributions** (not just uniform)
- **Margin↔Growth correlation**: high growth scenarios tend to come with better (or worse) margins depending on configuration
- **Growth↔DR negative correlation**: high growth environments tend to have higher discount rates
- **AR(1) autocorrelation**: growth rates persist year-to-year, preventing unrealistic whipsawing
- **Mean-reversion**: extreme growth rates revert toward the terminal rate over time

These parameters are all configurable per company to match the target's specific risk/uncertainty profile.

### 5. Penman Reformulation
The Formulas sheet implements Stephen Penman's equity valuation framework — reformulating the balance sheet into Operating Assets/Liabilities and Financial Assets/Obligations, then decomposing returns (RNOA, FLEV, SPREAD, ROCE) for a theoretically grounded analysis.

### 6. Three Terminal Value Methods
The model supports three TV methods (configurable via `TVMethod`):
1. **52-year perpetuity** — explicit year-by-year projection (most conservative)
2. **Gordon Growth Model** — closed-form terminal value (default)
3. **Exit Multiple** — EV/UFCF based terminal value

All three are cross-checked on the Valuation sheet regardless of which is selected as primary.

### 7. Toggle-Driven Methodology
The model adapts to different company types through binary toggles:

| Toggle | Purpose | When to Use |
|--------|---------|-------------|
| `UseWACC` | WACC vs. flat discount rate | WACC for publicly traded with clear comps; flat for private or opinionated view |
| `UseAssetHeavy` | Fixed+variable vs. simple % costs | On for manufacturers, capex-heavy; off for SaaS, services |
| `UseDetailedNWC` | DSO/DIO/DPO vs. simple NWC% | Detailed for companies with meaningful WC cycles; simple for asset-light |
| `UseSegmentGM` | Segment-level vs. consolidated margins | On for multi-segment with different margin profiles |
| `UseYoYGM` | Year-over-year GM curves vs. flat | On for businesses with expected margin evolution (e.g., scaling startups) |
| `DCFBasis` | Cash Proxy vs. UFCF | Cash Proxy for equity-focused; UFCF for enterprise-value approach |
| `TVMethod` | Perpetuity / Gordon / Exit Multiple | Choose based on conviction and comparability |
| `UseTargetCapStruct` | Market vs. target capital structure | Target for LBO-like scenarios; market for status quo |

---

## Emphasis Points for Portfolio

When building the GitHub page, these two areas deserve the most prominent treatment:

### ⭐ Monte Carlo Simulation
- **N-iteration VBA-driven simulation** (configurable, default 1,000) with configurable distributions (PERT, Normal)
- **Correlated parameter draws** — not independent; margin, growth, and discount rate are realistically interlinked
- **AR(1) autocorrelation + mean-reversion** — growth rates exhibit year-to-year persistence and long-run mean reversion
- **12+ stochastic variables** simultaneously randomized per iteration
- **Full output statistics** (mean, median, percentiles P5–P95, min/max)
- **Histogram visualization** with 50 bins showing the probability distribution
- **Two custom variable slots** for company-specific stochastic factors
- Demonstrates understanding of **stochastic modeling, correlation structures, and distribution theory** beyond basic spreadsheet work

### ⭐ Valuation Engine
- **Dual DCF approach** — Cash Proxy (primary) + UFCF (reference) with automatic bridging
- **Three terminal value methods** cross-checked against each other
- **DR vs. TGR sensitivity table** with base case highlighted
- **Scenario analysis** with probability-weighted IVPS (Bull/Base/Bear)
- **Asset-heavy cost modeling** capturing operating leverage dynamics
- **Penman reformulated financials** showing academic rigor (RNOA decomposition, residual income)
- **PP&E schedule** with maintenance vs. growth CapEx separation
- **Detailed working capital** (DSO/DIO/DPO) vs. simple NWC % toggle

### Additional Differentiators
- **Company-agnostic architecture** — swap any public company by updating one sheet; all 8 downstream sheets recompute automatically. Toggle-driven methodology adapts to asset-heavy manufacturers, asset-light SaaS, single-segment or multi-segment businesses, WACC or flat-rate discounting
- **From-scratch build** — not a template, every formula and named range is custom
- **275 named ranges** creating a clean API between VBA and spreadsheet
- **9-sheet architecture** with clear separation of concerns (inputs → projections → valuation → simulation → analysis)
- **Dashboard-driven UX** with macro buttons for one-click model execution
- **Version-controlled** (v6.4), indicating iterative refinement
- **Qualitative integration** — SWOT + VRIS + scenario probability framework alongside the quantitative model

---

*End of handoff document. The receiving agent should have everything needed to build the GitHub project page without accessing the original .xlsm file. Key framing for the portfolio: this is not a Sandisk-specific model — it is a reusable equity valuation framework demonstrated with Sandisk as the reference case.*
