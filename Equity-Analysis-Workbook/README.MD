# Rishi's Equity Model — Excel VBA

A from-scratch, company-agnostic equity valuation framework built entirely in Excel VBA. This is a full-stack financial modelling tool — not a template fill-in — designed to value any public company through multiple valuation methodologies, Monte Carlo simulation, and scenario analysis. Currently demonstrated with **Sandisk Corporation (SNDK)**.

The model is architected so that switching to a new company only requires updating the Assumptions sheet. All downstream sheets — Financials, Valuation, Scenarios, Monte Carlo — recompute automatically via 275+ named ranges that act as a clean API between the VBA engine and the spreadsheet layer.

---

## How It Works

The user populates a single Assumptions sheet with the target company's financials, valuation parameters, and uncertainty ranges. From there, macro buttons on the Dashboard drive the entire workflow — projecting financials, running the DCF, stress-testing through scenarios, and simulating thousands of possible outcomes via Monte Carlo.

The model adapts to different company types through binary toggles: asset-heavy manufacturers get a fixed+variable cost structure with a PP&E schedule, while asset-light businesses use simpler margin-based projections. Multi-segment conglomerates get independent growth and gross margin curves per division. Single-product companies use one segment.

---

## Screenshots

<!-- PLACEHOLDER — replace with actual screenshot paths once uploaded -->
![Dashboard](screenshots/Dashboard.png)
![Assumptions](screenshots/Assumptions.png)
![Financials](screenshots/Financials.png)
![Valuation](screenshots/Valuation.png)
![Monte Carlo Output](screenshots/MC_Output.png)
![Monte Carlo Histogram](screenshots/MC_Histogram.png)
![Scenarios](screenshots/Scenarios.png)

---

## Core Features

### Valuation Engine

The primary valuation runs on a **Cash Proxy DCF** (Net Income − CapEx + D&A), with Unlevered Free Cash Flow computed alongside as a cross-reference. The user can switch between the two via a toggle. Three terminal value methods are supported and cross-checked against each other on every run:

- **Gordon Growth Model** — closed-form terminal value (default)
- **Explicit Perpetuity** — 52-year year-by-year projection
- **Exit Multiple** — EV/UFCF based terminal value

The Valuation sheet outputs an intrinsic value per share (IVPS), an upside/downside signal against the current market price, and a full **Discount Rate vs. Terminal Growth Rate sensitivity table** with colour-coded over/undervaluation.

### Monte Carlo Simulation

This builds on the standalone Monte Carlo Simulation Engine (Project 2 in this portfolio) but takes the concept significantly further. Where the earlier project ran independent draws from basic distributions, this implementation adds:

- **Correlated parameter draws** — gross margin and revenue growth are linked, as are growth and discount rate, so the simulation reflects real-world co-movement rather than treating variables as independent
- **AR(1) autocorrelation** — growth rates exhibit year-to-year persistence, preventing unrealistic whipsawing between extreme values
- **Mean-reversion dynamics** — extreme growth rates pull back toward the distribution centre in terminal years
- **Phase-based shock regimes** — near-term, uncertain, and terminal phases each draw from independent distributions, modelling regime shifts over the forecast horizon
- **PERT, Normal, Triangular, Uniform, and Lognormal distributions** — each variable can use whichever distribution best fits its uncertainty profile
- **12+ stochastic variables** randomised simultaneously per iteration, including two custom variable slots for company-specific factors
- **Full output statistics** — mean, median, P5–P95 percentiles, min/max, and a 50-bin histogram visualisation

The simulation runs configurable iterations (default 10,000, max 50,000) and outputs the percentage of iterations that exceed the current market price — a direct probabilistic answer to "is this stock cheap?"

### Scenario Analysis

Bull/Base/Bear scenario analysis with user-defined adjustments to revenue growth, gross margin, tax rate, discount rate, terminal growth rate, and CapEx. Adjustments can be applied as flat deltas with an optional **time-decay rate** that fades toward the base case, or as **year-level overrides** for granular control over the growth and margin path in each scenario. The output is a probability-weighted IVPS alongside the individual scenario valuations.

### Segment-Level Revenue Modelling

Revenue is projected at the segment level (up to 4 segments) with independent year-by-year growth rates and gross margin curves per segment. This allows the model to capture businesses with distinct divisions growing at different rates — a semiconductor company with a high-growth NAND division and a mature enterprise segment, for example.

### Asset-Heavy Cost Structure

When toggled on, costs split into **fixed dollar amounts** (inflated annually) and **variable percentages of revenue**, capturing the operating leverage dynamics inherent in capital-intensive businesses. A PP&E schedule drives depreciation from the asset base rather than as a simple percentage of revenue, and CapEx decomposes into maintenance (sustaining the asset base) and growth (funding incremental revenue) components.

### Penman Reformulated Financials

The Formulas sheet implements Stephen Penman's equity valuation framework — reformulating the balance sheet into Operating Assets/Liabilities and Financial Assets/Obligations, then decomposing returns through RNOA, FLEV, SPREAD, and ROCE. This sits alongside 69 numbered formulas covering everything from residual income to DuPont decomposition.

### Qualitative Framework

A dedicated Qualitative sheet provides structured space for SWOT analysis, VRIS competitive advantage assessment, and bear/bull case narratives with a probability-of-occurrence framework.

---

## Workbook Architecture

The model is a 9-sheet workbook with clear separation of concerns:

| Sheet | Role |
|-------|------|
| **Dashboard** | VBA control centre — macro buttons for running individual modules or the full model |
| **Assumptions** | The only sheet the user modifies — all inputs, toggles, segment config, MC parameters |
| **Financials** | Multi-year projected Income Statement, Cash Flow, Working Capital, PP&E Schedule |
| **Valuation** | DCF output, equity bridge, IVPS, terminal value cross-checks, sensitivity table |
| **Scenarios** | Bull/Base/Bear with probability-weighted IVPS and sensitivity matrix |
| **MC_Output** | Monte Carlo statistics — mean, median, percentiles, price comparison |
| **MC_Histogram** | 50-bin distribution chart of simulated IVPS outcomes |
| **Formulas** | Reference sheet — 69 formulas, Penman reformulated balance sheet, computed ratios |
| **Qualitative** | SWOT, VRIS, bear/bull thesis, probability-of-occurrence diagram |

Data flows from Assumptions → Financials → Valuation → Scenarios / Monte Carlo. Named ranges create an abstraction layer so the VBA never hardcodes cell addresses.

---

## VBA Modules

| Module | Purpose |
|--------|---------|
| **Engine** | Core forecast builder and NPV calculator — handles segment revenue rollup, asset-heavy/light cost routing, NOL carryforward, debt amortisation, working capital |
| **Valuation** | Writes the DCF valuation sheet — equity bridge, terminal value cross-checks, sensitivity grid |
| **MonteCarlo** | Simulation engine — correlated draws, autocorrelation, mean-reversion, phase-based shocks, histogram generation |
| **Financials** | Writes the projected financial statements to the Financials sheet |
| **Scenarios** | Bull/Base/Bear scenario runner with time-decaying and year-level adjustments |
| **Distributions** | Statistical distribution samplers — PERT, Normal, Triangular, Uniform, Lognormal, plus QuickSort and percentile functions |
| **Constants** | Public types (ModelInputs, YearFinancials, ScenarioConfig, MCDistParams, MCPhaseConfig), design system colour palette, sheet name constants |
| **Helpers** | Input readers, validators, named range utilities, WACC calculator, share dilution logic |
| **modFormatStyles** | Shared design system — consistent formatting across all output sheets |
| **RunALL** | Master orchestrator — individual entry points and full-model execution |
| **AssetHeavy** | Toggle for switching between asset-heavy and asset-light cost structures |
| **ScenarioHelper** | Setup utility for scenario named ranges |
| **Movement** | Sheet navigation (Dashboard, Assumptions, Qualitative) |
| **NoteEditor** | Cell comment management — delete and restore notes |

---

## Toggle-Driven Methodology

The model adapts to different company types through binary toggles on the Assumptions sheet:

| Toggle | What It Does |
|--------|-------------|
| `UseWACC` | WACC (CAPM-derived) vs. flat discount rate |
| `UseAssetHeavy` | Fixed+variable cost structure with PP&E schedule vs. simple % of revenue |
| `UseDetailedNWC` | DSO/DIO/DPO working capital vs. simple NWC % of revenue |
| `UseSegmentGM` | Segment-level gross margins vs. consolidated |
| `UseYoYGM` | Year-over-year gross margin curves vs. flat margins |
| `DCFBasis` | Cash Proxy (NI−CapEx+D&A) vs. UFCF as primary valuation |
| `TVMethod` | Explicit Perpetuity / Gordon Growth / Exit Multiple |
| `UseTargetCapStruct` | Market capital structure vs. target D/E for WACC |

---

## Tech

**Excel VBA** · Named Range Architecture (275+) · Statistical Distributions (PERT, Normal, Triangular, Uniform, Lognormal) · Monte Carlo Simulation with Cross-Correlation, Autocorrelation & Phase-Based Shock Regimes · DCF Valuation · Scenario Analysis with Year-Level Overrides · Penman Reformulated Financials · Asset-Heavy Cost Modelling · Segment-Level Revenue Projection · Design System (modFormatStyles)

---

## What I Learned

<!-- Rishi to fill in -->
