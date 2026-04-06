# Rishi's Equity Model

Pretty proud of this one.

I feel like it's the first consolidation of my knowledge in a very competent manner. Building and using this was such a terrific experience. It showed me a lot of valuable tools and skills and got me to really familiarise myself with the granularity that comes with fundamental financial analysis.

Agentic Engineering was utilised in error handling, cleanup, and note making.

---

## How It Works

The user populates a single Assumptions sheet with the target company's financials, valuation parameters, and uncertainty ranges. From there, macro buttons on the Dashboard drive the entire workflow.

The model adapts to different company types through binary toggles:

- Asset-heavy manufacturers get a fixed+variable cost structure with a PP&E schedule
- Asset-light businesses use simpler margin-based projections
- Multi-segment conglomerates get independent growth and gross margin curves per division
- Single-product companies use one segment

---

## Output

### Dashboard
![Dashboard](screenshots/dashboard.png)

### Projected Financials
![Financials](screenshots/financials.png)

### Valuation
![Valuation](screenshots/valuation.png)

### Scenario Analysis
![Scenarios](screenshots/scenarios.png)

### Monte Carlo — Output Statistics
![MC Output](screenshots/mc-output.png)

### Monte Carlo — Histogram
![MC Histogram](screenshots/mc-histogram.png)

---

## Case Study

Here's a company I'm interested in run through the model. Full equity research write-ups live in the [Equity Research](../equity-research/) section of this portfolio.

### SanDisk

This company came to my attention for a few reasons. *Disclosure:* I am currently short this stock. 

Its explosive share appreciation following its spin-off from Western Digital is obviously quite jarring. This got me interested in how a firm selling a commoditised product can be priced like Nvidia (10x+ in under a year). 

The analysis on this company falls hand in hand with my prior knowledge in the tech and AI space. I've saved the thorough due diligence for my Equity Research GitHub section.

**Key outputs:**

| Metric | Value |
|--------|-------|
| IVPS (Base Case) | <!-- $ --> |
| Market Price at Analysis | <!-- $ --> |
| Implied Upside/Downside | <!-- % --> |
| MC Iterations | <!-- e.g. 10,000 --> |
| MC % Above Market Price | <!-- % --> |
| Bull / Base / Bear IVPS | <!-- $ / $ / $ --> |


> **[Full Write-Up →](../equity-research/)**

---

## Core Features

### Valuation Engine

The primary valuation runs on a **Cash Proxy DCF** (Net Income − CapEx + D&A). Unlevered Free Cash Flow is computed alongside as a cross-reference.

The user can switch between the two via a toggle.

Three terminal value methods are supported:

- **Gordon Growth Model** — closed-form terminal value
- **Explicit Perpetuity** — customisable year-by-year projection
- **Exit Multiple** — EV/UFCF based terminal value

The Valuation sheet outputs an intrinsic value per share (IVPS).

### Monte Carlo Simulation

This builds on the standalone Monte Carlo Simulation Engine (Project 2 in this portfolio) but takes the concept significantly further.

Where the earlier project ran independent draws from basic distributions, this implementation adds:

- **Correlated parameter draws**
- **Revenue autocorrelation**
- **Mean-reversion dynamics**
- **Phase-based shock regimes**
- **PERT, Normal, Triangular, Uniform, and Lognormal distributions**
- **12+ variables**
- **Full output statistics**

The simulation runs configurable iterations (max 50,000) and outputs the percentage of iterations that exceed the current market price — a direct probabilistic answer to "is this stock cheap?"

### Scenario Analysis

Bull/Base/Bear Scenario. You know the deal.

### Penman Reformulated Financials

The Formulas sheet implements Stephen Penman's equity valuation framework. This guy's insights are no joke.

---

## Workbook Architecture

The model is a 9-sheet workbook with clear separation of concerns:

| Sheet | Role |
|-------|------|
| **Dashboard** | VBA control centre |
| **Assumptions** | The only sheet the user modifies — all inputs, toggles, segment config, MC parameters |
| **Financials** | Multi-year projected Income Statement, Cash Flow, Working Capital, PP&E Schedule |
| **Valuation** | DCF output, equity bridge, IVPS, terminal value cross-checks, sensitivity table |
| **Scenarios** | Bull/Base/Bear with probability-weighted IVPS and sensitivity matrix |
| **MC_Output** | Monte Carlo statistics |
| **MC_Histogram** | 50-bin distribution chart of simulated IVPS outcomes |
| **Formulas** | Reference sheet — 69 formulas, Penman reformulated balance sheet, computed ratios |
| **Qualitative** | SWOT, VRIN, bear/bull thesis, probability-of-occurrence diagram |

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
