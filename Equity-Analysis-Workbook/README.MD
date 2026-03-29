# Rishi's Equity Model — Excel VBA

## What I Learned

This project really encapuslates everytihng I've learn't up untill this point. 

This is a comprehensive financial modelling tool designed to value any public company through multiple valuation methodologies, Monte Carlo simulation, and scenario analysis. 

Currently demonstrated with **Sandisk Corporation (SNDK)**.

The model is architected so that switching to a new company only requires updating the Assumptions sheet. 

All downstream sheets — Financials, Valuation, Scenarios, Monte Carlo — recompute automatically via 275+ named ranges that act as a clean API between the VBA engine and the spreadsheet layer.

---

## How It Works

The user populates a single Assumptions sheet with the target company's financials, valuation parameters, and uncertainty ranges. 

From there, macro buttons on the Dashboard drive the entire workflow  

Financials are projected, running the DCF, stress-tested through scenarios, and simulated in thousands of possible outcomes via Monte Carlo.

The model adapts to different company types through binary toggles: 
- Asset-heavy manufacturers get a fixed+variable cost structure with a PP&E schedule, while asset-light businesses use simpler margin-based projections.
- Multi-segment conglomerates get independent growth and gross margin curves per division. 
- Single-product companies use one segment.

---

## Screenshots


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

The primary valuation runs on a **Cash Proxy DCF** (Net Income − CapEx + D&A)

Unlevered Free Cash Flow is computed alongside as a cross-reference however cash proxy feeds exclusively into both monte carlo and scnario analysis. 

The user can switch between the two via a toggle. 

Three terminal value methods are supported and cross-checked against each other on every run:
- **Gordon Growth Model** — closed-form terminal value
- **Explicit Perpetuity** — 52-year year-by-year projection (deafult)
- **Exit Multiple** — EV/UFCF based terminal value

The Valuation sheet outputs an intrinsic value per share (IVPS), an upside/downside signal against the current market price, and a full **Discount Rate vs. Terminal Growth Rate sensitivity table** with colour-coded over/undervaluation.

### Monte Carlo Simulation

This builds on the standalone Monte Carlo Simulation Engine (Project 2 in this portfolio) but takes the concept significantly further. Where the earlier project ran independent draws from basic distributions, this implementation adds:

- **Correlated parameter draws** — gross margin and revenue growth are linked, as are growth and discount rate, so the simulation reflects real-world co-movement rather than treating variables as independent
- **AR(1) autocorrelation** — growth rates exhibit year-to-year persistence, preventing unrealistic whipsawing between extreme values (expecially good for tech companies)
- **Mean-reversion dynamics** — extreme growth rates pull back toward the distribution centre in terminal years
- **PERT, Normal, Triangular, Uniform, and Lognormal distributions** — each variable can use whichever distribution best fits its uncertainty profile
- **12+ variables** randomised simultaneously per iteration, including two custom variable slots for company-specific factors
- **Full output statistics** — mean, median, P5–P95 percentiles, min/max, and a 50-bin histogram visualisation

The simulation runs configurable iterations (default 1,000, max 50,000) and outputs the percentage of iterations that exceed the current market price 
- A direct probabilistic answer to "is this stock cheap?"

### Scenario Analysis

Bull/Base/Bear scenario analysis with user-defined adjustments to revenue growth, gross margin, tax rate, discount rate, terminal growth rate, and CapEx. 

Each scenario can optionally include a **time-decay rate**, where the adjustment fades toward the base case over the forecast horizon. 

The output is a probability-weighted IVPS alongside the individual scenario valuations.

### Segment-Level Revenue Modelling

Revenue is projected at the segment level (up to 4 segments) with independent year-by-year growth rates and gross margin curves per segment. 

This allows the model to capture businesses with distinct divisions growing at different rates.

### Asset-Heavy Cost Structure

When toggled on, costs split into **fixed dollar amounts** (inflated annually) and **variable percentages of revenue**, capturing the operating leverage dynamics inherent in capital-intensive businesses. 

A PP&E schedule drives depreciation from the asset base rather than as a simple percentage of revenue, and CapEx decomposes into maintenance (sustaining the asset base) and growth (funding incremental revenue) components.

### Penman Reformulated Financials

The Formulas sheet implements Stephen Penman's equity valuation framework. I took a class in the start of my thrid year based soley on this blokes work and it is nothing short of fascinating  

— Reformulating the balance sheet into Operating Assets/Liabilities and Financial Assets/Obligations 
Return decomposition through RNOA, FLEV, SPREAD, and ROCE. 

This sits alongside 69 numbered formulas covering everything from residual income to DuPont decomposition.

### Qualitative Framework

A dedicated Qualitative sheet provides structured space for SWOT analysis, VRIN competitive advantage assessment, and bear/bull case narratives with a probability-of-occurrence framework.

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
| **Qualitative** | SWOT, VRIN, bear/bull thesis, probability-of-occurrence diagram |

Data flows from Assumptions → Financials → Valuation → Scenarios / Monte Carlo. 

Named ranges create an abstraction layer so the VBA never hardcodes cell addresses.

---

## VBA Modules

| Module | Purpose |
|--------|---------|
| **Engine** | Core forecast builder and NPV calculator — handles segment revenue rollup, asset-heavy/light cost routing, NOL carryforward, debt amortisation, working capital |
| **Valuation** | Writes the DCF valuation sheet — equity bridge, terminal value cross-checks, sensitivity grid |
| **MonteCarlo** | Simulation engine — correlated draws, autocorrelation, mean-reversion, histogram generation |
| **Financials** | Writes the projected financial statements to the Financials sheet |
| **Scenarios** | Bull/Base/Bear scenario runner with time-decaying adjustments |
| **Distributions** | Statistical distribution samplers — PERT, Normal, Triangular, Uniform, Lognormal, plus QuickSort and percentile functions |
| **Constants** | Public types (ModelInputs, YearFinancials, ScenarioConfig, MCDistParams), design system colour palette, sheet name constants |
| **Helpers** | Input readers, validators, named range utilities, WACC calculator, share dilution logic |
| **modFormatStyles** | Shared design system — consistent formatting across all output sheets |
| **RunALL** | Master orchestrator — individual entry points and full-model execution |
| **AssetHeavy** | Toggle for switching between asset-heavy and asset-light cost structures |
| **ScenarioHelper** | Setup utility for scenario named ranges |
| **Movement** | Sheet navigation (Dashboard, Assumptions, Qualitative) |
| **NoteEditor** | Cell comment management — delete and restore notes |

---

## Tech

**Excel VBA** 
- Named Range Architecture (275+) 
- Statistical Distributions (PERT, Normal, Triangular, Uniform, Lognormal) 
- Monte Carlo Simulation with Cross-Correlation & Autocorrelation 
- DCF Valuation 
- Scenario Analysis 
- Penman Reformulated Financials 
- Asset-Heavy Cost Modelling 
- Segment-Level Revenue Projection 
- Design System (modFormatStyles)

---

