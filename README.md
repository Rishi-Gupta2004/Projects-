# Rishi Gupta's Finance Project Portfolio

Welcome. This is a collection of projects I've built as a student of Finance and more specifically equity valuation. 
This is mostly Excel VBA tools that solve real problems I've encountered in financial modelling and analysis - some python as well.

I'm just someone who loves the stock picking game i suppose.

Shout out to some of the individuals and firms who made me fall in love with the game (admiration for many more).

- Aswath Damodaran
- Martin Shkreli
- Mark Meldrum 

> 📄 **[About Me →](about-me/index.html)**  
> A bit more on who I am if your interested.

---

## Projects

### 1. [Rishi's Equity Model](Equity-Analysis-Workbook/)
<-- Date: 03/2026 -->

A comprehensive, company-agnostic equity valuation framework built from applying a lot of various experiences into a streamlined VBA powered workbook. 
I'm super proud of how this ones come. Currently on version **7.2** as of when your reading this.

Combines the following:
- A dual-basis DCF (Net Income + UFCF)
- Monte Carlo simulation with correlated draws and autocorrelation (evolving the concepts from the my previous standalone Monte Carlo project through integration)
- Three-scenario analysis, segment-level revenue modelling
- Asset-heavy model toggles (allows for a goods vs services switchup)
- Penman reformulated financials — all driven from a single Assumptions sheet via 275+ named ranges. Currently demonstrated with Sandisk Corporation.

**Tech:** Excel VBA 
- Named Range Architecture
- Monte Carlo (PERT/Normal/Correlated)
- DCF Valuation
- Scenario Analysis
- Penman Reformulation 
- Qualitative analysis 

---

### 2. [Monte Carlo Simulation Engine](monte-carlo-simulation/)
<-- Date: 10/2025 -->

A simulation macro that models project NPV under uncertainty. 

Inputs follow discrete, triangular, or uniform distributions. The tool runs the simulations, generates a histogram of outcomes, and surfaces the key stats. 

Monte Carlo sims blow my mind and has given me a true apreciation for maths.

**Tech:** Excel VBA 
- Statistical Distributions
- Histogram Generation
- Simulation Logic

---

### 3. [Discounted Cash Flow Model](dcf-model/)
<-- Date: 09/2025 -->

A VBA-powered DCF tool with two modes: a dynamic template generator that builds formatted DCF spreadsheets on the fly, and a standalone userform that computes NPV and IRR from direct inputs. Built to make valuation work faster and less error-prone.

**Tech:** Excel VBA 
- Subroutines
- Dynamic Cell Formatting
- Userforms

---

*More projects on the way.*
