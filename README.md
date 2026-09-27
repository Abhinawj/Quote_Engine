# 🛡️ BimaGrid Kenya

**Enterprise Motor & Motorcycle Insurance Quotation Terminal**

*A broker-grade motor insurance rating and quotation engine built with Streamlit and Python.*

---

## Overview

**BimaGrid Kenya** is a quotation and tariff comparison terminal designed for the Kenyan motor insurance market.

The application enables insurance brokers, agents, and underwriting teams to configure vehicle risks, calculate indicative premiums, apply statutory charges, compare insurer quotations, and generate client-ready quotation schedules.

The platform follows a structured two-stage workflow:

1. **Vehicle & Policy Configuration**
2. **Quotation Comparison & Executive Dashboard**

BimaGrid Kenya supports both **Comprehensive** and **Third Party Only (TPO)** insurance configurations, with separate underwriting rules for private vehicles and motorcycles.

> **Important:** BimaGrid Kenya is a quotation and rating support tool. Final coverage, premium, acceptance, inspection requirements, underwriting decisions, and policy terms remain subject to the applicable insurer's underwriting rules and the prevailing Kenyan regulatory framework.

---

## Key Capabilities

### 1. Two-Stage Quotation Workflow

#### Stage 1 — Vehicle & Policy Configuration

The first stage captures the vehicle and policy configuration required for quotation.

Key features include:

* Vehicle Category selection
* Make selection
* Model selection
* Engine / classification information
* Vehicle year
* Automatic vehicle valuation lookup
* Benchmark valuation based on the 2025 vehicle value
* Depreciation calculation
* Calculated Sum Insured
* Policy type selection
* Custom tariff/rate configuration
* Coverage extension selection
* Vehicle-specific underwriting checks

Vehicle information is populated dynamically from the backend vehicle valuation dataset.

#### Stage 2 — Quotation Comparison Dashboard

The second stage presents the calculated quotations in a consolidated comparison view.

The dashboard provides:

* Multi-insurer quotation comparison
* Gross premium comparison
* Statutory charges
* Total client outlay
* Coverage summary
* Deductible/excess information
* Underwriting conditions
* Comparison from lowest to highest calculated outlay
* Client communication options
* Excel quotation schedule generation
* WhatsApp briefing generation

---

# Motor Insurance Configuration

## Comprehensive Insurance

Comprehensive quotation calculations support:

* Vehicle Sum Insured
* Base premium rate
* Minimum premium checks
* Statutory charges
* Applicable policy extensions
* Excess/deductible requirements
* Anti-theft / tracking requirements
* Vehicle-age restrictions
* Additional service charges where applicable

Supported extensions may include:

* Windscreen Cover
* Audio / Sound System Cover
* Excess Protector
* Political Violence & Terrorism (PVT)
* Courtesy Car
* Medical Expenses
* Roadside Rescue

The availability and pricing of individual extensions remain subject to the configured insurer rules.

---

## Third Party Only (TPO)

TPO quotations are handled separately from Comprehensive insurance.

The TPO calculation framework includes:

* Statutory base premium
* Applicable levies
* Stamp duty
* Applicable service charges
* Total client outlay

Own-damage extensions are automatically suppressed for TPO quotations, including:

* Windscreen
* Sound System
* PVT
* Excess Protector
* Courtesy Car

The application also prevents Comprehensive-specific underwriting requirements from being applied to TPO quotations.

---

# Vehicle Age Rules

Vehicle age is used as part of the quotation validation process.

Vehicles exceeding the configured age threshold can be restricted to TPO based on the application's underwriting rules.

Current configuration:

> Vehicles older than **15 years** are restricted to TPO.

This rule should be treated as an application configuration and reviewed whenever insurer or regulatory underwriting requirements change.

---

# Motorcycle & Bodaboda Underwriting

BimaGrid Kenya maintains a separate underwriting flow for motorcycles and Bodaboda risks.

Motorcycle-specific functionality includes:

* Motorcycle vehicle categories
* Make and model selection
* Engine/classification details
* Motorcycle valuation
* Depreciated Sum Insured
* Motorcycle premium calculation
* Motorcycle-specific minimum premium
* Bike alarm requirements
* Pillion Passenger Personal Accident considerations
* Motorcycle roadside/recovery services

Passenger-car-specific extensions are not presented for motorcycles.

Examples include:

* Windscreen Cover
* Car Sound System Cover
* Courtesy Car

---

# Valuation & Depreciation Engine

Vehicle valuations are maintained in the backend vehicle dataset.

Each vehicle model has a benchmark:

```text
Base Value 2025 (KES)
```

The application calculates the Sum Insured using the configured depreciation schedule.

### Calculation

```text
Sum Insured =
Base Value 2025 × (1 - Depreciation Rate)
```

### Current Depreciation Schedule

| Vehicle Year | Years Elapsed | Depreciation |
| ------------ | ------------: | -----------: |
| 2025         |             0 |           0% |
| 2024         |             1 |          10% |
