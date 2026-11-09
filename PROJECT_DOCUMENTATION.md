# OmniShield Executive Summary Dashboard

## Comprehensive System Architecture & Engineering Documentation

---

### Table of Contents

1. [Executive Summary & Purpose](#1-executive-summary--purpose)
2. [High-Level Architecture](#2-high-level-architecture)
3. [Technology Stack](#3-technology-stack)
4. [End-to-End Data Pipeline](#4-end-to-end-data-pipeline)
5. [DuckDB Embedded Analytical Engine](#5-duckdb-embedded-analytical-engine)
6. [AI Actuarial & Underwriting Intelligence](#6-ai-actuarial--underwriting-intelligence)
7. [Tenure, Maturity, and Statistical Credibility](#7-tenure-maturity-and-statistical-credibility)
8. [Report Structure: Normal vs. Debug Mode](#8-report-structure-normal-vs-debug-mode)
9. [Comprehensive Tab Directory](#9-comprehensive-tab-directory)
10. [Actuarial Rules & Mathematical Formulas](#10-actuarial-rules--mathematical-formulas)
11. [Configuration & Environment Setup](#11-configuration--environment-setup)
12. [Operational Commands & CLI Guide](#12-operational-commands--cli-guide)
13. [Data Governance & Privacy Security](#13-data-governance--privacy-security)

---

### 1. Executive Summary & Purpose

The **OmniShield Executive Summary Dashboard** is an enterprise-grade automated reporting, actuarial transformation, and underwriting intelligence system. It ingests transactional warranty, service contract, cancellation, and claims data directly from production MongoDB data marts, transforms over **230,000+ contracts and 20,000+ claims** in memory using an embedded **DuckDB analytical engine**, and generates executive-ready, audit-compliant Microsoft Excel workbooks (`.xlsx`).

#### Key Capabilities

- **High-Performance Analytics**: Computes multi-period financial metrics (Current Month, YTD, Rolling 12-Month, Prior Calendar Year, Inception-to-Date) across multiple dimensions (Agents, Dealers, Products, Loss Codes, Vehicle Makes) in under 2 seconds via vectorized DuckDB SQL.
- **AI Actuarial Agent**: Employs Google Gemini (`gemini-3.8-flash`) and OpenAI (`gpt-4o-mini`) with a 100% deterministic mathematical fallback to deliver objective, factual executive summaries ($< 1$ paragraph) contextualized by program tenure and reserve earning maturity.
- **Two Operating Modes**:
  - **Normal Mode**: Standard corporate executive distribution workbook with sanitized samples and locked production formatting.
  - **Debug / Deep-Dive Mode**: Activated via `.env` flags (`DEBUG_AGENT_NAME` or `DEBUG_DEALER_NAME`). Prepends three specialized AI Underwriting tabs, unlocks 100% unsampled transactional audits, and exposes raw reconciliation math without disturbing regular production tab orders.

---

### 2. High-Level Architecture

```mermaid
flowchart TD
    subgraph Data Sources
        M1[(MongoDB: DwContract)]
        M2[(MongoDB: DwCancellation)]
        M3[(MongoDB: DwClaim)]
        M4[(MongoDB: PipelineAudits)]
    end

    subgraph Extraction & Ingestion
        Repo[DataRepository & MongoService]
        Sanitizer[Privacy Redaction & Stream Normalizer]
    end

    subgraph Analytical Processing
        DuckDB[(DuckDB In-Memory OLAP Engine)]
        SQL[Vectorized Actuarial SQL Aggregations]
        Model[ReportModel Synthesizer]
    end

    subgraph AI Intelligence Layer
        AiService[AiAnalystService]
        Gemini[Google Gemini API / gemini-3.8-flash]
        OpenAI[OpenAI API / gpt-4o-mini]
        Deterministic[Deterministic Actuarial Math Fallback]
    end

    subgraph Presentation & Delivery
        Excel[ExcelService & Worksheet Renderers]
        NormalWB[Production Executive Workbook]
        DebugWB[AI Deep-Audit Debug Workbook]
        SendGrid[SendGrid Email Dispatcher]
    end

    M1 & M2 & M3 & M4 --> Repo
    Repo --> Sanitizer
    Sanitizer --> DuckDB
    DuckDB --> SQL
    SQL --> Model
    Model --> AiService
    AiService --> Gemini & OpenAI & Deterministic
    Gemini & OpenAI & Deterministic --> AiService
    Model & AiService --> Excel
    Excel --> NormalWB & DebugWB
    NormalWB -.-> SendGrid
```

---

### 3. Technology Stack

| Layer                      | Component                           | Description                                                                                           |
| :------------------------- | :---------------------------------- | :---------------------------------------------------------------------------------------------------- |
| **Runtime & Language**     | Node.js (v20+) / TypeScript (v5.5+) | Strictly typed application with strict null checking and ES2022 targets.                              |
| **Primary Data Source**    | MongoDB (v7+)                       | Read-only connection to production reporting marts (`DwContract`, `DwCancellation`, `DwClaim`).       |
| **Analytical OLAP Engine** | DuckDB (v1.4+)                      | In-memory columnar SQL database executing complex group-bys, window functions, and reserve schedules. |
| **AI Intelligence**        | Google Gemini / OpenAI              | `gemini-3.8-flash` (via native fetch) and `gpt-4o-mini` (via OpenAI SDK).                             |
| **Workbook Generation**    | ExcelJS (v4.4+)                     | Vectorized workbook construction, conditional formatting data bars, custom palettes, and formulas.    |
| **Email Distribution**     | SendGrid Mail API                   | Automated transactional email delivery of generated reports with attachment handling.                 |

---

### 4. End-to-End Data Pipeline

The generation pipeline executes in four coordinated phases:

```mermaid
sequenceDiagram
    autonumber
    participant App as Runner (src/index.ts)
    participant Mongo as MongoDB Data Mart
    participant Duck as DuckDB Engine
    participant AI as AI Analyst Service
    participant XLS as Excel Service
    participant Output as File System / Email

    App->>Mongo: Stream raw contracts, cancellations, claims, and audits
    Mongo-->>App: Streamed document chunks (230k+ contracts, 20k+ claims)
    App->>Duck: Load normalized tables (contracts, cancellations, claims)
    App->>Duck: Execute actuarial SQL transformations
    Duck-->>App: Structured ReportModel (Scorecards, Dimensions, Trends)
    alt Debug Mode Active (DEBUG_AGENT_NAME / DEBUG_DEALER_NAME)
        App->>AI: Request Executive & Portfolio Overviews with Tenure context
        AI-->>App: Factual, non-emotional actuarial summary paragraphs
        App->>XLS: Generate AI Tabs (Portfolio, Analyst Lab, Underwriting Intel)
    end
    App->>XLS: Generate Production Sheets (Executive, Trends, Dimensions, Loss Codes)
    XLS-->>Output: Write .xlsx file to output/ directory
    opt Email Enabled
        Output->>Output: SendGrid SMTP Dispatch
    end
```

---

### 5. DuckDB Embedded Analytical Engine

Historically, data processing in Node.js applications utilized in-memory JavaScript `Map` and `Array` operations, which caused high garbage collection overhead and potential `Out Of Memory` (OOM) crashes when processing hundreds of thousands of records.

This project utilizes **DuckDB**, an embedded columnar SQL OLAP database:

#### How DuckDB Works in the Pipeline:

1. **Zero-Network In-Memory Tables**:
   Three tables are created in DuckDB: `contracts`, `cancellations`, and `claims`.
2. **Dynamic Reserve Earning**:
   Contract earning curves (Rule of 78, Straight Line, or custom actuarial schedules) are computed directly in DuckDB SQL using elapsed months between `effectiveDate` and `asOfDate`.
3. **Multi-Horizon Slicing**:
   A single SQL query evaluates multi-period boundaries using conditional aggregation:
   ```sql
   SUM(CASE WHEN activityDate BETWEEN ? AND ? THEN reserveAmount ELSE 0 END) as currentMonthReserve,
   SUM(CASE WHEN activityDate BETWEEN ? AND ? THEN reserveAmount ELSE 0 END) as ytdReserve,
   SUM(CASE WHEN activityDate BETWEEN ? AND ? THEN reserveAmount ELSE 0 END) as rolling12Reserve,
   SUM(reserveAmount) as itdReserve
   ```
4. **Instant Deduplication**:
   Handles transaction snapshot deduplication (latest `Contract#` and status) and claim payment deduplication in milliseconds.

---

### 6. AI Actuarial & Underwriting Intelligence

The system includes a dedicated actuarial AI service (`AiAnalystService`) designed to deliver **strictly objective, factual, non-emotional summaries** ($< 1$ paragraph) for executives.

#### Dual-Provider & Offline Fallback Architecture:

1. **Google Gemini (Priority 1)**:
   - Model: `gemini-3.8-flash`
   - Configured with `thinkingConfig: { thinkingBudget: 150 }` and `maxOutputTokens: 2048` to support internal reasoning without text truncation.
   - Low latency, high throughput, and cost-effective.
2. **OpenAI (Priority 2)**:
   - Model: `gpt-4o-mini`
   - Active if `OPENAI_API_KEY` is present.
3. **Deterministic Actuarial Generator (Priority 3 - 100% Offline)**:
   - If no API key is provided, or if network connectivity is offline, the system generates an exact mathematical paragraph derived directly from DuckDB metrics.

---

### 7. Tenure, Maturity, and Statistical Credibility

A central principle of the actuarial intelligence engine is **time frame and program maturity**:

> _"An agent with us for 5 years with an 85% loss ratio is completely different from a new agent with an 85% loss ratio who just started 6 months ago."_

```mermaid
flowchart TD
    Entity[Target Agent or Dealer] --> Calc[Compute Program Tenure from Earliest Contract]
    Calc --> T1{Tenure & Volume}

    T1 -->|Tenure >= 2 Years & Contracts >= 50| Seasoned[Seasoned Program: High Statistical Credibility]
    T1 -->|Tenure 1 - 2 Years| Developing[Developing Program: Moderate Credibility, Reserves Earning]
    T1 -->|Tenure < 1 Year| Immature[Immature Cohort: Low Credibility, High Unearned Reserve]

    Seasoned --> S1{ITD Loss Ratio}
    S1 -->|> 75%| Alert["Seasoned Adverse Experience (Pay Attention)"]
    S1 -->|<= 65%| Good["Seasoned Profitable (<65%)"]
    S1 -->|65% - 75%| Warning["Seasoned Moderate Margin"]

    Immature --> I1["Immature Cohort: Timing Volatility (Do NOT Overreact)"]
```

#### Actuarial Classifications:

1. **Seasoned Program ($\ge 2$ Years)**:
   - **Reserve Earned**: Typically $>80\text{–}95\%$ of written reserves are earned.
   - **Actuarial Credibility**: High statistical significance. If Inception-to-Date (ITD) loss ratio is unfavorable ($>75\%$), it represents seasoned adverse underwriting performance that **requires executive attention**.
2. **Developing Program ($1\text{–}2$ Years)**:
   - **Reserve Earned**: Typically $40\text{–}75\%$ earned.
   - **Actuarial Credibility**: Moderate. Trends are developing; monitor reserve velocity.
3. **Emerging / Unseasoned Program ($< 1$ Year)**:
   - **Reserve Earned**: Most reserves remain unearned ($< 35\%$ earned).
   - **Actuarial Credibility**: Low. Early claim spikes are common due to small-sample timing volatility and first-year vehicle break-in claims rather than structural program underpricing.

---

### 8. Report Structure: Normal vs. Debug Mode

The dashboard automatically configures its tab layout based on `.env` configuration:

```mermaid
graph TD
    subgraph Mode Detection in excel.service.ts
        ENV{DEBUG_AGENT_NAME or DEBUG_DEALER_NAME set?}
    end

    ENV -->|NO: Normal Production Mode| Normal[Normal Mode: 11 Standard Tabs]
    ENV -->|YES: Debug Mode Active| Debug[Debug Mode: 15 Tabs]

    subgraph Normal Mode Order
        Normal --> N1[1. Executive Summary]
        Normal --> N2[2. Rolling 12-Month Performance]
        Normal --> N3[3. Contract Activity - 3% Sample]
        Normal --> N4[4. Claims Activity - 3% Sample]
        Normal --> N5[5. Monthly Trends]
        Normal --> N6[6. Agent Performance]
        Normal --> N7[7. Dealer Performance]
        Normal --> N8[8. Product Performance]
        Normal --> N9[9. Loss Code Dashboard]
        Normal --> N10[10. Data Quality Audit]
        Normal --> N11[11. Definitions]
    end

    subgraph Debug Mode Order
        Debug --> D1["1. AI Portfolio Summary (Macro Enterprise)"]
        Debug --> D2["2. AI Analyst Lab (Target Overview & Loss Codes)"]
        Debug --> D3["3. AI Underwriting Intel (Sub-Dealers, TTL, Pricing)"]
        Debug --> D4[4. Executive Summary]
        Debug --> D5[...]
        Debug --> D6[Contract Activity - 100% Unsampled Target]
        Debug --> D7[Claims Activity - 100% Unsampled Target]
        Debug --> D8[15. Debug Math - Raw Verification]
    end
```

---

### 9. Comprehensive Tab Directory

#### 1. AI Portfolio Summary (`tab_ai_portfolio_summary`) — _Debug Only_

- **Enterprise Macro Scorecard**: Multi-horizon financial comparison across Current Month, YTD, Rolling 12, Prior Year, and Inception to Date.
- **Agent Risk & Volume Distribution**: Ranks top enterprise agents with **`Tenure & Active Since`**, **`Program Maturity`**, and **`Actuarial Credibility & Status`**.
- **Top 5 Portfolio Loss Exposures**: Company-wide claim drivers and dollar volumes.

#### 2. AI Analyst Lab (`tab_ai_agent_debug`) — _Debug Only_

- **Executive AI Overview Box**: Concise, factual summary ($< 1$ paragraph) explaining target tenure, earned reserve status, and statistical credibility.
- **Target Performance Scorecard**: Side-by-side comparison of Rolling 12-Month vs. Inception-to-Date metrics.
- **Top Loss Code Exposures**: Primary mechanical failure codes driving target claims.

#### 3. AI Underwriting Intel (`tab_ai_underwriting_intel`) — _Debug Only_

- **Section 1: Dealer Cohort Concentration**: Evaluates sub-dealers under a target agent, showing each dealer's tenure, active date, and credibility status (e.g. `Seasoned Adverse (Pay Attention)` vs. `Developing Cohort`).
- **Section 2: Time-to-Loss Velocity**: Analyzes claim emergence across four in-service windows ($0\text{–}90$ days, $91\text{–}180$ days, $181\text{–}365$ days, $365+$ days).
- **Section 3: Vehicle Make Distribution**: Ranks claims by manufacturer make and primary failing components.
- **Section 4: Rate Card Pricing Adequacy**: Compares target reserve and admin fees against company-wide portfolio benchmarks.

#### 4. Executive Summary (`tab_executive_dashboard`)

- Primary enterprise dashboard with macro KPIs, underwriting profit bars, and loss ratio alerts.

#### 5. Rolling 12-Month Performance (`tab_rolling_12_performance`)

- Trailing 12-month performance categorized by dealer, agent, and product line.

#### 6. Contract Activity (`tab_contract_activity`)

- Transactional contract log. In Normal Mode, outputs a sanitized 3% audit sample. In Debug Mode, outputs 100% of matching transactions for the debug target with un-redacted financial columns.

#### 7. Claims Activity (`tab_claim_activity`)

- Detailed claim payment log. Outputs 100% of transactions for the debug target in Debug Mode.

#### 8. Monthly Trends (`tab_monthly_trends`)

- Month-over-month historical progression of contracts written, earned reserves, claims paid, and loss ratios.

#### 9. Agent Performance (`tab_agent_performance`)

- Enterprise agent ranking sorted by volume, net reserve, and loss ratio.

#### 10. Dealer Performance (`tab_dealer_performance`)

- Top dealers ranked by Rolling 12-Month Net Written Reserve with conditional color-coded loss ratio bars.

#### 11. Product Performance (`tab_product_performance`)

- Performance metrics segmented by coverage types (VSC, GAP, Tire & Wheel, Ancillary).

#### 12. Loss Code Dashboard (`tab_loss_code_dashboard`)

- Detailed breakdown of all claim loss codes, component descriptions, paid claim shares, and top vehicle makes.

#### 13. Data Quality Audit (`tab_data_quality`)

- Pipeline audit validating MongoDB source integrity, record counts, and reconciliation variance.

#### 14. Definitions (`tab_definitions`)

- Formal corporate metric definitions, formulas, and exclusion rules.

#### 15. Debug Math (`tab_debug_math`) — _Debug Only_

- Granular mathematical proof reconciling every reserve, admin, and claim dollar for the debug target.

---

### 10. Actuarial Rules & Mathematical Formulas

Every figure in the dashboard adheres to controlled actuarial formulas:

$$\text{Net Written Reserve} = \text{Gross Written Reserve} - \text{Cancelled Reserve}$$

$$\text{Net Admin} = \text{Gross Written Admin} - \text{Cancelled Admin}$$

$$\text{Earned Reserve} = \sum_{i=1}^{N} \Big( \text{Reserve}_i \times \text{EarningFactor}(t_i, \text{Schedule}_i) \Big)$$

$$\text{Unearned Reserve} = \text{Net Written Reserve} - \text{Earned Reserve}$$

$$\text{Earned Loss Ratio} = \frac{\text{Claims Paid}}{\text{Earned Reserve}}$$

$$\text{Underwriting Profit} = \text{Earned Reserve} - \text{Claims Paid}$$

$$\text{Program Tenure (Years)} = \frac{\text{AsOfDate} - \text{EarliestContractDate}}{365.25 \text{ days}}$$

---

### 11. Configuration & Environment Setup

Configure application behavior via the `.env` file:

```dotenv
# MongoDB Connection
MONGODB_URI=mongodb+srv://user:pass@cluster.mongodb.net/?retryWrites=true&w=majority
MONGODB_CONTRACT_DB=PROD_OmniShield_DwContract
MONGODB_CONTRACT_COLLECTION=DwContract
MONGODB_CANCELLATION_DB=PROD_OmniShield_DwCancellation
MONGODB_CANCELLATION_COLLECTION=DwCancellation
MONGODB_CLAIM_DB=PROD_OmniShield_DwClaim
MONGODB_CLAIM_COLLECTION=DwClaim
MONGODB_AUDIT_DB=PROD_OmniShield_PipelineAudits
MONGODB_AUDIT_COLLECTION=PipelineAudits

# Debug Mode (Comment out for Normal Production Distribution)
DEBUG_AGENT_NAME=EXW
# DEBUG_DEALER_NAME=Power Motors

# AI Analyst Configuration
GEMINI_API_KEY=AIzaSy...your_gemini_api_key
GEMINI_MODEL=gemini-3.8-flash
# OPENAI_API_KEY=sk-...
# OPENAI_MODEL=gpt-4o-mini

# Thresholds
REPORT_WARNING_LOSS_RATIO=0.65
REPORT_HIGH_LOSS_RATIO=0.80

# SendGrid Email (Optional)
EMAIL_ENABLED=false
SENDGRID_API_KEY=SG...
EMAIL_FROM=reports@omnishield.com
EMAIL_TO=executive@omnishield.com
```

---

### 12. Operational Commands & CLI Guide

#### Run the Production Report

```bash
npm start
```

#### Run Report for a Specific Closed Month

```bash
npm run report -- --as-of 2026-10-31
```

#### Run Pipeline Test Script for an Agent or Dealer

```bash
npx ts-node -T scripts/debug/test-pipeline.ts EXW
```

#### Validate TypeScript and Code Quality

```bash
npm run check
```

---

### 13. Data Governance & Privacy Security

To protect sensitive consumer and dealer data, the pipeline enforces privacy restrictions:

- **Consumer PII Redacted**: Customer names, phone numbers, email addresses, and street addresses are excluded at the MongoDB cursor level.
- **VIN Exclusion**: Vehicle Identification Numbers (VINs) are stripped during extraction.
- **Financial Segregation**: Internal dealer packs and sales commissions are excluded from underwriting calculations.
- **3% Production Sampling**: In Normal Mode, raw transactional audit tabs sample 3% of transactions to prevent file bloat while maintaining audit reproducibility. In Debug Mode, only the targeted debugging entity is unsampled.
