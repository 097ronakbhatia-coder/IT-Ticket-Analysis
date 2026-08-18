# 🎫 IT Ticket Analysis Dashboard

### End-to-End IT Service Desk Analytics | 2016 – 2020

*Turning 97,498 raw support tickets into a data-driven decision framework for IT operations.*

![Excel](https://img.shields.io/badge/Tool-Excel-217346?style=flat-square&logo=microsoftexcel&logoColor=white)
![Data Analysis](https://img.shields.io/badge/Focus-Data%20Analysis-blue?style=flat-square)
![Dashboard](https://img.shields.io/badge/Deliverable-Interactive%20Dashboard-orange?style=flat-square)
![Status](https://img.shields.io/badge/Status-Completed-brightgreen?style=flat-square)


---

## 📑 Table of Contents

- [📌 Overview](#-overview)
- [🖼️ Dashboard Preview](#-dashboard-preview)
- [🔢 Key Metrics at a Glance](#-key-metrics-at-a-glance)
- [📂 Repository Contents](#-repository-contents)
- [🗃️ Dataset Description](#-dataset-description)
- [🧹 Data Quality Findings](#-data-quality-findings)
- [🔬 Analytical Approach](#-analytical-approach)
- [💡 Key Insights](#-key-insights)
  - [1. Ticket Volume & Growth](#1-ticket-volume--growth)
  - [2. Resolution Time — The Core Bottleneck](#2-resolution-time--the-core-bottleneck)
  - [3. Customer Satisfaction](#3-customer-satisfaction)
  - [4. Request Category Performance](#4-request-category-performance)
  - [5. Agent Performance](#5-agent-performance)
  - [6. Seasonality & Demand Patterns](#6-seasonality--demand-patterns)
  - [7. Demographic Segmentation](#7-demographic-segmentation)
- [🎯 Strategic Recommendations](#-strategic-recommendations)
- [❓ Business Questions Answered](#-business-questions-answered)
- [🚀 How to Use This Project](#-how-to-use-this-project)
- [🛠️ Tech Stack](#-tech-stack)
- [👤 Author](#-author)


---

## 📌 Overview

This project delivers a complete **IT Service Desk analytics engagement** — from raw data auditing to an executive-ready interactive dashboard and a full business-question deliverable — built entirely in **Microsoft Excel**.

The analysis covers **97,498 IT support tickets** logged by **50 IT agents** over a **5-year period (2016–2020)**, and answers a structured set of objective and subjective business questions that a real IT operations / investment committee would ask before allocating budget toward **hiring, training, or software upgrades**.

The central finding of the analysis is that process and tooling constraints appear to be more significant operational bottlenecks than overall staffing levels. Ticket volume increased by 123% between 2016 and 2020, while average resolution time remained broadly stable. This indicates an opportunity to improve ticket routing, automation, knowledge management, and workflow efficiency before relying primarily on additional headcount.

---

## 🖼️ Dashboard Preview

![IT Ticket Analysis Dashboard](/screenshot/IT_Ticket_Dashboard)

*A single-page, filter-driven Excel dashboard covering ticket volume, category mix, severity, priority, resolution time, satisfaction, and agent workload.*


---

## 🔢 Key Metrics at a Glance

| Metric | Value | Why It Matters |
|---|---:|---|
| **Total Tickets** | 97,498 | Full 5-year workload baseline |
| **IT Agents** | 50 | Total support team size |
| **Average Daily Volume** | 53.37 tickets/day | Overall daily ticket workload |
| **Average Resolution Time** | 4.55 days | Core service-efficiency KPI |
| **Average Satisfaction** | 4.10 / 5 | Overall customer satisfaction |
| **Unassigned Tickets** | 29,410 (30.2%) | Significant routing and prioritization gap |
| **Ticket Volume Growth** | +123% | Growth from 2016 to 2020 |
| **Average Agent Age** | 40.70 years | Team demographic benchmark |

---

## 📂 Repository Contents

| File | Description |
|---|---|
| `data/IT_Tickets.xlsx` | Raw source dataset containing the Tickets and IT Agents sheets |
| `excel/IT_Ticket_Analysis.xlsx` | Cleaned and modeled Excel workbook containing KPI, Pivot, and Dashboard sheets |
| `screenshots/IT_Ticket_Dashboard.png` | Static preview of the Excel dashboard |
| `documentation/IT_Ticket_Business_Questions.docx` | Full business-question analysis with data-backed answers and recommendations |
| `presentation/IT_Ticket_Analysis_Presentation.pptx` | Executive presentation summarizing methodology, findings, and recommendations |
| `README.md` | Project overview, methodology, insights, and navigation |

---

## 🗃️ Dataset Description

The raw data is distributed across **two related sheets**, joined via `Agent ID`.

**`Tickets` sheet (10 columns)**
| Column | Description |
|---|---|
| `ID Ticket` | Unique ticket identifier |
| `Fecha` | Date the ticket was logged |
| `Employee ID` | Requesting employee |
| `Agent ID` | Assigned support agent (foreign key → IT Agents) |
| `Request Category` | Hardware / Login Access / Software / System |
| `Issue Type` | IT Error / IT Request |
| `Severity` | 0-Unclassified → 4-Urgent |
| `Priority` | 0-Unassigned → 3-High |
| `Resolution Time (Days)` | Days taken to close the ticket |
| `Satisfaction Rate` | Customer rating, 1–5 |

**`IT Agents` sheet (6 columns)**
| Column | Description |
|---|---|
| `Agent ID` | Primary key |
| `Full Name` | Agent name |
| `Email` | Corporate email (all agents share the domain `fp20analytics.com`) |
| `Year / Month / Day of Birth` | Used to derive agent age for demographic analysis |

- **Total attributes:** 16 (10 + 6), or **15 unique** when the shared `Agent ID` key is counted once.
- **Categorical columns:** 7 total — 5 nominal (`ID Ticket`, `Request Category`, `Issue Type`, `Full Name`, `Email`) and 2 ordinal (`Severity`, `Priority`).
- **Agent–ticket relationship** is modeled using `VLOOKUP` (`=VLOOKUP(D2,'IT Agents'!$A:$B,2,0)`) to enrich every ticket with the agent's full name for performance analysis.
- **Email domain extraction** uses `MID(C2,FIND("@",C2)+1,LEN(C2))` — confirming a single-organization deployment (100% of agents on one domain).

---

## 🧹 Data Quality Findings

A data quality audit was performed across the Tickets and IT Agents sheets.

- 🔴 **Priority:** 29,410 records (30.2%) contain the value `0 - Unassiged`, representing both a spelling inconsistency and an unassigned priority status.
- 🟠 **Severity:** 356 records (0.37%) contain `0 - Unclasified`, which contains a spelling inconsistency and represents unclassified severity.
- 🟢 **Missing Values:** No NULL/missing values were identified in either the Tickets or IT Agents dataset.

The Priority issue is the most significant data-quality and operational finding because it affects 30.2% of the ticket population.

---

## 🔬 Analytical Approach

The engagement followed a structured, four-stage analytical methodology (detailed in the accompanying presentation):

1. **Data Profiling** — attribute inventory, data types, missing/inconsistent value audit
2. **Descriptive Analysis** — volume trends, category distribution, severity/priority breakdowns
3. **Diagnostic Analysis** — resolution time by category, agent performance quadrants, satisfaction drivers
4. 4. **Correlation & Segmentation Analysis** — relationship between severity and resolution time, demographic segmentation, and comparative trend analysis

All modeling was performed natively in **Excel** using **PivotTables, PivotCharts, VLOOKUP, slicers, and dashboard linking** — no external BI tool was required.

---

## 💡 Key Insights

### 1. Ticket Volume & Growth
- Ticket volume grew from **13,051 (2016)** to **29,088 (2020)** — a **123% increase** over five years.
- The steepest single-year jump was **+35.4% in 2020**, plausibly linked to the pandemic-driven shift to remote work.
- **December 2020** recorded the highest single-month volume ever (**2,609 tickets**).
- Average daily volume rose from **35.76 → 79.42 tickets/day**, more than doubling.

### 2. Resolution Time — The Core Bottleneck
- Overall average resolution time: **4.55 days** — and it has **not improved in 5 years**, despite ticket volume more than doubling.
- This flat trend, paired with accelerating demand, is the project's central diagnostic finding: **the ticketing system is not scaling with the business.**
- **30.2% of tickets (29,410)** were recorded with no assigned priority, indicating a significant gap in ticket triage and prioritization.
- Counterintuitively, **high-priority tickets resolve faster (3.49 days)** than low-priority ones (6.01 days) — proving the priority system *works when it's actually applied*, which makes the 30% unassigned gap even more costly.

### 3. Customer Satisfaction
- Average satisfaction improved steadily from **3.98 (2016) → 4.16 (2020)**.
- However, the average **masks hidden dissatisfaction**: while 52% of tickets scored 5/5, **10.2% scored 1/5**.
- Satisfaction is a **weak proxy for efficiency** — a 1-day resolution scores 4.09, while a 10+ day resolution still scores 4.05. Agents are compensating for slow tools with better communication, not faster service.

### 4. Request Category Performance
| Category | Volume | % Share | Avg. Resolution Time |
|---|---|---|---|
| System | 39,002 | 40% | 6.62 days |
| Login Access | 29,193 | 30% | **0.31 days** (fastest) |
| Software | 19,570 | 20% | 5.24 days |
| Hardware | 9,733 | 10% | **7.63 days** (slowest — 24× longer than Login Access) |

- **Hardware** is the smallest category by volume but the **most operationally costly** per ticket.
- **System** carries the highest ticket volume and a high average resolution time, making it a major opportunity for process improvement.
- **Login Access** proves fast, automated resolution is achievable — it's the internal benchmark every other category should be measured against.

### 5. Agent Performance
- Ticket distribution is relatively balanced across the 50 agents, with approximately 1,900–2,000 tickets per agent.
- **12 of 50 agents (24%)** fall below the 3.75 satisfaction threshold — a **systemic training gap**, not isolated underperformance.
- **Alfonso Barraza** is the most critical performance case, with **3.04/5 satisfaction** and a **5.00-day average resolution time**.
- Underperforming agents collectively manage **23,310 tickets (24% of total workload)** — meaningful enough to justify structured intervention rather than termination.
- **Top performers** (e.g., Diana Rojo, 4.60 satisfaction / 3.6-day resolution) are recommended as internal mentors rather than replacing lower performers with new hires.

### 6. Seasonality & Demand Patterns
- **Peak months:** August (consistent peak since 2018), October, and December (2,609 tickets in Dec 2020 — the highest month on record).
- **Stable / low-demand month:** January — consistently the lowest-volume month every year from 2017–2020, making it the safest window for system upgrades and training.
This indicates that peak demand patterns changed over the period and should be considered in capacity planning.

### 7. Demographic Segmentation
| Age Group | Ticket Volume | Avg. Satisfaction | Avg. Resolution |
|---|---|---|---|
| 29–33 | 23,483 (highest) | 4.21 (best) | 4.43 days |
| 34–38 | 13,586 (lowest) | 3.96 (worst) | 5.01 days (slowest) |
| 39–43 | ~balanced | ~team avg | ~team avg |
| 44–48 | 23,478 (high) | 4.00 | 4.83 days |
| 49–54 | 15,603 | 4.20 | **4.11 days (fastest)** |

- The **34–38** age segment shows the lowest satisfaction and longest average resolution time among the reported age groups, making it a potential priority for further investigation and targeted coaching.
- The **49–54** segment records the fastest average resolution time among the reported age groups and could be examined for practices that may be transferable to other groups.

---

## 🎯 Strategic Recommendations

Ranked by return on investment, based on cost-benefit analysis of resolution time and satisfaction metrics:

1. **🥇 Software / Ticketing System Upgrade — Highest Priority**
   Implement automated ticket prioritization, improved routing, and category-based resolution targets.
   **Proposed target:** Reduce average resolution time from 4.55 days to below 3.0 days within 12 months.

2. **🥈 Targeted Agent Training (Not Hiring)**
   Enroll the 12 underperforming agents (24% of the team) in a structured **90-day Performance Improvement Plan**, mentored by top performers (Diana Rojo, Javier D., Galindo Guadalupe). Do **not** terminate — Removing all 12 agents immediately could create significant capacity pressure, with the analysis estimating approximately 160 additional tickets per remaining agent per month.

3. **🥉 Category-Specific Fixes**
   - **Hardware:** 3-day resolution target with automatic vendor escalation at Day 2.
   - **System:** Build a knowledge base for the top 20 recurring issues (~9,750 tickets/year addressable).
   - **Software:** Standardize and automate solutions modeled on the Login Access process (target < 2 days).

4. **📅 Capacity Planning**
   Pre-staff for confirmed surge months (**August, October, December**); run upgrades and training during the stable window (**January–February**).

5. **🚫 Do Not Prioritize Hiring Initially**
   The team handled a 123% increase in ticket volume without a proportional increase in headcount. The analysis therefore recommends prioritizing workflow, routing, automation, and training improvements before expanding the team.

> **Recommended investment priority:** Software & workflow improvements → Targeted training → Additional hiring if capacity gaps remain.
---

## ❓ Business Questions Answered

The full deliverable (`IT_Ticket_Business_Questions.docx`) contains objective and subjective business questions covering data quality, ticket trends, category performance, agent performance, investment decisions, and operational recommendations.

**Objective**
- Attribute inventory, missing/inconsistent value audit
- Average daily ticket volume and category distribution
- Ticket count per agent, email domain extraction, agent name lookups
- Issue type breakdown, average resolution time, YoY volume trend
- Average agent age, severity–resolution time correlation, categorical column count

**Subjective / Strategic**
- Should investment go toward hiring, training, or software upgrades?
- Which agents need additional training?
- Do certain request categories take longer to resolve?
- How effective are current software tools?
- How has team performance changed over time?
- Would more tech investment improve resolution time and satisfaction?
- What agent-level KPIs matter, and should any agents be let go?
- How do employee demographics affect satisfaction and outcomes?

Each question is answered with **supporting data, a chart insight, key findings, recommendations, and a conclusion** — structured the way a business analyst would present findings to an investment or operations committee.

---

## 🚀 How to Use This Project

1. **Explore the raw data** — open `IT_Tickets.xlsx` to review the source `Tickets` and `IT Agents` sheets.
2. **Review the modeled workbook** — open `IT_Ticket_Analysis.xlsx` to see the `KPI`, `Pivot`, and `Dashboard` sheets, including all slicers and PivotTables used to build the dashboard.
3. **Interact with the dashboard** — use the Year, Request Category, Priority, Severity, and Issue Type slicers inside the `Dashboard` sheet to filter the view.
4. **Read the full analysis** — open `IT_Ticket_Business_Questions.docx` for the complete Q&A write-up with recommendations.
5. **Present the findings** — use `IT_Ticket_Analysis_Presentation.pptx` for a boardroom-ready summary of methodology and conclusions.

---

## 🛠️ Tech Stack

- **Microsoft Excel** — data cleaning, formulas, PivotTables, PivotCharts, slicers, KPI analysis, and dashboard development
- **Microsoft Word** — business-question analysis and documentation
- **Microsoft PowerPoint** — executive presentation and storytelling
- **GitHub** — project documentation, version control, and portfolio presentation

---


## 👤 Author

**Ronak Bhatia**

Business Analytics | Data Analysis | Excel | SQL | Python

This project demonstrates an end-to-end analytics workflow covering data quality analysis, KPI development, Excel dashboarding, business analysis, and strategic recommendations.
