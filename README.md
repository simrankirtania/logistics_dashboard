# SwiftRoute Logistics Analytics

> End-to-end logistics analytics project using **Excel, SQL, Python, and Power BI** to turn operational data into actionable business insights.

SwiftRoute Logistics is a portfolio analytics project built around a two-year logistics dataset covering orders, drivers, hubs, and vehicles.

The project follows a practical analytics workflow:

**Raw Data → Data Quality & Validation → SQL Analysis → Statistical & Predictive Analysis → Power BI Dashboard → Business Recommendations**

Rather than using one tool for everything, each stage is designed around the type of problem it solves best. Excel is used to validate and prepare the data, SQL answers operational questions, Python tests whether observed patterns are statistically meaningful, and Power BI turns validated metrics into an executive-facing dashboard.

---

## Table of Contents

- [Project Overview](#project-overview)
- [Business Problem](#business-problem)
- [Project Objectives](#project-objectives)
- [Dataset](#dataset)
- [Analytics Workflow](#analytics-workflow)
  - [1. Excel: Data Preparation & Quality Assurance](#1-excel-data-preparation--quality-assurance)
  - [2. SQL: Operational Analysis](#2-sql-operational-analysis)
  - [3. Python: Statistical & Predictive Analysis](#3-python-statistical--predictive-analysis)
  - [4. Power BI: Executive Dashboard](#4-power-bi-executive-dashboard)
- [Repository Structure](#repository-structure)
- [Data Model](#data-model)
- [Key Business Findings](#key-business-findings)
  - [Orders & Delivery](#orders--delivery)
  - [Hub Operations](#hub-operations)
  - [Drivers & Workforce](#drivers--workforce)
  - [Vehicles & Fleet](#vehicles--fleet)
- [Statistical & Predictive Analysis](#statistical--predictive-analysis)
- [Data Quality Findings](#data-quality-findings)
- [Business Recommendations](#business-recommendations)
- [Dashboard](#dashboard)
- [Business Report](#business-report)
- [What This Project Demonstrates](#what-this-project-demonstrates)
- [Limitations & Next Steps](#limitations--next-steps)
- [How to Explore the Repository](#how-to-explore-the-repository)
- [Tools & Technologies](#tools--technologies)
- [Project Scope](#project-scope)

---

## Project Overview

SwiftRoute Logistics needed both:

1. A **day-to-day operational dashboard** for monitoring logistics performance.
2. A deeper analytical layer to understand **why performance changes, which patterns are meaningful, and where management should focus resources**.

This project builds both.

The analysis covers:

- **27,979 orders**
- **55 drivers**
- **45 vehicles**
- **6 operational hubs**
- **January 2023 – December 2024**
- **44 business questions**
- Four major operational areas:
  - Orders & Delivery Performance
  - Hub Operations
  - Drivers & Workforce
  - Vehicles & Fleet

The project intentionally includes both positive and negative findings. Where a relationship was not statistically significant, it was reported as such rather than forcing a business conclusion from a weak pattern.

---

## Business Problem

The project is designed around questions that a logistics business could realistically ask:

- Are deliveries improving or deteriorating?
- What is driving delivery delays?
- Which hubs are operating above capacity?
- Is customer satisfaction affected by delivery delays?
- Are some drivers carrying disproportionately high workloads?
- Does driver experience translate into better performance?
- Are older vehicles more likely to break down?
- Does vehicle type affect delivery speed?
- Can delayed orders be predicted in advance?
- Which operational problems actually justify management intervention?

The goal is therefore not just to produce charts.

The goal is to move from:

**What happened?**

→ **Why did it happen?**

→ **Is the pattern statistically real?**

→ **What should the business do about it?**

---

## Project Objectives

### 1. Establish data reliability

Validate the four source datasets for:

- Missing values
- Duplicate keys
- Invalid values
- Referential integrity
- Unexpected relationships between tables

### 2. Understand operational performance

Analyze:

- Order volume
- On-time delivery
- Delivery time
- Customer satisfaction
- Delays
- Cancellations
- Hub capacity
- Driver workload
- Vehicle utilization

### 3. Test business hypotheses

Use statistical methods to determine whether observed relationships are meaningful rather than simply visual patterns.

### 4. Build an executive dashboard

Create a Power BI dashboard that allows stakeholders to monitor operational KPIs and drill into hubs, drivers, and vehicles.

### 5. Translate analysis into decisions

Turn analytical findings into practical recommendations around:

- Capacity planning
- Delay reduction
- Workforce allocation
- Fleet maintenance
- Data collection
- Future predictive analytics

---

# Dataset

The project uses four related datasets.

| Dataset | Records | Purpose |
|---|---:|---|
| `Orders.csv` | 27,979 | Transaction-level delivery and order performance |
| `Drivers.csv` | 55 | Driver master and performance attributes |
| `Hubs.csv` | 6 | Hub master and operational attributes |
| `Vehicles.csv` | 45 | Vehicle master and fleet attributes |

The datasets are connected through driver, vehicle, and hub identifiers. The Orders table contains transaction-level activity while the other three tables provide master/reference information.

### Orders

Key fields include:

- Order ID
- Order Date
- Actual Delivery Date
- Driver ID
- Driver Name
- Hub Name
- Vehicle Name
- Vehicle Type
- Order Status
- Is Delayed
- Is On Time
- Delay Reason
- Customer Satisfaction Score
- Delivery Time Hours
- Hub Processing Time Hours

### Drivers

Key fields include:

- Driver ID
- Driver Name
- Employment Type
- Hire Date
- Experience Years
- Performance Rating

### Hubs

Key fields include:

- Hub ID
- Hub Name
- Hub Capacity

### Vehicles

Key fields include:

- Vehicle ID
- Vehicle Code
- Vehicle Model
- Vehicle Status
- Purchase Date
- Breakdown Count
- Maintenance Alert

The detailed field definitions and observed values are documented in the project's [Data Dictionary](0.%20Datasets%20%26%20Dictionary/SwiftRoute_Data_Dictionary.pdf).

---

# Analytics Workflow

The project follows a four-stage analytical pipeline.

```text
                    RAW DATA
                        │
                        ▼
            ┌──────────────────────┐
            │ Excel                 │
            │ Data Preparation & QA │
            └──────────┬───────────┘
                       │
                       ▼
            ┌──────────────────────┐
            │ SQL / SQLite          │
            │ Operational Analysis  │
            └──────────┬───────────┘
                       │
                       ▼
            ┌──────────────────────┐
            │ Python                │
            │ Statistics & Modeling │
            └──────────┬───────────┘
                       │
                       ▼
            ┌──────────────────────┐
            │ Power BI              │
            │ Executive Dashboard   │
            └──────────┬───────────┘
                       │
                       ▼
              BUSINESS DECISIONS
```

---

## 1. Excel: Data Preparation & Quality Assurance

**Location:** [`2. Excel WorkBook`](2.%20Excel%20WorkBook)

Excel is the first analytical layer.

The workbook validates the raw extracts before they are used for downstream analysis and also provides quick first-pass answers using formulas such as:

- `SUMIFS`
- `COUNTIFS`
- `AVERAGEIFS`

The QA process checks:

- Row counts
- Duplicate identifiers
- Missing values
- Valid ranges
- Referential integrity
- Cancelled-order handling
- Driver relationships
- Vehicle relationships
- Hub relationships

The workbook also contains first-pass analyses for:

- Delay reasons
- Cancellations
- Hub capacity utilization
- Driver workload
- Vehicle utilization

[View the QA documentation](2.%20Excel%20WorkBook/SwiftRoute_Data_Preparation_Quality%20Assurance.pdf)

[Open the Excel workbook](2.%20Excel%20WorkBook/SwiftRoute_Data_Preparation_Quality%20Assurance.xlsx)

---

## 2. SQL: Operational Analysis

**Location:** [`3. SQL & Python Notebooks/A. SwiftRoute_SQL_Analysis.ipynb`](3.%20SQL%20%26%20Python%20Notebooks/A.%20SwiftRoute_SQL_Analysis.ipynb)

The SQL layer moves beyond simple aggregation and focuses on operational questions requiring:

- Table joins
- Common Table Expressions (CTEs)
- Aggregations
- Ranking
- Window functions
- `RANK`
- `LAG`
- Trend analysis
- Cross-table comparisons

Examples of questions answered include:

- Which drivers handle the most deliveries?
- Which hubs are consistently above capacity?
- How does workload vary across drivers?
- Does employment type relate to service performance?
- How does cancellation rate vary?
- What is the leading delay reason by hub?
- Does the Power BI month-over-month calculation reproduce independently in SQL?

SQL was also used as an independent validation layer for dashboard metrics.

---

## 3. Python: Statistical & Predictive Analysis

**Location:** [`3. SQL & Python Notebooks/B. SwiftRoute_Python_Analysis.ipynb`](3.%20SQL%20%26%20Python%20Notebooks/B.%20SwiftRoute_Python_Analysis.ipynb)

Python is used to answer a different question:

> Are the patterns observed in the operational data actually statistically meaningful?

The analysis uses:

- `pandas`
- `scipy.stats`
- `scikit-learn`

Statistical techniques include:

- Chi-square tests
- ANOVA
- Independent-samples t-tests
- Pearson correlation
- Multiple linear regression
- Random Forest classification
- Quantile-based risk segmentation
- Trend analysis

This stage is particularly important because visual differences are not automatically meaningful.

The project therefore reports both:

**Significant findings**

and

**Negative findings where the data does not support a meaningful relationship.**

---

## 4. Power BI: Executive Dashboard

**Location:** [`4. Power BI Dashboard`](4.%20Power%20BI%20Dashboard)

The Power BI layer converts the validated analysis into an interactive operational dashboard.

The dashboard contains four major views:

### Overview

Provides high-level KPIs including:

- Total Orders
- On-Time Delivery Rate
- Customer Satisfaction
- Average Delivery Time
- Hub performance
- Driver performance
- Vehicle utilization

### Hubs Overview

Provides:

- Hub capacity vs. orders
- Hub performance ranking
- Processing time by weekday
- Hub-level operational comparison

### Drivers Overview

Provides:

- Driver experience vs. performance rating
- Drivers with the highest delay share
- Individual driver profiles
- Monthly delivery trends

### Vehicles Overview

Provides:

- Active vs. maintenance status
- Orders by vehicle model
- Vehicle age vs. breakdowns
- Breakdown counts by model
- Breakdown counts by vehicle code
- Orders by vehicle type

The dashboard uses DAX measures and interactive filters to allow stakeholders to move from high-level KPIs into operational detail.

[View the dashboard PDF](4.%20Power%20BI%20Dashboard/SwiftRoute%20Logistics%20Dashboard.pdf)

[Open the Power BI file](4.%20Power%20BI%20Dashboard/SwiftRoute%20Logistics%20Dashboard.pbix)

---

# Repository Structure

```text
SwiftRoute-Logistics-Analytics/
│
├── 0. Datasets & Dictionary/
│   ├── Drivers.csv
│   ├── Hubs.csv
│   ├── Orders.csv
│   ├── SwiftRoute_Data_Dictionary.pdf
│   └── Vehicles.csv
│
├── 1. Business Questions/
│   └── SwiftRoute_Business_Questions_Catalog.pdf
│
├── 2. Excel WorkBook/
│   ├── SwiftRoute_Data_Preparation_Quality Assurance.pdf
│   └── SwiftRoute_Data_Preparation_Quality Assurance.xlsx
│
├── 3. SQL & Python Notebooks/
│   ├── A. SwiftRoute_SQL_Analysis.ipynb
│   └── B. SwiftRoute_Python_Analysis.ipynb
│
├── 4. Power BI Dashboard/
│   ├── SwiftRoute Logistics Dashboard.pbix
│   └── SwiftRoute Logistics Dashboard.pdf
│
├── 5. Business Report/
│   └── SwiftRoute_Business_Report.pdf
│
├── LICENSE
└── README.md
```

---

# Data Model

The project uses `Orders` as the central transactional table.

```text
                         ┌───────────────┐
                         │    Drivers    │
                         │               │
                         │   Driver ID   │
                         └───────┬───────┘
                                 │
                                 │ Driver ID
                                 │
┌───────────────┐         ┌──────▼───────┐         ┌───────────────┐
│     Hubs      │         │    Orders    │         │   Vehicles    │
│               │◄────────│              │────────►│               │
│    Hub Name   │         │  Order ID    │         │ Vehicle Code  │
│   Capacity    │         │  Driver ID   │         │ Vehicle Model │
└───────────────┘         │  Hub Name    │         │   Status      │
                          │ Vehicle Name  │         └───────────────┘
                          │ Delay / CSAT  │
                          │ Delivery Time │
                          └──────────────┘
```

A key data-quality decision was made here:

> **Driver-level analysis uses Driver ID rather than Driver Name.**

The order-level Driver Name field did not reliably match the driver master table. This was identified during QA, and all downstream driver joins therefore use the ID field.

---

# Key Business Findings

## Orders & Delivery

### Delivery delays are the strongest customer-experience lever

Delayed orders had an average CSAT of:

- **2.94** for delayed orders
- **4.50** for on-time orders

That represents a **1.56-point CSAT gap** on a 5-point scale, with **p < 0.0001**.

A multiple regression model also found that the delay indicator dominated the other tested factors.

**Business implication:** Reducing delays should be treated as the highest-priority lever for improving customer satisfaction.

---

### Order volume is essentially flat

Across the two-year period, monthly order volume remains around approximately **1,150 orders per month**.

The trend analysis produced:

**R² = 0.008**

indicating virtually no underlying growth trend.

**Business implication:** Capacity planning should focus on the existing demand baseline rather than assuming rapid volume growth.

---

### Cancellations are not currently a major issue

The overall cancellation rate is:

**0.9%**

or:

**252 out of 27,979 orders**

The rate varies only narrowly across hubs and vehicle types.

**Business implication:** Cancellation reduction does not currently appear to be a priority intervention area.

---

### No single delay reason dominates company-wide

All ten recorded delay reasons occur at roughly **9–11%** of delays.

Examples include:

- Road Construction
- Vehicle Breakdown
- Package Sorting Error
- Driver Unavailable
- Hub Processing Delay
- Traffic Congestion
- Severe Weather
- Customer Not Home
- Incorrect Address
- Multiple Delivery Stops

The leading reason differs by hub.

**Business implication:** A single company-wide delay intervention is unlikely to solve the problem. Delay reduction needs to be approached at the hub/root-cause level.

---

# Hub Operations

## Dallas Main Hub is structurally over capacity

Dallas Main Hub operates at approximately:

**121% of rated monthly capacity**

while the other hubs operate around **75–77%**.

More importantly, SQL trend analysis shows that Dallas exceeds 100% capacity in nearly every month across the two-year period.

This is therefore not simply a temporary spike.

**Business implication:** SwiftRoute should consider:

- Increasing Dallas capacity
- Rebalancing orders
- Shifting some volume toward hubs with available capacity

---

## Hub performance differences are not statistically significant

The dashboard shows hub on-time performance ranging from approximately:

**83.3% → 76.7%**

However, a chi-square test produced:

**p = 0.174**

Therefore, the observed difference does not reach the 5% significance threshold.

**Business implication:** The dashboard ranking is useful for monitoring, but leadership should not automatically classify the lowest-ranked hub as operationally underperforming.

---

## Hub processing time does not explain delivery time

The correlation between hub processing time and total delivery time was approximately:

**r ≈ 0.00**

with:

**p = 0.68**

**Business implication:** Faster hub processing alone is unlikely to materially improve end-to-end delivery time based on the variables available in this dataset.

---

# Drivers & Workforce

## Driver workload is highly uneven

Across two years:

- Busiest driver: **1,281 deliveries**
- Least-used driver: **345 deliveries**
- Approximate spread: **3.7×**

This suggests a meaningful workload imbalance.

**Business implication:** Dispatch allocation and roster distribution should be reviewed for fairness and operational efficiency.

---

## Experience is associated with higher performance ratings

Driver experience and performance rating show:

**r = 0.45**

with:

**p < 0.001**

This is a moderate, statistically significant positive relationship.

**Business implication:** Driver retention and experience may have operational value beyond simply reducing hiring and training costs.

---

## Driver tenure does not meaningfully predict delays

Delay rates remain relatively flat across tenure groups:

- 0–2 years: **20.7%**
- 2–4 years: **21.0%**
- 4+ years: **21.5%**

**Business implication:** Tenure by itself should not be treated as a primary lever for reducing delivery delays.

---

# Vehicles & Fleet

## Vehicle age is strongly associated with breakdowns

Vehicle age has a statistically significant positive correlation with breakdown frequency:

**r = 0.65**

with:

**p < 0.001**

Vehicle age alone explains approximately **42% of breakdown variance**.

**Business implication:** Vehicle age provides a strong basis for prioritizing inspections and replacement decisions.

---

## Maintenance alerts add very little beyond vehicle age

A regression model comparing:

**Age only → R² = 0.42**

versus:

**Age + Maintenance Alerts → R² = 0.43**

shows that the existing maintenance-alert field adds very little predictive information beyond vehicle age.

**Business implication:** An age-based maintenance strategy may be more useful than relying heavily on the current alert field.

---

## The oldest third of the fleet is the priority group

Quantile-based segmentation identifies the oldest third of vehicles as carrying a disproportionate share of breakdowns.

**Business implication:** Fleet management can use vehicle age as a practical first-pass maintenance prioritization rule.

---

## Four vehicles have zero linked deliveries

Four of the 45 vehicles have no linked deliveries over the two-year period.

The QA process confirmed that this is a genuine utilization finding rather than a vehicle-code join problem.

**Business implication:** These vehicles should be reviewed for:

- Idle capacity
- Maintenance status
- Deployment decisions
- Fleet right-sizing

---

# Statistical & Predictive Analysis

One of the strongest parts of the project is the distinction between **observed patterns** and **statistically supported relationships**.

| Question | Method | Result |
|---|---|---|
| Do hub on-time rates differ significantly? | Chi-square | **p = 0.174, not significant** |
| Does vehicle type affect delivery time? | ANOVA | **p = 0.58, not significant** |
| Does delay affect CSAT? | T-test | **p < 0.0001, significant** |
| Does driver experience relate to rating? | Pearson correlation | **r = 0.45, p < 0.001** |
| Does vehicle age relate to breakdowns? | Pearson correlation | **r = 0.65, p < 0.001** |
| Does hub processing time relate to delivery time? | Pearson correlation | **r ≈ 0.00, p = 0.68** |
| Can CSAT be explained by operational factors? | Multiple regression | **R² = 0.54** |
| Can order-level delay be predicted? | Random Forest | **AUC = 0.50** |

### An important negative finding

The Random Forest model produced:

**AUC = 0.50**

which is effectively chance-level prediction.

Instead of treating this as a failed project, the result identifies a data limitation.

The current dataset does not contain enough real-time information to predict individual delivery delays reliably.

Potentially useful future variables include:

- Live traffic conditions
- Weather severity
- Real-time hub queue length
- Dispatch-time operational conditions

This is an important distinction:

> A good analytics project should identify where prediction works, but also where the available data is not sufficient to support prediction.

---

# Data Quality Findings

The QA stage identified several important characteristics of the dataset.

### Core integrity checks passed

The datasets contained:

- No duplicate Order IDs
- No duplicate Driver IDs
- No duplicate Vehicle Codes
- No orphaned Driver IDs
- No unmatched Hub Names
- No unmatched Vehicle Codes
- Valid CSAT scores within the expected 1–5 range

### Cancelled orders

There are **252 cancelled orders**.

Their missing delivery time and delivery date are expected because cancelled orders were never completed.

### Driver name mismatch

The most important data-quality issue was the mismatch between the `Driver Name` stored in Orders and the corresponding driver in the Drivers master table.

The QA check found this mismatch across the driver records.

As a result:

```text
DO NOT JOIN:
Orders.Driver Name → Drivers.Driver Name

USE:
Orders.Driver ID → Drivers.DriverID
```

This decision was applied throughout the SQL and Python analysis.

---

# Business Recommendations

Based on the combined Excel, SQL, Python, and Power BI analysis, the project points toward five main priorities.

## 1. Prioritize delay reduction

Delay has the strongest demonstrated relationship with customer satisfaction.

Focus improvement efforts on reducing the overall delay rate rather than optimizing variables that have not demonstrated meaningful effects.

---

## 2. Rebalance Dallas Main Hub

Dallas Main Hub is consistently operating above rated capacity.

Management should evaluate:

- Additional capacity
- Staffing/dock expansion
- Volume redistribution
- Greater use of available capacity at other hubs

---

## 3. Review driver workload allocation

The 3.7× workload gap between the busiest and least-used drivers suggests that dispatch allocation deserves investigation.

The next step should be to determine whether the imbalance is caused by:

- Route geography
- Hub assignment
- Driver availability
- Dispatch logic
- Vehicle assignment

---

## 4. Introduce age-based fleet maintenance prioritization

Vehicle age is the strongest available predictor of breakdown frequency.

The oldest third of the fleet should be prioritized for:

- Preventive inspection
- Maintenance review
- Replacement evaluation

---

## 5. Improve operational data collection

The current data supports descriptive and statistical analysis well, but it is insufficient for reliable order-level delay prediction.

Future data collection should consider:

- Live traffic
- Weather conditions
- Hub queue length
- Dispatch-time conditions
- More granular delay duration
- Detailed cancellation reasons

Better operational data would create a stronger foundation for predictive analytics.

---

# Dashboard

The Power BI dashboard provides four linked views:

### Overview

High-level operational KPIs and comparisons.

### Hubs

Capacity, order volume, performance ranking, and processing time.

### Drivers

Experience, performance, delay share, driver profiles, and delivery trends.

### Vehicles

Fleet status, vehicle utilization, breakdowns, vehicle age, and model-level performance.

The dashboard is designed for stakeholders who need answers without opening the underlying SQL or Python notebooks.

[Open the dashboard PDF](4.%20Power%20BI%20Dashboard/SwiftRoute%20Logistics%20Dashboard.pdf)

---

# Business Report

The final business report consolidates the project's findings into a structured narrative.

It provides:

- Business question
- Analytical method
- Finding
- Business implication

The report covers the four operational categories:

1. Orders & Delivery Performance
2. Hub Operations
3. Drivers & Workforce
4. Vehicles & Fleet

[Read the Business Report](5.%20Business%20Report/SwiftRoute_Business_Report.pdf)

---

# What This Project Demonstrates

This project demonstrates the ability to work across the complete analytics lifecycle rather than treating analysis as a single-tool exercise.

### Data Preparation

- Raw data validation
- Data quality checks
- Referential integrity
- Key validation
- Data issue identification

### SQL

- Relational thinking
- Joins
- CTEs
- Aggregations
- Window functions
- Ranking
- Time-series analysis

### Python

- Data manipulation with pandas
- Statistical hypothesis testing
- Correlation analysis
- Regression
- Classification
- Model evaluation
- Segmentation

### Power BI

- Data modeling
- DAX measures
- KPI design
- Interactive filtering
- Drill-through analysis
- Executive dashboard design

### Business Analysis

Most importantly, the project focuses on the connection between:

**Data → Evidence → Insight → Business Decision**

It also demonstrates the ability to communicate negative findings honestly instead of assuming that every observed difference represents a real business effect.

---

# Limitations & Next Steps

This project uses a synthetic portfolio dataset.

The analysis is therefore intended to demonstrate an end-to-end analytics workflow rather than represent SwiftRoute's actual operational data.

Several limitations also emerge from the available fields:

- No live traffic data
- No detailed weather severity data
- No real-time hub queue information
- Limited information about individual delay incidents
- Limited cancellation-reason information
- No explicit route-level characteristics

### Potential next phase

A future version of the project could incorporate:

```text
Live Traffic
     +
Weather
     +
Route Distance
     +
Hub Queue Length
     +
Driver Availability
     +
Vehicle Availability
     ↓
Real-Time Delay Risk Model
```

This would provide the additional signal required to move from retrospective analysis toward operational prediction.

---

# How to Explore the Repository

If you are reviewing this project for the first time, the recommended path is:

### Step 1 — Understand the data

Start with:

[`SwiftRoute_Data_Dictionary.pdf`](0.%20Datasets%20%26%20Dictionary/SwiftRoute_Data_Dictionary.pdf)

This explains the four datasets and their fields.

### Step 2 — Review the QA process

Open:

[`SwiftRoute_Data_Preparation_Quality%20Assurance.pdf`](2.%20Excel%20WorkBook/SwiftRoute_Data_Preparation_Quality%20Assurance.pdf)

This shows how the raw data was validated and documents the driver-ID issue.

### Step 3 — Explore the SQL analysis

Open:

[`A. SwiftRoute_SQL_Analysis.ipynb`](3.%20SQL%20%26%20Python%20Notebooks/A.%20SwiftRoute_SQL_Analysis.ipynb)

This contains the operational queries and deeper relational analysis.

### Step 4 — Explore the Python analysis

Open:

[`B. SwiftRoute_Python_Analysis.ipynb`](3.%20SQL%20%26%20Python%20Notebooks/B.%20SwiftRoute_Python_Analysis.ipynb)

This contains the statistical testing and predictive/segmentation work.

### Step 5 — View the dashboard

Open:

[`SwiftRoute Logistics Dashboard.pdf`](4.%20Power%20BI%20Dashboard/SwiftRoute%20Logistics%20Dashboard.pdf)

This provides a quick visual overview of the final stakeholder-facing output.

### Step 6 — Read the final business report

Finish with:

[`SwiftRoute_Business_Report.pdf`](5.%20Business%20Report/SwiftRoute_Business_Report.pdf)

This brings the entire analysis together into business findings and implications.

---

# Tools & Technologies

### Data & Spreadsheet Analysis

- Microsoft Excel
- Excel formulas
- Pivot-style analysis
- Data Quality Assurance

### Database & Querying

- SQL
- SQLite
- SQL joins
- CTEs
- Window functions
- `RANK`
- `LAG`

### Statistical & Predictive Analysis

- Python
- pandas
- SciPy
- scikit-learn
- Statistical hypothesis testing
- Regression
- Random Forest
- Quantile segmentation

### Business Intelligence

- Microsoft Power BI
- DAX
- Interactive dashboards
- KPI reporting
- Data visualization

### Documentation & Reporting

- Jupyter Notebooks
- PDF reporting
- Business-question documentation

---

# Project Scope

| Category | Coverage |
|---|---|
| Orders | 27,979 |
| Drivers | 55 |
| Vehicles | 45 |
| Hubs | 6 |
| Time Period | Jan 2023 – Dec 2024 |
| Business Questions | 44 |
| Analytical Stages | 4 |
| Dashboard Views | 4 |

---

## Final Takeaway

The central lesson from the SwiftRoute analysis is that **not every operational difference deserves an intervention**.

The strongest evidence points toward three priorities:

1. **Reduce delivery delays because they have the clearest relationship with customer satisfaction.**
2. **Address the structural capacity imbalance at Dallas Main Hub.**
3. **Use vehicle age to prioritize fleet maintenance and replacement decisions.**

At the same time, the analysis shows where the current data does **not** support strong conclusions, such as vehicle-type effects on delivery time, hub-level on-time differences, driver tenure as a delay lever, and individual-order delay prediction.

That distinction between **what the data shows, what the data proves, and what the data cannot yet answer** is at the core of this project.

---

**Author:** Simran Kirtania — MSc Economics, aspiring Analyst
📧 simrankirtania02@gmail.com · 🔗 [LinkedIn](https://www.linkedin.com/in/simrankirtania/) · 🌐 [Portfolio](https://simrankirtaniaportfolio.netlify.app)
