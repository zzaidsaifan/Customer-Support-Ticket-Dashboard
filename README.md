# Customer Support SLA Dashboard

## Overview

The Customer Support SLA Dashboard is an Excel-based Business Intelligence project designed to monitor customer support operations, evaluate Service Level Agreement (SLA) compliance, and measure support team performance. The project leverages Excel formulas, Pivot Tables, Pivot Charts, Dashboard Design, and VBA automation to transform raw ticket data into actionable business insights.

This dashboard helps organizations track ticket resolution efficiency, identify SLA breaches, analyze support trends, and improve customer satisfaction through data-driven decision-making.

---

## Business Problem

Customer support teams often struggle to monitor service performance across large volumes of support tickets. Without proper tracking, SLA violations can negatively impact customer satisfaction and operational efficiency.

This project addresses these challenges by providing:

- SLA Compliance Monitoring
- Ticket Resolution Analysis
- Support Agent Performance Evaluation
- Customer Satisfaction Tracking
- Automated Reporting and Dashboard Refresh

---

## Project Objectives

- Measure SLA compliance across ticket priorities.
- Track ticket resolution performance.
- Identify categories with the highest SLA breach rates.
- Evaluate support agent productivity.
- Analyze customer satisfaction trends.
- Create an interactive dashboard for management reporting.
- Automate dashboard updates using VBA.

---

## Dataset Information

The dataset contains customer support ticket records with the following attributes:

| Column Name | Description |
|------------|-------------|
| Ticket ID | Unique identifier for each support ticket |
| Customer ID | Customer reference number |
| Priority | Ticket priority level (Critical, High, Medium, Low) |
| Issue Type | Category of customer issue |
| Created Date | Ticket creation date |
| Resolved Date | Ticket closure date |
| Resolution Time (Hours) | Total time taken to resolve the ticket |
| Satisfaction Score | Customer feedback score |
| Resolution Category | Open or Closed status |
| Agent Name | Assigned support representative |

---

## SLA Framework

A separate SLA Master Table was created to define expected resolution times for each ticket priority.

| Priority | SLA Resolution Time (Hours) |
|-----------|----------------------------|
| Critical | 4 |
| High | 8 |
| Medium | 24 |
| Low | 72 |

### SLA Calculation Process

1. Assign SLA hours based on ticket priority.
2. Compare actual resolution time against SLA hours.
3. Classify tickets as:
   - Met
   - Breached

### Formula Used

```excel
=XLOOKUP([@Priority],SLA_Master!A:A,SLA_Master!B:B)
```

```excel
=IF([@[Resolution Time (Hours)]]<=[@[SLA Hours]],"Met","Breached")
```

---

## Key Performance Indicators (KPIs)

The dashboard tracks the following KPIs:

- Total Tickets
- Open Tickets
- Closed Tickets
- Average Resolution Time
- SLA Compliance Rate (%)
- SLA Breach Rate (%)
- Average Customer Satisfaction Score

### SLA Compliance Formula

```excel
=COUNTIF(SLA_Status,"Met")/COUNTA(SLA_Status)
```

### SLA Breach Formula

```excel
=COUNTIF(SLA_Status,"Breached")/COUNTA(SLA_Status)
```

---

## Analysis Performed

### 1. SLA Compliance by Priority
Analyzes how effectively support teams meet SLA targets across different priority levels.

### 2. SLA Compliance by Issue Type
Identifies issue categories that contribute most to SLA violations.

### 3. Resolution Time Analysis
Measures average ticket resolution time across priorities and issue types.

### 4. Agent Performance Analysis
Evaluates:
- Tickets handled
- Average resolution time
- SLA compliance rate
- Customer satisfaction score

### 5. Monthly Ticket Trend Analysis
Tracks ticket volumes and operational workload over time.

---

## Dashboard Features

### KPI Summary Section
- Total Tickets
- Open Tickets
- Closed Tickets
- SLA Compliance %
- Average Resolution Time
- Customer Satisfaction Score

### Interactive Visualizations
- Monthly Ticket Trend
- SLA Compliance by Priority
- SLA Breaches by Issue Type
- Agent Performance Comparison
- SLA Status Distribution

### User Interaction
- Slicers
- Dynamic Filters
- Interactive Pivot Charts
- Conditional Formatting

---

## VBA Automation

The project includes VBA macros to automate dashboard maintenance and reporting.

### Automation Features

- Refresh Pivot Tables
- Refresh Dashboard Data
- Update KPI Metrics
- Generate Reports
- Highlight SLA Breaches
- Export Dashboard to PDF

### Sample VBA Macro

```vb
Sub RefreshDashboard()

Dim pt As PivotTable
Dim ws As Worksheet

For Each ws In Worksheets
    For Each pt In ws.PivotTables
        pt.RefreshTable
    Next pt
Next ws

MsgBox "Dashboard Refreshed Successfully!"

End Sub
```

---

## Business Value

This dashboard provides organizations with:

- Improved SLA visibility
- Faster identification of support bottlenecks
- Better workload management
- Enhanced customer satisfaction monitoring
- Data-driven decision-making capabilities
- Improved operational efficiency

---

## Tools and Technologies

- Microsoft Excel
- Pivot Tables
- Pivot Charts
- Excel Formulas
- XLOOKUP
- COUNTIF
- SUMIFS
- Conditional Formatting
- VBA Macros

---

## Project Workflow

```text
Raw Ticket Data
       ↓
Data Cleaning
       ↓
SLA Master Table
       ↓
SLA Calculations
       ↓
Pivot Table Analysis
       ↓
Dashboard Development
       ↓
VBA Automation
       ↓
Business Insights
```

---


## Future Enhancements

- Power BI Integration
- Automated Email Reporting
- Real-Time Ticket Monitoring
- Predictive SLA Analytics
- Machine Learning-Based Ticket Classification

---

## Skills Demonstrated

- Data Analysis
- Business Intelligence
- Dashboard Development
- KPI Reporting
- SLA Monitoring
- Data Visualization
- Excel Automation
- VBA Programming

---

## Author

**Zaid Arif Saifan**

MSc Financial Engineering, HEC Montréal

**Areas of Interest:** Financial Analytics, Business Intelligence, Data Analysis, Excel Automation, Risk Analytics, and Dashboard Development.

---

### ⭐ If you found this project useful, please consider starring the repository.
