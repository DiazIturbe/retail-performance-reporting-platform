# Store Performance Reporting Platform

### Executive Retail Analytics & KPI Reporting

Transforming raw Tableau exports into executive-ready retail performance reports.

**Python • Streamlit • Tableau • Pandas • HTML Reporting**

---

## Executive Summary

The **Store Performance Reporting Platform** automates the creation of weekly retail performance reports by transforming multiple Tableau exports into a single executive-ready dashboard.

Designed for retail operations, the platform streamlines data validation, KPI calculations, report generation, and visualization, allowing managers to focus on decision-making rather than manual reporting.

---

## The Challenge

Weekly retail reporting often requires manually downloading, validating, and consolidating multiple Tableau exports before producing management reports.

This repetitive process is time-consuming, prone to errors, and makes it difficult to maintain a consistent reporting standard across reporting periods.

---

## The Solution

The Store Performance Reporting Platform automates the reporting workflow by:

- Validating uploaded Tableau exports
- Consolidating multiple datasets into a unified data model
- Calculating operational KPIs
- Generating executive dashboards
- Producing interactive HTML reports
- Exporting formatted Excel reports

The result is a standardized reporting process that produces consistent, executive-ready outputs in minutes.

---

# Dashboard Overview

The executive dashboard provides a high-level view of store performance, highlighting the most important KPIs, budget performance, sales contribution, and operational metrics in a single page.

![Executive Dashboard](assets/readme/executive_dashboard.png)

---

# Executive Summary

Automatically generated executive summaries highlight the week's most important insights, allowing managers to understand store performance at a glance.

![Executive Summary](assets/readme/executive_summary.png)

---

# Department Analysis

Detailed department and subdepartment analysis provides visibility into sales performance, budget attainment, product contribution, and operational KPIs.

![Department Analysis](assets/readme/department_analysis.png)

---

# Top Products & Brand Performance

The platform automatically identifies the top-performing products for every department and subdepartment, supporting merchandising and inventory decisions.

![Product Performance](assets/readme/brands_products.png)

---

## Workflow

```text
          Tableau Exports
                  │
                  ▼
        Upload & Validation
                  │
                  ▼
        Data Processing Engine
                  │
                  ▼
       KPI Calculations & Metrics
                  │
      ┌───────────┴───────────┐
      ▼                       ▼
 HTML Executive Report   Excel Report
                  │
                  ▼
      Management Decision Support
```

---

## Key Capabilities

- Automated Tableau export validation
- Guided upload workflow
- Executive KPI dashboard
- Department-level performance analysis
- Budget gap analysis
- Sales mix visualization
- Brand performance summaries
- Interactive HTML reporting
- Excel report generation

---

## Technology Stack

### Programming

- Python

### Frameworks

- Streamlit

### Data Processing

- Pandas
- NumPy

### Visualization

- Matplotlib

### Excel Reporting

- OpenPyXL
- XlsxWriter

---

## Repository Structure

```text
📁 app.py
Main Streamlit application

📁 upload_classifier.py
Upload validation

📁 vm_report_data_model_v2.py
Data model construction

📁 report_calculations.py
KPI calculations

📁 chart_builder.py
Chart generation

📁 html_report_builder.py
HTML report generation
```

---

## Roadmap

- PDF report enhancements
- Historical trend analysis
- District performance comparisons
- Interactive filtering
- Performance optimization
- Configurable KPI templates
- PDF report enhancements
- Historical trend analysis
- District performance comparisons
- Interactive filtering
- Performance optimization
- Configurable KPI templates

---

## Author

**Diego Díaz Iturbe**

Data Analytics • Automation • Cloud • GIS
