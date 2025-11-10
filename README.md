# Amazon Sales Analysis

Automated analysis of Amazon sales data with Excel dashboard generation.

## Overview

This project analyzes Amazon sales transactions and generates an interactive Excel dashboard with charts and insights. Each run creates a timestamped Excel file for version tracking.

## Prerequisites

1. Python 3.8 or higher
2. Required packages: pandas, numpy, openpyxl
3. Sales data file: `Amazon Sale Report.csv` in `assignment/` folder

## Installation

Install required packages:
```bash
pip install -r requirements.txt
```

## How to Run

```bash
python scripts/analyze_and_generate_dashboard.py
```

Output will be generated in `outputs/` folder with timestamp: `Amazon_Sales_Dashboard_YYYYMMDD_HHMMSS.xlsx`

## Project Structure

```
Amazon-Sales-Analysis/
├── assignment/              # Source data
├── scripts/                 # Analysis script
├── outputs/                 # Generated dashboards
└── requirements.txt         # Dependencies
```

## Output Dashboard Contains

1. Summary & Insights - Key findings and recommendations
2. Visual Dashboard - Metric cards with KPIs
3. Data Quality - Data cleaning report
4. Category Analysis - Product performance
5. Geography Analysis - State and city distribution
6. Order Status - Order lifecycle tracking
7. Size Analysis - Popular sizes
8. Sales Trend - Daily revenue patterns
9. Fulfillment Analysis - Delivery methods
10. B2B vs B2C - Customer segments

## Output Features

- Interactive charts editable in Excel
- Timestamped filenames for version control
- Professional formatting with color-coded metrics
- All analysis in single Excel file
- Charts can be copied to PowerPoint


Author : Madhu Kumari
Email: madhukumari09957@gmail.com