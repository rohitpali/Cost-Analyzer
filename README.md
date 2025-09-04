# Courier Cost Analyzer

A Python-based data pipeline that reconciles courier invoices against company order records to detect overcharges and undercharges.

## Features

• Automated reconciliation of courier bills with internal records
• Weight slab calculation and charge validation
• Summary table for correct, overcharged, and undercharged orders
• Excel report generation with detailed calculations
• Business insights and visualizations (bar, pie, histogram)

## Architecture

The system integrates multiple Excel files:

1. **Company Order Report** (SKU, quantity)
2. **SKU Master** (product weights)
3. **Pincode Zones** (delivery mapping)
4. **Courier Invoice** (billed charges)
5. **Courier Rates** (zone-wise slab charges)

The pipeline merges these datasets, applies weight slab rounding logic, computes expected charges, compares them with billed charges, and generates reports + visual insights.

## Setup

1. Clone repository
2. Place all Excel input files in the 'Data/' folder
3. Install dependencies: `pip install -r requirements.txt`
4. Run analysis.py in VS Code or terminal
5. Outputs:
   • Reconciliation_Result.xlsx
   • summary_bar.png
   • summary_pie.png
   • diff_hist.png

## Usage

Run: `python analysis.py`

This generates a detailed reconciliation report and visualizations.

## Input Files Structure

Ensure the following Excel files are placed in the `Data/` folder:

- `Company X - Order Report.xlsx` - Contains order details with SKU and quantities
- `Company X - SKU Master.xlsx` - Product weight information
- `Company X - Pincode Zones.xlsx` - Delivery zone mappings
- `Courier Company - Invoice.xlsx` - Courier billing details
- `Courier Company - Rates.xlsx` - Zone-wise rate slabs

## Output Files

### Excel Report
- **Reconciliation_Result.xlsx**
  - Summary sheet with order counts and amounts by category
  - Calculations sheet with detailed reconciliation data

### Visualizations
- **summary_bar.png** - Bar chart showing order counts by category
- **summary_pie.png** - Pie chart showing amount distribution
- **diff_hist.png** - Histogram of charge differences

## Dependencies

```
pandas
numpy
matplotlib
xlsxwriter
```

## Business Insights

The tool provides automated analysis to:
- Identify correctly charged orders
- Detect overcharges for potential recovery
- Track undercharges for consistency monitoring
- Generate actionable business recommendations

## License

This project is for internal business use.