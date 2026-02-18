# Price Elasticity Modelling

This project analyzes product price elasticity using sales data, providing actionable insights for pricing strategy optimization. The workflow includes data cleaning, exploratory data analysis (EDA), elasticity estimation at SKU and category levels, visualizations, and automated report generation.

## Features
- Data cleaning and quality checks
- Exploratory data analysis (EDA) with visualizations
- SKU-level and category-level price elasticity estimation
- Actionable business recommendations
- Automated summary and Word report generation with embedded plots

## Outputs
- Excel files: SKU and category elasticity results
- PNG plots: Key visualizations for report-ready insights
- Word report: Submission-ready document with tables, plots, and recommendations

## How to Run
1. Place your data file in the project directory.
2. Run `price_elasticity_alaysis.py` in your Python environment.
3. Outputs will be saved in the `outputs/` directory.

## Requirements
- Python 3.7+
- pandas, numpy, matplotlib, seaborn, scikit-learn, python-docx, lxml

Install requirements with:
```
pip install pandas numpy matplotlib seaborn scikit-learn python-docx lxml
```

## Usage Example
```bash
python price_elasticity_alaysis.py
```

## Project Structure
```
Price_Elasticity_Modelling/
├── price_elasticity_alaysis.py
├── outputs/
│   ├── sku_level_elasticity.xlsx
│   ├── category_level_elasticity.xlsx
│   ├── category_avg_elasticity.png
│   ├── top10_sensitive_skus.png
│   ├── top10_insensitive_skus.png
│   ├── summary_report.txt
│   └── Price_Elasticity_Report.docx
└── ...
```

## License
MIT License

## Author
Your Name
