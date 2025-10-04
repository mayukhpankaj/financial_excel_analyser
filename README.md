# Excel AI Financial Analyzer

An intelligent tool for extracting and analyzing financial data from Excel files using Streamlit and Python.

<img width="3840" height="1535" alt="image" src="https://github.com/user-attachments/assets/4b325250-9120-4aeb-a0ab-8b285d911806" />


## Features


- 🚀 Automatic extraction of key financial metrics from Excel files
- 📊 Intelligent pattern matching to find financial data in various formats
- 💹 Automatic calculation of financial ratios and metrics
- 📈 Interactive visualizations of financial data
- 🎨 Clean, modern user interface

## Installation

1. Clone this repository
2. Install the required dependencies:

```bash
pip install -r requirements.txt
```

## Usage

1. Run the Streamlit app:

```bash
streamlit run app.py
```

2. Open your browser and navigate to the URL shown in the terminal (usually http://localhost:8501)
3. Upload an Excel file containing financial data
4. View the extracted metrics and visualizations

## Supported Financial Metrics

The tool can automatically detect and extract:

- Revenue/Sales
- Net Income
- EBITDA
- Total Assets
- Total Liabilities
- Equity
- Debt
- Tax Rates

And calculates:
- Debt-to-Equity Ratio
- Net Profit Margin
- Return on Equity (ROE)

## How It Works

The application uses pattern matching to identify financial terms in the Excel file, regardless of their exact position. It then extracts the corresponding values and performs financial calculations to provide meaningful insights.

## Requirements

- Python 3.8+
- See `requirements.txt` for Python package dependencies

## License

MIT
