import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
from typing import Dict, Any, List, Tuple, Optional, Union
import re
from dataclasses import dataclass, field
from datetime import datetime
import json
from sklearn.preprocessing import LabelEncoder
from collections import defaultdict

# Set page config
st.set_page_config(
    page_title="Excel AI Analyzer",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

@dataclass
class ColumnAnalysis:
    """Class to hold analysis results for a single column in the dataset."""
    name: str
    dtype: str
    unique_values: int
    sample_values: list
    is_numeric: bool
    numeric_stats: Optional[Dict[str, float]] = None
    potential_metrics: List[str] = field(default_factory=list)
    null_count: int = 0
    null_percentage: float = 0.0
    is_date: bool = False
    date_format: Optional[str] = None
    is_currency: bool = False
    currency_symbol: Optional[str] = None
    is_percentage: bool = False
    is_id: bool = False
    is_categorical: bool = False
    categories: Optional[List[str]] = field(default_factory=list)
    description: str = ""
    suggestions: List[str] = field(default_factory=list)

@dataclass
class FinancialMetric:
    """Class to represent a financial metric with metadata."""
    name: str
    value: Any
    unit: str = ""
    description: str = ""
    confidence: float = 1.0
    source: str = ""
    time_period: Optional[str] = None
    trend: Optional[float] = None
    previous_value: Optional[float] = None
    is_derived: bool = False
    formula: Optional[str] = None
    category: str = "General"
    importance: str = "medium"  # low, medium, high
    visualization_type: str = "number"  # number, line, bar, pie, etc.
    
    def to_dict(self) -> Dict[str, Any]:
        """Convert the FinancialMetric to a dictionary for serialization."""
        return {
            "name": self.name,
            "value": self.value,
            "unit": self.unit,
            "description": self.description,
            "confidence": self.confidence,
            "source": self.source,
            "time_period": self.time_period,
            "trend": self.trend,
            "previous_value": self.previous_value,
            "is_derived": self.is_derived,
            "formula": self.formula,
            "category": self.category,
            "importance": self.importance,
            "visualization_type": self.visualization_type
        }

class ExcelAIAnalyzer:
    """Main class for analyzing Excel files containing financial data."""
    
    # Common financial terms and patterns for metric detection
    FINANCIAL_TERMS = {
        'revenue': ['revenue', 'sales', 'income', 'turnover'],
        'expense': ['expense', 'cost', 'expenditure'],
        'profit': ['profit', 'net income', 'net profit'],
        'asset': ['asset', 'inventory', 'property', 'equipment'],
        'liability': ['liability', 'debt', 'loan', 'payable'],
        'equity': ['equity', 'shareholder', 'stockholder'],
        'cash': ['cash', 'bank', 'balance'],
        'date': ['date', 'period', 'year', 'month', 'day'],
        'id': ['id', 'code', 'number', 'ref', 'reference'],
        'quantity': ['qty', 'quantity', 'amount', 'volume'],
        'price': ['price', 'rate', 'cost', 'value'],
        'discount': ['discount', 'rebate', 'deduction']
    }
    
    # Comprehensive financial ratios and their formulas
    FINANCIAL_RATIOS = {
        # Core Valuation Metrics
        'payback_period': {
            'name': 'Payback Period',
            'formula': 'initial_investment / annual_cash_flow',
            'description': 'Time required to recover the initial investment.',
            'category': 'Valuation',
            'importance': 'high'
        },
        'roic': {
            'name': 'Return on Invested Capital',
            'formula': 'nopat / (total_debt + total_equity)',
            'description': 'Measures how efficiently a company allocates its capital.',
            'category': 'Valuation',
            'importance': 'high'
        },
        'eva': {
            'name': 'Economic Value Added',
            'formula': 'nopat - (total_capital * wacc)',
            'description': 'Measures a company\'s financial performance based on residual wealth.',
            'category': 'Valuation',
            'importance': 'high'
        },
        'fcff': {
            'name': 'Free Cash Flow to Firm',
            'formula': 'ebit * (1 - tax_rate) + depreciation - capex - change_in_nwc',
            'description': 'Cash flow available to all investors after accounting for capital expenditures.',
            'category': 'Valuation',
            'importance': 'high'
        },
        
        # Profitability Metrics
        'gross_margin': {
            'name': 'Gross Margin',
            'formula': '(revenue - cogs) / revenue',
            'description': 'Measures how much out of every dollar of sales a company keeps as gross profit.',
            'category': 'Profitability',
            'importance': 'high'
        },
        'operating_margin': {
            'name': 'Operating Margin',
            'formula': 'operating_income / revenue',
            'description': 'Shows what percentage of revenue is left after paying for variable costs of production.',
            'category': 'Profitability',
            'importance': 'high'
        },
        'ebitda_margin': {
            'name': 'EBITDA Margin',
            'formula': 'ebitda / revenue',
            'description': 'Measures a company\'s operating profitability as a percentage of revenue.',
            'category': 'Profitability',
            'importance': 'high'
        },
        'net_profit_margin': {
            'name': 'Net Profit Margin',
            'formula': 'net_income / revenue',
            'description': 'Shows what percentage of revenue is actual profit after all expenses.',
            'category': 'Profitability',
            'importance': 'high'
        },
        'roa': {
            'name': 'Return on Assets',
            'formula': 'net_income / total_assets',
            'description': 'Indicates how efficiently a company uses its assets to generate profit.',
            'category': 'Profitability',
            'importance': 'high'
        },
        'roe': {
            'name': 'Return on Equity',
            'formula': 'net_income / shareholders_equity',
            'description': 'Measures the profitability of a business in relation to the equity.',
            'category': 'Profitability',
            'importance': 'high'
        },
        
        # Leverage & Risk Metrics
        'debt_to_equity': {
            'name': 'Debt to Equity Ratio',
            'formula': 'total_debt / shareholders_equity',
            'description': 'Indicates the relative proportion of shareholders\' equity and debt used to finance assets.',
            'category': 'Leverage',
            'importance': 'high'
        },
        'interest_coverage': {
            'name': 'Interest Coverage Ratio',
            'formula': 'ebit / interest_expense',
            'description': 'Measures a company\'s ability to meet its interest payments.',
            'category': 'Leverage',
            'importance': 'high'
        },
        'debt_to_ebitda': {
            'name': 'Debt to EBITDA',
            'formula': 'total_debt / ebitda',
            'description': 'Measures a company\'s ability to pay off its debt.',
            'category': 'Leverage',
            'importance': 'medium'
        },
        'cost_of_equity': {
            'name': 'Cost of Equity (CAPM)',
            'formula': 'risk_free_rate + (beta * market_risk_premium)',
            'description': 'The return a company requires to decide if an investment meets capital return requirements.',
            'category': 'Leverage',
            'importance': 'high'
        },
        'after_tax_cost_of_debt': {
            'name': 'After-Tax Cost of Debt',
            'formula': 'interest_rate * (1 - tax_rate)',
            'description': 'The effective rate that a company pays on its current debt after tax savings.',
            'category': 'Leverage',
            'importance': 'high'
        },
        
        # Liquidity & Efficiency Metrics
        'current_ratio': {
            'name': 'Current Ratio',
            'formula': 'current_assets / current_liabilities',
            'description': 'Measures a company\'s ability to pay short-term obligations.',
            'category': 'Liquidity',
            'importance': 'high'
        },
        'quick_ratio': {
            'name': 'Quick Ratio',
            'formula': '(cash + accounts_receivable) / current_liabilities',
            'description': 'Measures the ability to meet short-term obligations with its most liquid assets.',
            'category': 'Liquidity',
            'importance': 'high'
        },
        'cash_conversion_cycle': {
            'name': 'Cash Conversion Cycle',
            'formula': 'dso + dio - dpo',
            'description': 'Measures how long it takes for a company to convert its investments in inventory into cash.',
            'category': 'Efficiency',
            'importance': 'medium'
        },
        'asset_turnover': {
            'name': 'Asset Turnover',
            'formula': 'revenue / total_assets',
            'description': 'Measures a company\'s ability to generate sales from its assets.',
            'category': 'Efficiency',
            'importance': 'medium'
        },
        'eps': {
            'name': 'Earnings per Share',
            'formula': 'net_income / shares_outstanding',
            'description': 'Portion of a company\'s profit allocated to each outstanding share of common stock.',
            'category': 'Market',
            'importance': 'high'
        },
        
        # Market & Shareholder Value Metrics
        'pe_ratio': {
            'name': 'Price-to-Earnings Ratio',
            'formula': 'share_price / eps',
            'description': 'Measures a company\'s current share price relative to its per-share earnings.',
            'category': 'Market',
            'importance': 'high'
        },
        'ev_ebitda': {
            'name': 'Enterprise Value to EBITDA',
            'formula': 'enterprise_value / ebitda',
            'description': 'Measures the return a company makes on its assets.',
            'category': 'Market',
            'importance': 'high'
        },
        'dividend_yield': {
            'name': 'Dividend Yield',
            'formula': 'dividend_per_share / share_price',
            'description': 'Shows how much a company pays out in dividends each year relative to its stock price.',
            'category': 'Market',
            'importance': 'medium'
        },
        'tsr': {
            'name': 'Total Shareholder Return',
            'formula': '(ending_price - beginning_price + dividends) / beginning_price',
            'description': 'The total return to shareholders, including both capital gains and dividends.',
            'category': 'Market',
            'importance': 'high'
        },
        'wacc': {
            'name': 'Weighted Average Cost of Capital',
            'formula': '(e_v * re) + (d_v * rd * (1 - t))',
            'description': 'The average rate of return a company is expected to pay to all its security holders.',
            'category': 'Valuation',
            'importance': 'high'
        },
        'beta': {
            'name': 'Beta',
            'formula': 'cov(ra, rm) / var(rm)',
            'description': 'Measures the volatility of an investment compared to the market.',
            'category': 'Risk',
            'importance': 'medium'
        }
    }
    def __init__(self):
        """Initialize the Excel AI Analyzer."""
        self.data: Optional[pd.DataFrame] = None
        self.metrics: Dict[str, FinancialMetric] = {}
        self.sheet_names: List[str] = []
        self.column_analysis: List[ColumnAnalysis] = []
        self.sheet_data: Dict[str, pd.DataFrame] = {}
        self.current_sheet: str = ""
        self._initialize_metrics()
    
    def _initialize_metrics(self):
        """Initialize the metrics dictionary with common financial metrics."""
        # Add common financial metrics
        common_metrics = [
            ('revenue', 'Revenue', 'USD', 'Total revenue from sales', 'income_statement'),
            ('gross_profit', 'Gross Profit', 'USD', 'Revenue minus cost of goods sold', 'income_statement'),
            ('net_income', 'Net Income', 'USD', 'Total profit after all expenses', 'income_statement'),
            ('total_assets', 'Total Assets', 'USD', 'Sum of all assets', 'balance_sheet'),
            ('total_liabilities', 'Total Liabilities', 'USD', 'Sum of all liabilities', 'balance_sheet'),
            ('total_equity', 'Total Equity', 'USD', 'Shareholders\' equity', 'balance_sheet'),
            ('operating_cash_flow', 'Operating Cash Flow', 'USD', 'Cash generated from operations', 'cash_flow'),
            ('free_cash_flow', 'Free Cash Flow', 'USD', 'Cash available to investors', 'cash_flow'),
        ]
        
        for metric_id, name, unit, desc, category in common_metrics:
            self.metrics[metric_id] = FinancialMetric(
                name=name,
                value=None,
                unit=unit,
                description=desc,
                category=category,
                confidence=0.0
            )
    
    def load_excel(self, file_path: str, sheet_name: Optional[str] = None) -> bool:
        """
        Load and process the uploaded Excel file.
        
        Args:
            file_path: Path to the Excel file
            sheet_name: Optional name of a specific sheet to load. If None, loads the first sheet.
            
        Returns:
            bool: True if loading was successful, False otherwise
        """
        try:
            xls = pd.ExcelFile(file_path)
            self.sheet_names = xls.sheet_names
            
            # Load all sheets into a dictionary
            self.sheet_data = {}
            for sheet in self.sheet_names:
                try:
                    df = pd.read_excel(file_path, sheet_name=sheet)
                    # Clean column names: lowercase, remove special chars, replace spaces with underscores
                    df.columns = [re.sub(r'[^\w\s]', '', str(col)).strip().lower().replace(' ', '_') 
                                for col in df.columns]
                    self.sheet_data[sheet] = df
                except Exception as e:
                    st.warning(f"Could not load sheet '{sheet}': {str(e)}")
            
            if not self.sheet_data:
                st.error("No valid sheets found in the Excel file.")
                return False
                
            # For main analysis, use the specified sheet or the first one by default
            if sheet_name and sheet_name in self.sheet_names:
                self.current_sheet = sheet_name
            else:
                self.current_sheet = self.sheet_names[0]
                
            self.data = self.sheet_data[self.current_sheet]
            
            # Perform initial analysis
            self.analyze_columns()
            self.extract_financial_metrics()
            
            return True
            
        except Exception as e:
            st.error(f"Error loading Excel file: {str(e)}")
            return False
    
    def analyze_columns(self) -> None:
        """Analyze all columns in the current dataframe."""
        if self.data is None or self.data.empty:
            return
            
        self.column_analysis = []
        
        for col in self.data.columns:
            col_data = self.data[col].dropna()
            is_numeric = pd.api.types.is_numeric_dtype(col_data)
            
            # Basic column stats
            col_analysis = ColumnAnalysis(
                name=col,
                dtype=str(self.data[col].dtype),
                unique_values=len(self.data[col].unique()),
                sample_values=self._get_sample_values(col_data, is_numeric),
                is_numeric=is_numeric,
                null_count=self.data[col].isnull().sum(),
                null_percentage=round((self.data[col].isnull().sum() / len(self.data)) * 100, 2)
            )
            
            # Additional analysis for numeric columns
            if is_numeric and len(col_data) > 0:
                col_analysis.numeric_stats = {
                    'min': float(col_data.min()),
                    'max': float(col_data.max()),
                    'mean': float(col_data.mean()),
                    'median': float(col_data.median()),
                    'std': float(col_data.std()),
                    'sum': float(col_data.sum())
                }
                
                # Check for currency or percentage
                if any(symbol in col.lower() for symbol in ['$', '€', '£', '¥', 'usd', 'eur', 'gbp']):
                    col_analysis.is_currency = True
                    if '$' in col: col_analysis.currency_symbol = 'USD'
                    elif '€' in col: col_analysis.currency_symbol = 'EUR'
                    elif '£' in col: col_analysis.currency_symbol = 'GBP'
                    elif '¥' in col: col_analysis.currency_symbol = 'JPY'
                
                if '%' in col or 'pct' in col.lower() or 'percent' in col.lower():
                    col_analysis.is_percentage = True
            
            # Check for date columns
            if pd.api.types.is_datetime64_any_dtype(self.data[col]):
                col_analysis.is_date = True
                # Try to infer date format from sample values
                if not col_data.empty and pd.notna(col_data.iloc[0]):
                    try:
                        col_analysis.date_format = self._infer_date_format(col_data.iloc[0])
                    except:
                        pass
            
            # Check for ID columns
            if any(term in col.lower() for term in self.FINANCIAL_TERMS['id']):
                col_analysis.is_id = True
            
            # Check for categorical data
            unique_ratio = col_analysis.unique_values / len(self.data)
            if (0 < unique_ratio <= 0.2) or (col_analysis.unique_values <= 20):
                col_analysis.is_categorical = True
                col_analysis.categories = sorted([str(x) for x in self.data[col].dropna().unique()])
            
            # Identify potential metrics based on column name
            self._identify_potential_metrics(col_analysis)
            
            self.column_analysis.append(col_analysis)
    
    def _get_sample_values(self, series: pd.Series, is_numeric: bool, n: int = 5) -> list:
        """Get sample values from a series."""
        if len(series) == 0:
            return []
            
        # For numeric types, get a sample of distinct values
        if is_numeric:
            unique_vals = series.unique()
            if len(unique_vals) <= n:
                return sorted(unique_vals.tolist())
            else:
                # For many unique values, get min, max, and some in between
                sample = [min(unique_vals), max(unique_vals)]
                remaining = n - 2
                if remaining > 0:
                    step = max(1, len(unique_vals) // (remaining + 1))
                    for i in range(1, remaining + 1):
                        idx = min(i * step, len(unique_vals) - 1)
                        sample.append(unique_vals[idx])
                return sorted(sample)
        else:
            # For non-numeric, just get the first n unique values
            return series.drop_duplicates().head(n).tolist()
    
    def _infer_date_format(self, date_val) -> str:
        """Infer the date format from a date value."""
        if pd.isna(date_val):
            return "Unknown"
            
        if isinstance(date_val, str):
            # Try to parse the string date
            try:
                date_val = pd.to_datetime(date_val)
            except:
                return "Unknown"
        
        # Format the date based on its components
        if date_val.hour != 0 or date_val.minute != 0 or date_val.second != 0:
            return "%Y-%m-%d %H:%M:%S"
        return "%Y-%m-%d"
    
    def _identify_potential_metrics(self, col_analysis: ColumnAnalysis) -> None:
        """Identify potential financial metrics based on column name and content."""
        col_name = col_analysis.name.lower()
        
        # Check for financial terms in column name
        for metric_type, terms in self.FINANCIAL_TERMS.items():
            if any(term in col_name for term in terms):
                col_analysis.potential_metrics.append(metric_type)
        
        # Additional checks for specific financial metrics
        if col_analysis.is_numeric and col_analysis.numeric_stats:
            # Check for monetary values
            if (col_analysis.numeric_stats['min'] >= 0 and 
                col_analysis.numeric_stats['max'] > 1000 and 
                'pct' not in col_name and 
                'ratio' not in col_name):
                col_analysis.potential_metrics.append('monetary_value')
            
            # Check for percentages
            if (0 <= col_analysis.numeric_stats['min'] <= 1 and 
                0 <= col_analysis.numeric_stats['max'] <= 1 and 
                col_analysis.numeric_stats['max'] > 0.1):
                col_analysis.potential_metrics.append('percentage')
        
        # Remove duplicates
        col_analysis.potential_metrics = list(dict.fromkeys(col_analysis.potential_metrics))
    
    def extract_financial_metrics(self) -> None:
        """Extract financial metrics from the analyzed columns."""
        if not self.column_analysis:
            return
        
        # First pass: extract direct metrics from columns
        for col in self.column_analysis:
            if not col.potential_metrics:
                continue
                
            # Get the first non-null value as a sample
            sample_value = None
            if not self.data[col.name].empty:
                non_null = self.data[col.name].dropna()
                if len(non_null) > 0:
                    sample_value = non_null.iloc[0]
            
            # Create metric entries based on potential metrics
            for metric_type in col.potential_metrics:
                if metric_type in ['revenue', 'sales', 'income']:
                    self._update_metric('revenue', sample_value, col.name, 'USD')
                elif metric_type in ['expense', 'cost']:
                    self._update_metric('expense', sample_value, col.name, 'USD')
                elif 'profit' in metric_type:
                    self._update_metric('gross_profit' if 'gross' in metric_type else 'net_income', 
                                      sample_value, col.name, 'USD')
                elif 'asset' in metric_type:
                    self._update_metric('total_assets', sample_value, col.name, 'USD')
                elif 'liability' in metric_type or 'debt' in metric_type:
                    self._update_metric('total_liabilities', sample_value, col.name, 'USD')
                elif 'equity' in metric_type:
                    self._update_metric('total_equity', sample_value, col.name, 'USD')
        
        # Second pass: calculate derived metrics (ratios, etc.)
        self._calculate_derived_metrics()
    
    def _update_metric(self, metric_id: str, value: Any, source: str, unit: str = '') -> None:
        """Update a metric with a new value if it's better than the current one."""
        if pd.isna(value) or value is None:
            return
            
        # Calculate confidence based on source and value
        confidence = 0.7  # Base confidence
        
        # Increase confidence if the metric name is in the source column name
        if metric_id in source.lower():
            confidence += 0.2
            
        # Increase confidence for positive monetary values
        if isinstance(value, (int, float)) and value > 0:
            confidence += 0.1
        
        # Cap confidence at 1.0
        confidence = min(1.0, confidence)
        
        # Only update if we have higher confidence or no previous value
        if (metric_id not in self.metrics or 
            confidence > self.metrics[metric_id].confidence or 
            self.metrics[metric_id].value is None):
                
            self.metrics[metric_id] = FinancialMetric(
                name=metric_id.replace('_', ' ').title(),
                value=value,
                unit=unit,
                source=source,
                confidence=confidence
            )
    
    def _calculate_derived_metrics(self) -> None:
        """Calculate derived financial metrics based on existing metrics."""
        # Calculate financial ratios
        for ratio_id, ratio_info in self.FINANCIAL_RATIOS.items():
            try:
                # Skip if we already have this ratio
                if ratio_id in self.metrics and self.metrics[ratio_id].value is not None:
                    continue
                    
                # Evaluate the formula
                formula = ratio_info['formula']
                required_metrics = [m for m in self.metrics.keys() if m in formula]
                
                # Check if we have all required metrics
                if all(m in self.metrics and self.metrics[m].value is not None for m in required_metrics):
                    # Create a local namespace with metric values
                    local_vars = {m: self.metrics[m].value for m in required_metrics}
                    
                    # Safely evaluate the formula
                    try:
                        value = eval(formula, {"__builtins__": {}}, local_vars)
                        
                        # Add the derived metric
                        self.metrics[ratio_id] = FinancialMetric(
                            name=ratio_info['name'],
                            value=value,
                            unit='%' if 'margin' in ratio_id.lower() or 'return' in ratio_id.lower() else '',
                            description=ratio_info['description'],
                            category=ratio_info['category'],
                            importance=ratio_info.get('importance', 'medium'),
                            is_derived=True,
                            formula=formula,
                            confidence=0.9  # High confidence for calculated metrics
                        )
                    except Exception as e:
                        st.warning(f"Could not calculate {ratio_id}: {str(e)}")
                        
            except Exception as e:
                st.warning(f"Error processing ratio {ratio_id}: {str(e)}")
    
    def get_metrics_by_category(self) -> Dict[str, List[FinancialMetric]]:
        """Group metrics by category for display."""
        categories = {}
        for metric in self.metrics.values():
            if metric.value is None:
                continue
                
            if metric.category not in categories:
                categories[metric.category] = []
            categories[metric.category].append(metric)
            
        return categories
    
    def get_data_summary(self) -> Dict[str, Any]:
        """Generate a summary of the data and analysis."""
        if self.data is None:
            return {}
            
        return {
            'file_info': {
                'sheets': self.sheet_names,
                'current_sheet': self.current_sheet,
                'total_rows': len(self.data),
                'total_columns': len(self.data.columns),
                'total_numeric_columns': sum(1 for col in self.column_analysis if col.is_numeric),
                'total_date_columns': sum(1 for col in self.column_analysis if col.is_date)
            },
            'metrics_summary': {
                'total_metrics': len([m for m in self.metrics.values() if m.value is not None]),
                'categories': list(set(m.category for m in self.metrics.values() if m.value is not None))
            }
        }
    
    def generate_visualizations(self) -> List[go.Figure]:
        """Generate visualizations for the analyzed data."""
        if self.data is None or not hasattr(self, 'column_analysis'):
            return []
            
        figures = []
        
        # 1. Time series plot for date columns with numeric data
        date_cols = [col.name for col in self.column_analysis if col.is_date]
        numeric_cols = [col.name for col in self.column_analysis if col.is_numeric]
        
        if date_cols and numeric_cols:
            # Use the first date column as x-axis
            date_col = date_cols[0]
            
            # Create a line plot for each numeric column over time
            for num_col in numeric_cols[:5]:  # Limit to first 5 numeric columns
                try:
                    fig = px.line(
                        self.data.sort_values(date_col),
                        x=date_col,
                        y=num_col,
                        title=f"{num_col.replace('_', ' ').title()} Over Time"
                    )
                    fig.update_layout(
                        xaxis_title=date_col.replace('_', ' ').title(),
                        yaxis_title=num_col.replace('_', ' ').title(),
                        showlegend=False
                    )
                    figures.append(fig)
                except:
                    continue
        
        # 2. Bar chart for categorical data
        categorical_cols = [col.name for col in self.column_analysis 
                          if col.is_categorical and 1 < col.unique_values <= 10]
        
        for cat_col in categorical_cols[:3]:  # Limit to first 3 categorical columns
            try:
                fig = px.bar(
                    self.data[cat_col].value_counts().reset_index(),
                    x='index',
                    y=cat_col,
                    title=f"Distribution of {cat_col.replace('_', ' ').title()}",
                    labels={'index': cat_col.replace('_', ' ').title(), cat_col: 'Count'}
                )
                figures.append(fig)
            except:
                continue
        
        # 3. Pie chart for a categorical column with percentages
        if categorical_cols:
            cat_col = categorical_cols[0]
            try:
                counts = self.data[cat_col].value_counts(normalize=True) * 100
                fig = px.pie(
                    values=counts.values,
                    names=counts.index,
                    title=f"Percentage Distribution of {cat_col.replace('_', ' ').title()}"
                )
                figures.append(fig)
            except:
                pass
        
        return figures

def main():
    """Main Streamlit application for Excel AI Analyzer."""
    st.title("📊 Excel AI Analyzer")
    st.markdown("Upload an Excel file to analyze its financial data and extract key metrics.")
    
    # Initialize session state
    if 'analyzer' not in st.session_state:
        st.session_state.analyzer = ExcelAIAnalyzer()
    
    # File uploader
    uploaded_file = st.file_uploader("Choose an Excel file", type=["xlsx", "xls"])
    
    if uploaded_file is not None:
        try:
            with st.spinner("Analyzing your Excel file..."):
                # Load and analyze the Excel file
                if st.session_state.analyzer.load_excel(uploaded_file):
                    st.success("File loaded and analyzed successfully!")
                    
                    # Show file summary
                    summary = st.session_state.analyzer.get_data_summary()
                    
                    # Display summary in expandable section
                    with st.expander("📄 File Summary", expanded=True):
                        col1, col2, col3 = st.columns(3)
                        with col1:
                            st.metric("Sheets", len(summary['file_info']['sheets']))
                        with col2:
                            st.metric("Rows", summary['file_info']['total_rows'])
                        with col3:
                            st.metric("Columns", summary['file_info']['total_columns'])
                    
                    # Show available sheets with a selector
                    if len(st.session_state.analyzer.sheet_names) > 1:
                        selected_sheet = st.selectbox(
                            "Select a sheet to analyze:",
                            st.session_state.analyzer.sheet_names,
                            index=st.session_state.analyzer.sheet_names.index(st.session_state.analyzer.current_sheet)
                        )
                        
                        if selected_sheet != st.session_state.analyzer.current_sheet:
                            st.session_state.analyzer.load_excel(uploaded_file, selected_sheet)
                    
                    # Tabs for different sections
                    tab1, tab2, tab3, tab4 = st.tabs(["📊 Metrics", "📈 Visualizations", "🔍 Data Explorer", "📝 Column Analysis"])
                    
                    with tab1:
                        st.header("Extracted Financial Metrics")
                        
                        # Group metrics by category
                        metrics_by_category = st.session_state.analyzer.get_metrics_by_category()
                        
                        if not metrics_by_category:
                            st.warning("No financial metrics were automatically detected in this file.")
                        else:
                            for category, metrics in metrics_by_category.items():
                                with st.expander(f"{category} ({len(metrics)} metrics)", expanded=True):
                                    # Sort metrics by importance (high to low)
                                    metrics_sorted = sorted(
                                        metrics,
                                        key=lambda x: (x.importance == 'high', x.importance == 'medium', x.importance == 'low'),
                                        reverse=True
                                    )
                                    
                                    # Display metrics in columns
                                    cols = st.columns(3)
                                    for i, metric in enumerate(metrics_sorted):
                                        with cols[i % 3]:
                                            # Format the value based on its type
                                            value = metric.value
                                            if isinstance(value, (int, float)):
                                                if abs(value) >= 1_000_000:
                                                    value_str = f"${value/1_000_000:,.2f}M"
                                                elif abs(value) >= 1_000:
                                                    value_str = f"${value/1_000:,.2f}K"
                                                else:
                                                    value_str = f"${value:,.2f}"
                                                
                                                # Add percentage sign if it's a percentage
                                                if metric.unit == '%':
                                                    value_str = f"{value*100:.2f}%"
                                            else:
                                                value_str = str(value)
                                            
                                            st.metric(
                                                label=metric.name,
                                                value=value_str,
                                                help=f"{metric.description} (Source: {metric.source})"
                                            )
                    
                    with tab2:
                        st.header("Data Visualizations")
                        
                        # Generate and display visualizations
                        with st.spinner("Generating visualizations..."):
                            figures = st.session_state.analyzer.generate_visualizations()
                            
                            if not figures:
                                st.info("No visualizations could be generated. This might be due to insufficient data or incompatible column types.")
                            else:
                                for fig in figures:
                                    st.plotly_chart(fig, use_container_width=True)
                    
                    with tab3:
                        st.header("Data Explorer")
                        
                        # Show the raw data with filters
                        st.dataframe(
                            st.session_state.analyzer.data,
                            use_container_width=True,
                            hide_index=True
                        )
                    
                    with tab4:
                        st.header("Column Analysis")
                        
                        # Display detailed column analysis
                        for col in st.session_state.analyzer.column_analysis:
                            with st.expander(f"{col.name} ({col.dtype})", expanded=False):
                                col1, col2 = st.columns([1, 2])
                                
                                with col1:
                                    st.markdown("**Basic Information**")
                                    st.write(f"**Type:** {col.dtype}")
                                    st.write(f"**Unique Values:** {col.unique_values}")
                                    st.write(f"**Null Values:** {col.null_count} ({col.null_percentage}%)")
                                    
                                    if col.is_numeric and col.numeric_stats:
                                        st.markdown("**Numeric Statistics**")
                                        st.write(f"Min: {col.numeric_stats['min']:,.2f}")
                                        st.write(f"Max: {col.numeric_stats['max']:,.2f}")
                                        st.write(f"Mean: {col.numeric_stats['mean']:,.2f}")
                                        st.write(f"Median: {col.numeric_stats['median']:,.2f}")
                                        st.write(f"Sum: {col.numeric_stats['sum']:,.2f}")
                                        
                                    if col.potential_metrics:
                                        st.markdown("**Potential Metrics**")
                                        for metric in col.potential_metrics:
                                            st.write(f"- {metric.replace('_', ' ').title()}")
                                
                                with col2:
                                    st.markdown("**Sample Values**")
                                    if col.sample_values:
                                        st.code("\n".join(map(str, col.sample_values[:10])))  # Show first 10 samples
                                    
                                    # Show a small preview of the data distribution
                                    if col.is_numeric and len(col.sample_values) > 1:
                                        try:
                                            fig = px.histogram(
                                                st.session_state.analyzer.data[col.name].dropna(),
                                                title=f"Distribution of {col.name}",
                                                labels={'value': col.name}
                                            )
                                            st.plotly_chart(fig, use_container_width=True, use_container_height=200)
                                        except Exception as e:
                                            st.warning(f"Could not generate distribution: {str(e)}")
                else:
                    st.error("Failed to analyze the Excel file. Please check the file format and try again.")
        except Exception as e:
            st.error(f"An error occurred: {str(e)}")
            st.exception(e)  # Show detailed error for debugging
    else:
        # Show welcome/instructions when no file is uploaded
        st.markdown("""
        ### How to use this tool:
        1. Upload an Excel file containing financial data
        2. The tool will automatically analyze the data and extract key metrics
        3. Explore the different tabs to view:
           - 📊 **Metrics**: Extracted financial metrics and KPIs
           - 📈 **Visualizations**: Interactive charts and graphs
           - 🔍 **Data Explorer**: Raw data with filtering options
           - 📝 **Column Analysis**: Detailed analysis of each column
        
        ### Supported Data Types:
        - Numeric data (revenue, costs, etc.)
        - Date fields (for time series analysis)
        - Categorical data (regions, departments, etc.)
        - Currency values
        
        ### Tips for Best Results:
        - Ensure your data has clear column headers
        - Use consistent formatting for dates and currencies
        - Remove any empty rows or columns before uploading
        """)

if __name__ == "__main__":
    main()
