"""
Financial Statements Generator
==============================
Generates P&L statements and calculates key financial metrics

Business Logic:
- Groups GL transactions by account type and period
- Calculates standard financial metrics (margins, ratios)
- Supports monthly, quarterly, and annual reporting
"""

import pandas as pd
import numpy as np
from datetime import datetime
import yaml


class FinancialStatements:
    """
    Automated P&L and metrics generator

    -------------
    1. Takes raw GL transactions
    2. Groups them by time period (month/quarter/year)
    3. Calculates Revenue, COGS, OpEx
    4. Computes financial metrics (margins, ratios)
    5. Returns structured financial statements
    """
    
    def __init__(self, gl_df, config_path='config/settings.yaml'):
        """
        Initialize the financial statements generator
        ------------------
        1. Makes a copy of data
        2. Loads configuration settings
        3. Converts dates to proper datetime format
        4. Stores data range for reference
        
        Parameters:
        -----------
        gl_df : DataFrame
            General Ledger transactions with columns:
            - date, account_type, account_name, amount
        
        config_path : str
            Path to settings.yaml configuration file
        """
        # Step 1: Make a safe copy (don't modify original data)
        self.gl_df = gl_df.copy()
        
        # Step 2: Load configuration
        with open(config_path, 'r') as f:
            self.config = yaml.safe_load(f)
        
        # Step 3: Ensure dates are datetime objects (for time-based grouping)
        date_col = self.config['data']['date_column']
        self.gl_df[date_col] = pd.to_datetime(self.gl_df[date_col])
        
        # Step 4: Store the data range (useful for reports)
        self.start_date = self.gl_df[date_col].min()
        self.end_date = self.gl_df[date_col].max()
        
        print(f"Financial Statements module initialized")
        print(f"  Data range: {self.start_date.strftime('%Y-%m-%d')} to {self.end_date.strftime('%Y-%m-%d')}")
        print(f"  Total transactions: {len(self.gl_df):,}")
    def generate_pnl(self, start_date=None, end_date=None, period='monthly'):
        """
        Generate Profit & Loss Statement
        -------------------------------
        1. Filter data by date range (if specified)
        2. Keep only P&L accounts (Revenue, COGS, OpEx)
        3. Add period column (month/quarter/year)
        4. Group transactions by period and account
        5. Sum amounts for each group
        6. Pivot into wide format (periods as columns)
        
        Parameters:
        -----------
        start_date : str or datetime, optional
            Filter to start from this date
            Example: '2024-01-01' or datetime(2024, 1, 1)
            
        end_date : str or datetime, optional
            Filter to end at this date
            
        period : str
            How to group time: 'monthly', 'quarterly', or 'annual'
            
        Returns:
        --------
        DataFrame with P&L organized by period
        Rows = Accounts (Revenue, COGS, OpEx)
        Columns = Time periods (Jan, Feb, Mar...)
        
        Example Output:
        ---------------
                              2024-01    2024-02    2024-03
        Revenue
          Product Sales       100,000    120,000    110,000
          Service Revenue      50,000     55,000     60,000
        COGS
          Cost of Products     40,000     48,000     44,000
        OpEx
          Salaries            30,000     30,000     30,000
        """
        
        # STEP 1: Create working copy
        df = self.gl_df.copy()
        date_col = self.config['data']['date_column']
        
        # STEP 2: Filter by date range (if user specified)
        if start_date:
            df = df[df[date_col] >= pd.to_datetime(start_date)]
            print(f"  Filtered from: {start_date}")
            
        if end_date:
            df = df[df[date_col] <= pd.to_datetime(end_date)]
            print(f"  Filtered to: {end_date}")
        
        # STEP 3: Filter to P&L accounts only
        # P&L = Profit & Loss = Revenue - COGS - OpEx
        pnl_df = df[df['account_type'].isin(['Revenue', 'COGS', 'OpEx'])].copy()
        
        print(f"  P&L transactions: {len(pnl_df):,} (from {len(df):,} total)")
        
        # STEP 4: Add period column based on user choice
        if period == 'monthly':
            # Convert date to month period (2024-01-15 → 2024-01)
            pnl_df['period'] = pnl_df[date_col].dt.to_period('M')
            print(f"  Grouping by: Month")
            
        elif period == 'quarterly':
            # Convert to quarter period (2024-01-15 → 2024Q1)
            pnl_df['period'] = pnl_df[date_col].dt.to_period('Q')
            print(f"  Grouping by: Quarter")
            
        elif period == 'annual':
            # Convert to year period (2024-01-15 → 2024)
            pnl_df['period'] = pnl_df[date_col].dt.to_period('Y')
            print(f"  Grouping by: Year")
            
        else:
            raise ValueError(f"period must be 'monthly', 'quarterly', or 'annual', got '{period}'")
        
        # STEP 5: Group and sum
        # This is like: SELECT period, account_type, SUM(amount) GROUP BY period, account_type
        pnl_summary = pnl_df.groupby([
            'period',           # Group by time period
            'account_type',     # Group by Revenue/COGS/OpEx
            'account_category', # Group by Sales/Marketing/etc
            'account_name'      # Group by specific account
        ])['amount'].sum().reset_index()
        
        # STEP 6: Pivot to wide format
        # Before: One row per period-account combo
        # After: Accounts as rows, periods as columns
        pnl_pivot = pnl_summary.pivot_table(
            index=['account_type', 'account_category', 'account_name'],  # Rows
            columns='period',      # Columns
            values='amount',       # Cell values
            fill_value=0          # If no data for a period, show 0
        )
        
        print(f"P&L generated: {len(pnl_pivot)} accounts × {len(pnl_pivot.columns)} periods")
        
        return pnl_pivot
    def calculate_metrics(self, pnl_df):
        """
        Calculate key financial metrics from P&L

        1. Revenue (total sales)
        2. COGS (cost to deliver)
        3. Gross Profit (Revenue - COGS)
        4. Gross Margin % (Gross Profit / Revenue × 100)
        5. Operating Expenses (overhead costs)
        6. Operating Income / EBITDA (Gross Profit - OpEx)
        7. Operating Margin % (Operating Income / Revenue × 100)
        
        Parameters:
        -----------
        pnl_df : DataFrame
            P&L from generate_pnl() with periods as columns
            
        Returns:
        --------
        DataFrame with metrics for each period
        
        Example Output:
        ---------------
        Metric                     2024-01    2024-02    2024-03
        Revenue                    125,000    150,000    140,000
        COGS                       20,000     25,000     23,000
        Gross Profit               105,000    125,000    117,000
        Gross Margin %             84.0%      83.3%      83.6%
        Operating Expenses         30,000     32,000     31,000
        Operating Income (EBITDA)  75,000     93,000     86,000
        Operating Margin %         60.0%      62.0%      61.4%
        """
        
        # Dictionary to store all metrics for all periods
        metrics = {}
        
        # Loop through each time period (each column in P&L)
        for period in pnl_df.columns:
            print(f"  Calculating metrics for: {period}")
            
            # Get all amounts for this specific period
            period_data = pnl_df[period]
            
            # METRIC 1: Total Revenue
            # Sum all accounts where account_type = 'Revenue'
            # .index.get_level_values(0) gets the first level of multi-index (account_type)
            revenue = period_data[
                period_data.index.get_level_values(0) == 'Revenue'
            ].sum()
            
            # METRIC 2: Total COGS
            cogs = period_data[
                period_data.index.get_level_values(0) == 'COGS'
            ].sum()
            
            # METRIC 3: Total Operating Expenses
            opex = period_data[
                period_data.index.get_level_values(0) == 'OpEx'
            ].sum()
            
            # METRIC 4: Gross Profit (Revenue minus cost to deliver)
            gross_profit = revenue - cogs
            
            # METRIC 5: Gross Margin % (what % of revenue is profit after COGS)
            # Avoid division by zero with conditional
            if revenue > 0:
                gross_margin_pct = (gross_profit / revenue) * 100
            else:
                gross_margin_pct = 0
            
            # METRIC 6: Operating Income / EBITDA (profit after all expenses)
            operating_income = gross_profit - opex
            
            # METRIC 7: Operating Margin % (what % of revenue is final profit)
            if revenue > 0:
                operating_margin_pct = (operating_income / revenue) * 100
            else:
                operating_margin_pct = 0
            
            # Store all metrics for this period
            metrics[str(period)] = {
                'Revenue': revenue,
                'COGS': cogs,
                'Gross Profit': gross_profit,
                'Gross Margin %': gross_margin_pct,
                'Operating Expenses': opex,
                'Operating Income (EBITDA)': operating_income,
                'Operating Margin %': operating_margin_pct
            }
        
        # Convert dictionary to DataFrame
        # .T = Transpose (flip rows and columns)
        # Before: {period1: {metrics}, period2: {metrics}}
        # After: DataFrame with periods as rows, metrics as columns
        metrics_df = pd.DataFrame(metrics).T
        
        print(f"Calculated {len(metrics_df.columns)} metrics for {len(metrics_df)} periods")
        
        return metrics_df
    def export_to_excel(self, pnl_df, metrics_df, filename):
        """
        Export P&L and metrics to Excel file
        ------------------
        Excel workbook with 2 sheets:
        1. "P&L Detail" - Full account-level breakdown
        2. "Key Metrics" - Summary metrics
        
        Parameters:
        -----------
        pnl_df : DataFrame
            P&L from generate_pnl()
        metrics_df : DataFrame
            Metrics from calculate_metrics()
        filename : str
            Output filename (e.g., 'monthly_pnl_202401.xlsx')
        """
        
        # Create Excel file with multiple sheets
        with pd.ExcelWriter(filename, engine='openpyxl') as writer:
            # Sheet 1: Detailed P&L
            pnl_df.to_excel(writer, sheet_name='P&L Detail')

            # Sheet 2: Summary Metrics (transpose so metric names are first column)
            metrics_long = metrics_df.transpose().reset_index()
            metrics_long.columns = ['Metric'] + [str(col) for col in metrics_long.columns[1:]]
            metrics_long.to_excel(writer, sheet_name='Key Metrics', index=False)

        print(f"Exported to: {filename}")
        print(f"  - P&L Detail: {len(pnl_df)} accounts")
        print(f"  - Key Metrics: {len(metrics_df)} periods")


# Demo/test function
if __name__ == "__main__":
    """
    Test the FinancialStatements module
    
    This runs when you execute this file directly:
    python src/analytics/financial_statements.py
    """
    import sys
    sys.path.append('.')
    from src.core.data_loader import DataLoader
    
    print("="*60)
    print("TESTING FINANCIAL STATEMENTS MODULE")
    print("="*60)
    
    # Load data
    loader = DataLoader()
    gl_df = loader.load_gl_transactions()
    
    # Create financial statements
    fs = FinancialStatements(gl_df)
    
    # Generate monthly P&L
    print("\nGenerating monthly P&L...")
    pnl = fs.generate_pnl(period='monthly')
    print("\nFirst 10 accounts:")
    print(pnl.head(10))
    
    # Calculate metrics
    print("\nCalculating metrics...")
    metrics = fs.calculate_metrics(pnl)
    print("\nKey Metrics:")
    print(metrics)
    
    # Export
    print("\nExporting to Excel...")
    fs.export_to_excel(pnl, metrics, 'outputs/reports/test_pnl.xlsx')
    
    print("\n" + "="*60)
    print("TEST COMPLETE!")
    print("="*60)