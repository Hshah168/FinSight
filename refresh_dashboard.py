
"""
One-Click Dashboard Refresh
============================
We run this to update Power BI with new data

Usage:
    python refresh_dashboard.py
"""

import os
from matplotlib.pylab import var
import pandas as pd
from datetime import datetime

from sklearn import metrics

from src.analytics.variance_analysis import VarianceAnalyzer

print("="*60)
print("FINSIGHT DASHBOARD REFRESH")
print("="*60)

# --- Power BI Data Preparation Logic ---

class PowerBIPrep:
    """
    Prepare data for Power BI consumption
    Create the following tables in data/powerbi/:
    --------
    1. Fact_Transactions - Transaction-level data
    2. Dim_Accounts - Account master data
    3. Dim_Date - Date dimension
    4. Fact_Budget - Budget data
    5. Metrics_Summary - Pre-calculated KPIs
    6. Variance_Analysis - Budget vs Actual
    """
    def __init__(self, output_folder='data/powerbi'):
        self.output_folder = output_folder
        os.makedirs(output_folder, exist_ok=True)
        print(f"Power BI Prep initialized")
        print(f"  Output folder: {output_folder}")

    def prepare_all_tables(self, gl_df, budget_df):

        print("\n" + "="*60)
        print("PREPARING POWER BI DATA TABLES")
        print("="*60)

        # Table 1: Fact Transactions
        print("\n1. Creating Fact_Transactions...")
        fact_transactions = self._create_fact_transactions(gl_df)
        fact_transactions.to_csv(f'{self.output_folder}/Fact_Transactions.csv', index=False)
        print(f"   Saved: {len(fact_transactions):,} rows")

        # Table 2: Dim Accounts
        print("\n2. Creating Dim_Accounts...")
        dim_accounts = self._create_dim_accounts(gl_df)
        dim_accounts.to_csv(f'{self.output_folder}/Dim_Accounts.csv', index=False)
        print(f"   Saved: {len(dim_accounts):,} accounts")

        # Table 3: Dim Date
        print("\n3. Creating Dim_Date...")
        dim_date = self._create_dim_date(gl_df)
        dim_date.to_csv(f'{self.output_folder}/Dim_Date.csv', index=False)
        print(f"   Saved: {len(dim_date):,} dates")
        
        # Table 4: Fact Budget
        print("\n4. Creating Fact_Budget...")
        #print dataframe
        print(budget_df.head())
        fact_budget = self._create_fact_budget(budget_df)
        fact_budget.to_csv(f'{self.output_folder}/Fact_Budget.csv', index=False)
        print(f"   Saved: {len(fact_budget):,} rows")

        # Table 5: Metrics Summary
        print("\n5. Creating Metrics_Summary...")
        metrics = self._create_metrics_summary(gl_df, budget_df)
        metrics.to_csv(f'{self.output_folder}/Metrics_Summary.csv', index=False)
        print(f"   Saved: {len(metrics):,} periods")
        print(gl_df['account_type'].unique())

        # Table 6: Variance Analysis
        print("\n6. Creating Variance_Analysis...")
        variance = self._create_variance_table(gl_df, budget_df)
        variance.to_csv(f'{self.output_folder}/Variance_Analysis.csv', index=False)
        print(f"   Saved: {len(variance):,} variances")

        print("\n" + "="*60)
        print("ALL TABLES CREATED")
        print("="*60)
        print(f"\n Files saved to: {self.output_folder}/")
        print("\nNext step: Import these CSV files into Power BI")

    def _create_fact_transactions(self, gl_df):
        df = gl_df.copy()
        df['date'] = pd.to_datetime(df['date'])
        df['Year'] = df['date'].dt.year
        df['Quarter'] = df['date'].dt.quarter
        df['Month'] = df['date'].dt.month
        df['MonthName'] = df['date'].dt.strftime('%B')
        df['YearMonth'] = df['date'].dt.strftime('%Y-%m')
        df['YearQuarter'] = df['date'].dt.year.astype(str) + '-Q' + df['date'].dt.quarter.astype(str)
        fact = df[[
            'transaction_id',
            'date',
            'account_code',
            'amount',
            'department',
            'cost_center',
            'Year',
            'Quarter',
            'Month',
            'MonthName',
            'YearMonth',
            'YearQuarter'
        ]].copy()
        fact.columns = [
            'TransactionID',
            'Date',
            'AccountCode',
            'Amount',
            'Department',
            'CostCenter',
            'Year',
            'Quarter',
            'Month',
            'MonthName',
            'YearMonth',
            'YearQuarter'
        ]
        return fact

    def _create_dim_accounts(self, gl_df):
        df = gl_df.copy()
        dim = df[['account_code', 'account_name', 'account_type', 'account_category']].drop_duplicates()
        dim.columns = ['AccountCode', 'Account_Name', 'Account_Type', 'AccountCategory']
        return dim

    def _create_dim_date(self, gl_df):
        """
        Create date dimension for Power BI
        Columns:
        - Date (primary key)
        - Year, Quarter, Month, MonthName
        - YearMonth, YearQuarter
        - DayOfWeek, DayName, IsWeekend
        - IsCurrentYear, IsCurrentQuarter, IsCurrentMonth
        """
        min_date = pd.to_datetime(gl_df['date']).min()
        max_date = pd.to_datetime(gl_df['date']).max()
        date_range = pd.date_range(start=min_date, end=max_date, freq='D')
        dim = pd.DataFrame({
            'Date': date_range,
            'Year': date_range.year,
            'Quarter': date_range.quarter,
            'Month': date_range.month,
            'MonthName': date_range.strftime('%B'),
            'YearMonth': date_range.strftime('%Y-%m'),
            'YearQuarter': date_range.year.astype(str) + '-Q' + date_range.quarter.astype(str),
            'DayOfWeek': date_range.dayofweek,
            'DayName': date_range.strftime('%A'),
            'IsWeekend': date_range.dayofweek.isin([5, 6])
        })
        today = pd.Timestamp.now()
        dim['IsCurrentYear'] = (dim['Year'] == today.year)
        dim['IsCurrentQuarter'] = ((dim['Year'] == today.year) & (dim['Quarter'] == today.quarter))
        dim['IsCurrentMonth'] = ((dim['Year'] == today.year) & (dim['Month'] == today.month))
        return dim

    def _create_fact_budget(self, budget_df):
        df = budget_df.copy()
        df['date'] = pd.to_datetime(df['date'])
        df['Year'] = df['date'].dt.year
        df['Quarter'] = df['date'].dt.quarter
        df['Month'] = df['date'].dt.month
        df['MonthName'] = df['date'].dt.strftime('%B')
        df['YearMonth'] = df['date'].dt.strftime('%Y-%m')
        df['YearQuarter'] = df['date'].dt.year.astype(str) + '-Q' + df['date'].dt.quarter.astype(str)

        print("Budget Data columns:", df.columns.tolist())
        
        # Use only columns that exist in budget_data.csv
        fact = df[[
            'date',
            'yearmonth',
            'year',
            'month',
            'quarter',
            'account_code',
            'account_name',
            'department',
            'budgetamount',
            'account_type',
            'account_category'
        ]].copy()
        # Rename columns for Power BI convention
        fact = fact.rename(columns={
            'date': 'Date',
            'account_code': 'AccountCode',
            'account_name': 'AccountName',
            'department': 'Department',
            'budgetamount': 'BudgetAmount',
            'account_type': 'AccountType',
            'account_category': 'AccountCategory',
            'year': 'Year',
            'month': 'Month',
            'quarter': 'Quarter',
            'yearmonth': 'YearMonth',
        })
        return fact

    
    def _create_metrics_summary(self, gl_df, budget_df):
    
        gl_df = gl_df.copy()
        gl_df.columns = gl_df.columns.str.strip().str.lower()

        gl_df['date'] = pd.to_datetime(gl_df['date'])
        gl_df['Period'] = gl_df['date'].dt.strftime('%Y-%m')

        # --- NORMALIZE ACCOUNT TYPE ---
        gl_df['account_type'] = gl_df['account_type'].str.strip().str.lower()

        # Map common variations
        gl_df['account_type_mapped'] = gl_df['account_type'].replace({
            'revenue': 'revenue',
            'cogs': 'cogs',
            'cost of goods sold': 'cogs',
            'opex': 'expense',
            'operating expense': 'expense',
            'expenses': 'expense'
        })

        # --- AGGREGATE ACTUALS ---
        grouped_actual = (
            gl_df.groupby(['Period', 'account_type_mapped'])['amount']
            .sum()
            .reset_index()
        )

        pivot_actual = grouped_actual.pivot(
            index='Period',
            columns='account_type_mapped',
            values='amount'
        ).fillna(0).reset_index()

        # Ensure required columns exist
        for col in ['revenue', 'cogs', 'expense']:
            if col not in pivot_actual.columns:
                pivot_actual[col] = 0

        # BASIC METRICS 
        pivot_actual['Total Revenue'] = pivot_actual['revenue']
        pivot_actual['Total Expense'] = pivot_actual['expense']
        pivot_actual['COGS'] = pivot_actual['cogs']

        pivot_actual['Gross Profit'] = pivot_actual['revenue'] - pivot_actual['cogs']

        pivot_actual['Gross Margin %'] = (
            pivot_actual['Gross Profit'] / pivot_actual['revenue']
        ).replace([float('inf'), -float('inf')], 0).fillna(0)

        pivot_actual['EBITDA'] = (
            pivot_actual['revenue']
            - pivot_actual['cogs']
            - pivot_actual['expense']
        )

        # REVENUE GROWTH
        pivot_actual = pivot_actual.sort_values('Period')
        pivot_actual['Revenue Growth %'] = (
            pivot_actual['Total Revenue']
            .pct_change()
            .fillna(0)
        )

        # BUDGET VARIANCE
        budget_df = budget_df.copy()
        budget_df.columns = budget_df.columns.str.strip().str.lower()
        budget_df['date'] = pd.to_datetime(budget_df['date'])
        budget_df['Period'] = budget_df['date'].dt.strftime('%Y-%m')

        if 'budgetamount' in budget_df.columns:
            budget_df.rename(columns={'budgetamount': 'budget_amount'}, inplace=True)

        budget_grouped = (
            budget_df.groupby('Period')['budget_amount']
            .sum()
            .reset_index()
            .rename(columns={'budget_amount': 'Budget'})
        )

        pivot_actual = pivot_actual.merge(budget_grouped, on='Period', how='left')
        pivot_actual['Budget'] = pivot_actual['Budget'].fillna(0)

        pivot_actual['Variance vs Budget'] = (
            pivot_actual['Total Revenue'] - pivot_actual['Budget']
        )

        #STACK FORMAT FOR POWER BI
        metric_columns = [
            'Total Revenue',
            'Total Expense',
            'COGS',
            'Gross Profit',
            'Gross Margin %',
            'EBITDA',
            'Revenue Growth %',
            'Variance vs Budget'
        ]

        metrics_list = []

        for metric in metric_columns:
            temp = pivot_actual[['Period', metric]].copy()
            temp['Metric'] = metric
            temp.rename(columns={metric: 'Amount'}, inplace=True)
            metrics_list.append(temp)

        metrics = pd.concat(metrics_list, ignore_index=True)
        metrics = metrics[['Period', 'Metric', 'Amount']]

        return metrics

    def _create_variance_table(self, gl_df, budget_df):

        # STANDARDIZE GL
        gl_df = gl_df.copy()
        gl_df.columns = gl_df.columns.str.strip().str.lower()

        if 'date' in gl_df.columns:
            gl_df['date'] = pd.to_datetime(gl_df['date'])
            gl_df['YearMonth'] = gl_df['date'].dt.strftime('%Y-%m')

        #STANDARDIZE BUDGET
        budget_df = budget_df.copy()
        budget_df.columns = budget_df.columns.str.strip().str.lower()

        if 'date' in budget_df.columns:
            budget_df['date'] = pd.to_datetime(budget_df['date'])
            budget_df['YearMonth'] = budget_df['date'].dt.strftime('%Y-%m')

        if 'budgetamount' in budget_df.columns:
            budget_df.rename(columns={'budgetamount': 'budget_amount'}, inplace=True)

        if 'account_code' not in gl_df.columns:
            raise Exception("GL missing account_code")

        if 'account_code' not in budget_df.columns:
            raise Exception("Budget missing account_code")

        # Now analyzer receives clean schema
        va = VarianceAnalyzer(gl_df, budget_df)

        variance = va.calculate_variances(YearMonth='monthly')
        variance_flagged = va.flag_significant_variances(variance)

        return variance_flagged

# DataLoader
class DataLoader:
    def __init__(self):
        self.base_path = os.path.join(os.getcwd(), "data")
        self.gl_path = os.path.join(self.base_path, "gl_transactions.csv")
        self.budget_path = os.path.join(self.base_path, "budget_data.csv")
        self.coa_path = os.path.join(self.base_path, "chart_of_accounts.csv")
        self.dept_path = os.path.join(self.base_path, "department_mapping.csv")
        print("DataLoader initialized")

    def load_gl_transactions(self):
        print(f"Loading GL transactions from: {self.gl_path}")
        try:
            df = pd.read_csv(self.gl_path)
            print(f"Loaded {len(df):,} transactions")
            return df
        except FileNotFoundError as e:
            print(f"Error loading file: {e}")
            print("Trying alternative formats...")
            return pd.DataFrame()

    def load_budget_data(self):
        print("Loading Budget Data...")
        df = pd.read_csv(self.budget_path)
        print(f"Loaded {len(df)} Budget records")
        return df

# End DataLoader

# Step 1: Prepare Power BI data
print("\n1. Preparing Power BI data...")
loader = DataLoader()
gl_df = loader.load_gl_transactions()
budget_df = loader.load_budget_data()
prep = PowerBIPrep(output_folder='data/powerbi')
prep.prepare_all_tables(gl_df, budget_df)
print("   Data preparation complete")

# Step 2: Open Power BI (if installed)
print("\n2. Opening Power BI Desktop...")
powerbi_path = r"C:\Program Files\Microsoft Power BI Desktop\bin\PBIDesktop.exe"
template_path = r"templates\FinSight_Dashboard_Template.pbix"

if os.path.exists(powerbi_path) and os.path.exists(template_path):
    os.startfile(template_path)
    print("   Power BI opened")
    print("\n3. In Power BI: Click 'Refresh' button")
else:
    print("    Power BI not found or template missing")
    print(f"   Expected: {powerbi_path}")
    print(f"\n   Please open manually: {template_path}")

print("\n" + "="*60)
print("REFRESH PROCESS COMPLETE")
print("="*60)

# Quick Reference Card
print("""
╔════════════════════════════════════════════════════════════╗
║          FINSIGHT MONTHLY DASHBOARD UPDATE                 ║
╠════════════════════════════════════════════════════════════╣
║                                                            ║
║  1. Place new CSV files in: data/input/                    ║
║     - gl_transactions.csv                                  ║
║     - budget_data.csv                                      ║
║                                                            ║
║  2. Run: python refresh_dashboard.py                       ║
║                                                            ║
║  3. In Power BI: Click "Refresh" button                    ║
║                                                            ║
║  4. Dashboard updated with new data                        ║
║                                                            ║
╚════════════════════════════════════════════════════════════╝
""")