"""
Variance Analysis Engine
========================
Compares Actual vs Budget and explains differences

Business Logic:
- Calculates dollar and percentage variances
- Flags significant variances based on thresholds
- Categorizes as favorable or unfavorable
- Provides drill-down by account, department, period
"""

import pandas as pd
import numpy as np
from datetime import datetime


class VarianceAnalyzer:
    """
    Budget vs Actual variance analysis
    -------------
    1. Takes actual GL data and budget data
    2. Compares them side-by-side
    3. Calculates variances (difference)
    4. Identifies which variances need attention
    5. Provides explanations and context
    """
    
    def __init__(self, actual_df, budget_df, config_path='config/settings.yaml'):
        """
        Initialize variance analyzer
        ------------------
        1. Store actual and budget data
        2. Load variance thresholds from config
        3. Prepare data for comparison
        
        Parameters:
        -----------
        actual_df : DataFrame
            Actual GL transactions (from data_loader)
        budget_df : DataFrame
            Budget data with planned amounts
        config_path : str
            Path to settings.yaml
        """
        self.actual_df = actual_df.copy()
        self.budget_df = budget_df.copy()

        # Standardize column names
        self.actual_df.columns = self.actual_df.columns.str.lower().str.strip()
        self.budget_df.columns = self.budget_df.columns.str.lower().str.strip()

        # Ensure budget column is consistent
        if 'budgetamount' in self.budget_df.columns:
            self.budget_df.rename(columns={'budgetamount': 'budget_amount'}, inplace=True)

        # Define thresholds
        self.thresholds = {
            'revenue': 10,
            'expenses': 15,
            'overall': 20
        }
    
    def calculate_variances(self, YearMonth='monthly'):
        """
        Calculate Budget vs Actual variances
        
        Step-by-Step Process:
        ---------------------
        1. Group actual data by YearMonth and account
        2. Group budget data by YearMonth and account
        3. Merge them together (actual + budget)
        4. Calculate variance ($)
        5. Calculate variance (%)
        6. Determine favorable vs unfavorable
        
        Parameters:
        -----------
        YearMonth : str
            'monthly', 'quarterly', or 'annual'
            
        Returns:
        --------
        DataFrame with columns:
        - YearMonth, account_type, account_name
        - budget_amount, actual_amount
        - variance_dollar, variance_percent
        - variance_type (Favorable/Unfavorable)
        
        Example Output:
        ---------------
        Period   Account       Budget   Actual   Var($)   Var(%)  Type
        2024-01  Revenue       100,000   85,000  -15,000   -15%   Unfavorable
        2024-01  Marketing      50,000   65,000  +15,000   +30%   Unfavorable
        2024-01  Salaries       80,000   80,000        0     0%   On Target
        """
        print(f"\nCalculating {YearMonth} variances...")
        
        # STEP 1: Prepare actual data
        actual_df = self.actual_df.copy()
        date_col = 'date'
        actual_df[date_col] = pd.to_datetime(actual_df[date_col])
        
        # Add YearMonth column
        if YearMonth == 'monthly':
            actual_df['YearMonth'] = actual_df[date_col].dt.to_period('M')
        elif YearMonth == 'quarterly':
            actual_df['YearMonth'] = actual_df[date_col].dt.to_period('Q')
        elif YearMonth == 'annual':
            actual_df['YearMonth'] = actual_df[date_col].dt.to_period('Y')
        
        # Sum actual by YearMonth and account
        actual_summary = actual_df.groupby([
            'YearMonth',
            'account_type',
            'account_name'
        ])['amount'].sum().reset_index()
        actual_summary.rename(columns={'amount': 'actual_amount'}, inplace=True)
        
        # STEP 2: Prepare budget data
        budget_df = self.budget_df.copy()
        
        #STANDARDIZE COLUMN NAMES ONCE
        budget_df = self.budget_df.copy()
        budget_df.columns = budget_df.columns.str.strip().str.lower()

        # Convert old naming to standard naming
        if 'budgetamount' in budget_df.columns:
            budget_df.rename(columns={'budgetamount': 'budget_amount'}, inplace=True)

        if 'YearMonth' in budget_df.columns:
            budget_df.rename(columns={'YearMonth': 'YearMonth'}, inplace=True)

        # Update self.budget_df so rest of function uses clean schema
        self.budget_df = budget_df
        # Ensure budget has YearMonth column
        if 'YearMonth' not in budget_df.columns:

            # If budget has a date column, convert it
            if 'date' in budget_df.columns or 'YearMonth' in budget_df.columns:
                date_col_budget = 'date' if 'date' in budget_df.columns else 'YearMonth'
                budget_df[date_col_budget] = pd.to_datetime(budget_df[date_col_budget])
                
                if YearMonth == 'monthly':
                    budget_df['YearMonth'] = budget_df[date_col_budget].dt.to_period('M')
                elif YearMonth == 'quarterly':
                    budget_df['YearMonth'] = budget_df[date_col_budget].dt.to_period('Q')
                elif YearMonth == 'annual':
                    budget_df['YearMonth'] = budget_df[date_col_budget].dt.to_period('Y')
        
        # Sum budget by YearMonth and account
        budget_column = None

        for col in ['amount', 'budget_amount', 'budget', 'value']:
            if col in budget_df.columns:
                budget_column = col
                break

        if budget_column is None:
            raise Exception(f"No valid budget column found. Available columns: {budget_df.columns}")

        # Determine groupby columns - if budget has account_type, include it
        budget_groupby_cols = ['YearMonth', 'account_name']
        if 'account_type' in budget_df.columns:
            budget_groupby_cols = ['YearMonth', 'account_type', 'account_name']
        
        budget_summary = budget_df.groupby(budget_groupby_cols)[budget_column].sum().reset_index()

        # Standardize name for downstream processing
        budget_summary.rename(columns={budget_column: 'budgetamount'}, inplace=True)
        
        # Ensure YearMonth columns have matching data types
        # Convert both to period type to match
        if not hasattr(budget_summary['YearMonth'].dtype, 'name') or budget_summary['YearMonth'].dtype.name != 'period[M]':
            budget_summary['YearMonth'] = pd.to_datetime(budget_summary['YearMonth']).dt.to_period('M')
        
        if not hasattr(actual_summary['YearMonth'].dtype, 'name') or actual_summary['YearMonth'].dtype.name != 'period[M]':
            actual_summary['YearMonth'] = pd.to_datetime(actual_summary['YearMonth']).dt.to_period('M')
        
        print(f"  Actual periods: {actual_summary['YearMonth'].nunique()}")
        print(f"  Budget periods: {budget_summary['YearMonth'].nunique()}")
        
        # STEP 3: Merge actual and budget
        # Determine merge keys based on available columns
        merge_keys = ['YearMonth', 'account_type', 'account_name']
        if 'account_type' not in budget_summary.columns:
            merge_keys = ['YearMonth', 'account_name']
        
        variance_df = pd.merge(
            actual_summary,
            budget_summary,
            on=merge_keys,
            how='outer',  # Keep all records even if budget or actual is missing
            indicator=True  # Shows which records matched
        )
        
        # Fill missing values with 0
        variance_df['actual_amount'].fillna(0, inplace=True)
       # Normalize columns
        budget_df.columns = budget_df.columns.str.strip().str.lower()

        if 'budgetamount' in budget_df.columns:
            budget_df.rename(columns={'budgetamount': 'budget_amount'}, inplace=True)
        
        # STEP 4: Calculate variance ($)
        variance_df['variance_dollar'] = variance_df['actual_amount'] - variance_df['budgetamount']
        
        # STEP 5: Calculate variance (%)
        # Avoid division by zero
        variance_df['variance_percent'] = np.where(
            variance_df['budgetamount'] != 0,
            (variance_df['variance_dollar'] / variance_df['budgetamount']) * 100,
            0
        )
        
        # STEP 6: Determine favorable vs unfavorable
        variance_df['variance_type'] = variance_df.apply(
            self._categorize_variance, axis=1
        )
        
        # Sort by absolute variance (biggest variances first)
        variance_df['abs_variance'] = variance_df['variance_dollar'].abs()
        variance_df.sort_values('abs_variance', ascending=False, inplace=True)
        variance_df.drop('abs_variance', axis=1, inplace=True)
        
        print(f"Calculated variances for {len(variance_df)} account-period combinations")
        
        return variance_df
    
    def _categorize_variance(self, row):
        """
        Categorize variance as Favorable or Unfavorable
        
        The Logic:
        ----------
        Revenue/Income:
          - Actual > Budget = FAVORABLE
          - Actual < Budget = UNFAVORABLE
          - Actual = Budget = ON TARGET
        
        Expenses/Costs:
          - Actual > Budget = UNFAVORABLE
          - Actual < Budget = FAVORABLE
          - Actual = Budget = ON TARGET
        """
        variance = row['variance_dollar']
        accounttype = row['account_type']
        
        # If variance is zero, it's on target
        if variance == 0:
            return 'On Target'
        
        # Revenue/Income: positive variance is good
        if accounttype == 'Revenue':
            if variance > 0:
                return 'Favorable'
            else:
                return 'Unfavorable'
        
        # Expenses: negative variance is good
        elif accounttype in ['COGS', 'OpEx']:
            if variance > 0:
                return 'Unfavorable'
            else:
                return 'Favorable'
        
        return 'Neutral'
    
    def flag_significant_variances(self, variance_df):
        """
        Flag variances that need attention
        -----------------------------
        Uses thresholds
        - Revenue variance > 10% → FLAG
        - Expense variance > 15% → FLAG
        - Overall variance > 20% → FLAG

        
        Parameters:
        -----------
        variance_df : DataFrame
            Output from calculate_variances()
            
        Returns:
        --------
        DataFrame with additional column:
        - requires_attention (True/False)
        - attention_reason (text explanation)
        """
        
        df = variance_df.copy()
        
        # Initialize flag column
        df['requires_attention'] = False
        df['attention_reason'] = ''
        
        # Get thresholds
        revenue_threshold = self.thresholds['revenue']
        expense_threshold = self.thresholds['expenses']
        overall_threshold = self.thresholds['overall']
        
        # Flag revenue variances
        revenue_mask = (
            (df['account_type'] == 'Revenue') & 
            (df['variance_percent'].abs() > revenue_threshold)
        )
        df.loc[revenue_mask, 'requires_attention'] = True
        df.loc[revenue_mask, 'attention_reason'] = f'Revenue variance >{revenue_threshold}%'
        
        # Flag expense variances
        expense_mask = (
            (df['account_type'].isin(['COGS', 'OpEx'])) & 
            (df['variance_percent'].abs() > expense_threshold)
        )
        df.loc[expense_mask, 'requires_attention'] = True
        df.loc[expense_mask, 'attention_reason'] = f'Expense variance >{expense_threshold}%'
        
        # Count flagged items
        flagged_count = df['requires_attention'].sum()
        
        print(f"\n Flagged {flagged_count} significant variances:")
        print(f"   Revenue threshold: >{revenue_threshold}%")
        print(f"   Expense threshold: >{expense_threshold}%")
        
        return df
    
    def get_top_variances(self, variance_df, n=10, variance_type=None):
        """
        Get the top N biggest variances
        
        Parameters:
        -----------
        variance_df : DataFrame
            Variance data from calculate_variances()
        n : int
            Number of top variances to return
        variance_type : str, optional
            Filter by type: 'Favorable', 'Unfavorable', or None for all
            
        Returns:
        --------
        DataFrame with top N variances sorted by absolute value
        """
        
        df = variance_df.copy()
        
        # Filter by type if specified
        if variance_type:
            df = df[df['variance_type'] == variance_type]
        
        # Sort by absolute variance
        df['abs_variance'] = df['variance_dollar'].abs()
        df_sorted = df.sort_values('abs_variance', ascending=False).head(n)
        df_sorted.drop('abs_variance', axis=1, inplace=True)
        
        return df_sorted
    
    def summarize_by_category(self, variance_df):
        """
        Roll up variances by category
        
        Returns:
        --------
        DataFrame with summary by accounttype and category
        """
        
        summary = variance_df.groupby(['account_type', 'account_name']).agg({
            'budget_amount': 'sum',
            'actual_amount': 'sum',
            'variance_dollar': 'sum'
        }).reset_index()
        
        # Recalculate variance %
        summary['variance_percent'] = np.where(
            summary['budget_amount'] != 0,
            (summary['variance_dollar'] / summary['budget_amount']) * 100,
            0
        )
        
        # Categorize
        summary['variance_type'] = summary.apply(
            self._categorize_variance, axis=1
        )
        
        return summary
    
    def export_variance_report(self, variance_df, filename):
        """
        Export variance analysis to Excel
        
        Creates professional variance report with:
        - Summary sheet
        - Detail sheet
        - Flagged items sheet
        
        Parameters:
        -----------
        variance_df : DataFrame
            Variance data
        filename : str
            Output Excel filename
        """
        
        with pd.ExcelWriter(filename, engine='openpyxl') as writer:
            
            # Sheet 1: Summary by category
            summary = self.summarize_by_category(variance_df)
            summary.to_excel(writer, sheet_name='Summary', index=False)
            
            # Sheet 2: All variances
            variance_df.to_excel(writer, sheet_name='Detail', index=False)
            
            # Sheet 3: Flagged items only
            if 'requires_attention' in variance_df.columns:
                flagged = variance_df[variance_df['requires_attention'] == True]
                flagged.to_excel(writer, sheet_name='Requires Attention', index=False)
        
        print(f"Variance report exported to: {filename}")

if __name__ == "__main__":
    """
    Test the VarianceAnalyzer module
    """
    import sys
    sys.path.append('.')
    from src.core.data_loader import DataLoader
    
    print("="*60)
    print("TESTING VARIANCE ANALYSIS MODULE")
    print("="*60)
    
    # Load data
    loader = DataLoader()
    actual_df = loader.load_gl_transactions()
    budget_df = loader.load_budget_data()
    
    # Create analyzer
    va = VarianceAnalyzer(actual_df, budget_df)
    
    # Calculate variances
    print("\nCalculating monthly variances...")
    variances = va.calculate_variances(period='monthly')
    
    print("\nTop 10 variances:")
    print(variances.head(10))
    
    # Flag significant ones
    print("\nFlagging significant variances...")
    variances_flagged = va.flag_significant_variances(variances)
    
    # Show flagged items
    flagged = variances_flagged[variances_flagged['requires_attention'] == True]
    print(f"\nItems requiring attention:")
    print(flagged[['period', 'account_name', 'variance_dollar', 'variance_percent', 'attention_reason']].head())
    
    # Export
    print("\nExporting variance report...")
    va.export_variance_report(variances_flagged, 'outputs/reports/test_variance_report.xlsx')
    
    print("\n" + "="*60)
    print("TEST COMPLETE!")
    print("="*60)
