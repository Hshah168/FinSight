"""
Universal Data Loader - Loads ANY financial data format
"""

import pandas as pd
import yaml
import os
from pathlib import Path
from datetime import datetime


class DataLoader:
    """
    Flexible data loader that adapts to different CSV formats
    """
    
    def __init__(self, config_path='config/settings.yaml'):
        """Initialize with configuration"""
        with open(config_path, 'r') as f:
            self.config = yaml.safe_load(f)
        
        self.input_folder = self.config['data']['input_folder']
        self.processed_folder = self.config['data']['processed_folder']
        
        # Create folders if they don't exist
        os.makedirs(self.processed_folder, exist_ok=True)
    
    def load_gl_transactions(self, filename=None):
        """
        Load GL transactions from CSV
        
        Parameters:
        -----------
        filename : str, optional
            Custom filename, or uses default from config
        Returns:
        --------
        DataFrame with GL transactions
        """
        if filename is None:
            filename = self.config['data']['gl_transactions_file']
        filepath = os.path.join(self.input_folder, filename)
        print(f" Loading GL transactions from: {filepath}")
        loaded = False
        try:
            df = pd.read_csv(
                filepath,
                parse_dates=[self.config['data']['date_column']],
                low_memory=False
            )
            loaded = True
        except Exception as e:
            print(f" Error loading file: {e}")
            print(" Trying alternative formats...")
            # Try different encodings
            for encoding in ['utf-8', 'latin1', 'cp1252']:
                try:
                    df = pd.read_csv(filepath, encoding=encoding)
                    print(f"Loaded with {encoding} encoding")
                    loaded = True
                    break
                except:
                    continue
        if loaded:
            print(f"Loaded {len(df):,} transactions")
            print(f"  Columns: {', '.join(df.columns[:5])}...")
            print(f"  Date range: {df[self.config['data']['date_column']].min()} to {df[self.config['data']['date_column']].max()}")
            return df
        else:
            print("Failed to load GL transactions. Returning empty DataFrame.")
            return pd.DataFrame()
    
    def load_budget_data(self, filename=None):
        """Load budget data"""
        if filename is None:
            filename = self.config['data']['budget_file']
        
        filepath = os.path.join(self.input_folder, filename)
        
        print(f" Loading budget data from: {filepath}")
        try:
            df = pd.read_csv(filepath)
            print(f"Loaded {len(df):,} budget records")
            return df
        except Exception as e:
            print(f" Error loading file: {e}")
            print(" Returning empty DataFrame for budget data.")
            return pd.DataFrame()
    
    def detect_columns(self, df):
        """
        Intelligently detect column types
        Useful when column names vary
        """
        column_map = {}
        
        # Look for date columns
        for col in df.columns:
            if 'date' in col.lower():
                column_map['date'] = col
            elif 'amount' in col.lower() or 'value' in col.lower():
                column_map['amount'] = col
            elif 'account' in col.lower() and 'code' in col.lower():
                column_map['account_code'] = col
            elif 'account' in col.lower() and 'name' in col.lower():
                column_map['account_name'] = col
            elif 'dept' in col.lower() or 'department' in col.lower():
                column_map['department'] = col
        
        print(" Detected columns:")
        for standard, actual in column_map.items():
            print(f"  {standard}: {actual}")
        
        return column_map
    
    def save_processed(self, df, filename):
        """Save processed data"""
        filepath = os.path.join(self.processed_folder, filename)
        df.to_csv(filepath, index=False)
        print(f" Saved processed data: {filepath}")


# Test the loader
if __name__ == "__main__":
    loader = DataLoader()
    
    try:
        gl_df = loader.load_gl_transactions()
        budget_df = loader.load_budget_data()
        
        print("\n Data loading successful!")
        print(f"GL Transactions shape: {gl_df.shape}")
        print(f"Budget Data shape: {budget_df.shape}")
        
    except Exception as e:
        print(f"\n Error: {e}")
        print(" Make sure data files are in data/input/ folder")