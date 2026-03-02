import pandas as pd
import os


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
            # Try alternative paths or formats here if needed
            # If still not found, handle gracefully:
            return pd.DataFrame()  # or raise a clear error

    def load_budget_data(self):

        print("Loading Budget Data...")
        df = pd.read_csv(self.budget_path)

        print(f"Loaded {len(df)} Budget records")
        return df

    def load_chart_of_accounts(self):

        print("Loading Chart of Accounts...")
        df = pd.read_csv(self.coa_path)

        print(f"Loaded {len(df)} COA records")
        return df

    def load_department_mapping(self):

        print("Loading Department Mapping...")
        df = pd.read_csv(self.dept_path)

        print(f"Loaded {len(df)} Department records")
        return df

    def load_all(self):

        gl = self.load_gl_transactions()
        budget = self.load_budget_data()
        coa = self.load_chart_of_accounts()
        dept = self.load_department_mapping()

        return {
            "gl": gl,
            "budget": budget,
            "coa": coa,
            "dept": dept
        }