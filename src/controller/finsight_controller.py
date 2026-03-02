"""
FinSight Automation Controller
"""

import os
import pandas as pd
from datetime import datetime


# Import modules
from src.ingestion.data_loader import DataLoader
from src.analytics.financial_statements import FinancialStatements
from src.analytics.variance_analysis import VarianceAnalyzer
from src.reporting.excel_reporter import ExcelReporter


class FinSightController:

    def __init__(self):

        self.raw_data_path = "data/powerbi"
        self.processed_path = "data/processed"
        self.output_path = "outputs/reports"

        os.makedirs(self.output_path, exist_ok=True)

        print("FinSight Controller Initialized")

    # Step 1: Detect Data
    def detect_data(self):

        print("\nDetecting financial data...")

        if not os.path.exists(self.raw_data_path):
            os.makedirs(self.raw_data_path)
        print("Created data/raw folder")

        files = os.listdir(self.raw_data_path)

        required_files = [
            "gl_transactions.csv",
            "budget_data.csv",
            "chart_of_accounts.csv",
            "department_mapping.csv"
        ]

        for file in required_files:
            if file not in files:
                raise Exception(f"Missing required file: {file}")

        print("All required data files detected.")

    # Step 2: Load Data
    def load_data(self):

        print("\nLoading financial data...")
        from src.ingestion.data_loader import DataLoader
        print("Import successful")

        loader = DataLoader()
        print("Loader created")

        data = loader.load_all()
        print("load_all executed")

        self.gl_df = data["gl"]
        self.budget_df = data["budget"]
        self.coa_df = data["coa"]
        self.dept_df = data["dept"]

        print("All data loaded successfully")

    # Step 3: Generate Financial Statements
    def generate_financials(self):

        print("\nGenerating financial statements...")

        fs = FinancialStatements(self.gl_df)

        self.pnl = fs.generate_pnl(period="monthly")

        self.metrics = fs.calculate_metrics(self.pnl)

        print("Financial statements generated.")

    # Step 4: Generate Variances
    def generate_variances(self):

        print("\nGenerating variance analysis...")

        va = VarianceAnalyzer(self.gl_df, self.budget_df)

        self.variances = va.calculate_variances(YearMonth="monthly")

        print("Variance analysis complete.")

    # Step 5: Export Reports
    def generate_report(self):

        print("\nExporting reports...")

        timestamp = datetime.now().strftime("%Y%m%d_%H%M")

        filename = f"{self.output_path}/FinSight_Report_{timestamp}.xlsx"

        reporter = ExcelReporter()

        reporter.create_monthly_package(
            self.pnl,
            self.variances,
            self.metrics,
            filename
        )

        print(f"Report generated: {filename}")

    # Step 6: Export Power BI Datasets
    def export_powerbi_data(self):

        print("\nExporting Power BI datasets...")

        powerbi_path = os.path.join(os.getcwd(), "outputs", "powerbi")
        os.makedirs(powerbi_path, exist_ok=True)

        pnl_path = os.path.join(powerbi_path, "pnl.csv")
        var_path = os.path.join(powerbi_path, "variances.csv")
        metrics_path = os.path.join(powerbi_path, "metrics.csv")

        self.pnl.to_csv(pnl_path, index=False)
        self.variances.to_csv(var_path, index=False)
        self.metrics.to_csv(metrics_path, index=False)

        print("pnl.csv saved")
        print("variances.csv saved")
        print("metrics.csv saved")

    # Master Execution Function
    def run(self):

        print("\n" + "="*50)
        print("RUNNING FINSIGHT AUTOMATION PIPELINE")
        print("="*50)

        self.detect_data()

        self.load_data()

        self.generate_financials()

        self.generate_variances()

        self.generate_report()

        self.export_powerbi_data()

        print("\n FinSight Pipeline Complete")