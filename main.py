"""
FinSight Main Entry Point
=========================

Single command center for FinSight.

When you run:
    python main.py

Steps performed:
1. Detect new financial data
2. Process GL and budget data
3. Generate financial statements
4. Calculate KPIs and variances
5. Export executive Excel report
"""

import sys
import os

# Ensure project root is in Python path for imports
PROJECT_ROOT = os.path.dirname(os.path.abspath(__file__))
sys.path.append(PROJECT_ROOT)

# Import FinSight Controller
from src.controller.finsight_controller import FinSightController


def main():

    print("\n" + "="*60)
    print("FinSight Financial Intelligence Platform")
    print("="*60)

    controller = FinSightController()

    controller.run()

    print("\nFinSight execution completed successfully.")
    print("="*60)


# Entry point
if __name__ == "__main__":
    main()