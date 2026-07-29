"""
Author: Woobin Park
================================================================================
EXCEL COMPARTIVE DIFF CHECKER (100% PRIVATE & LOCAL)
================================================================================
Description:
    This script compares two Excel (.xlsx) workbooks cell-by-cell without 
    uploading data to any servers or training AI models. It captures all text, 
    date, and numerical discrepancies and outputs them into a newly generated, 
    timestamped Excel workbook with perfectly sized columns.

Instructions for Use:
    Step 1: Open this script in any text editor (like Notepad, VS Code, or IDLE).
    Step 2: Update the 'file1' and 'file2' variables below with the full path 
            to your target Excel files. Keep the 'r' before the quotes.
    Step 3: Run the script by double-clicking the file or executing it from 
            the terminal: python "excel diff checker.py"
    Step 4: Check the directory where this script lives. A clean report named 
            "differences_YYYYMMDD_HHMMSS.xlsx" will be waiting for you.
================================================================================
"""

import os
import warnings
from datetime import datetime
import pandas as pd
from openpyxl.utils import get_column_letter # Fixes the version error cleanly

# Paste your targeted file paths here
file1 = r"C:\Users\Unisem\OneDrive\Documents\scripts\Daily Report_06.21.2026.xlsx"
file2 = r"C:\Users\Unisem\OneDrive\Documents\scripts\Daily Report_06.22.2026.xlsx"

try:
    # Hides the data validation warnings in the terminal
    warnings.simplefilter(action='ignore', category=UserWarning) 

    # 1. DYNAMIC FILE NAMING: Creates "differences_20260521_161205.xlsx"
    current_time = datetime.now().strftime("%Y%m%d_%H%M%S")
    filename = f"differences_{current_time}.xlsx"
    
    # 2. LOCAL DIRECTORY FIX: Tells Python to use the folder where this script lives
    script_dir = os.path.dirname(os.path.abspath(__file__))
    output_file = os.path.join(script_dir, filename)

    # Extract dynamic column labels from the input paths
    label1 = os.path.splitext(os.path.basename(file1))[0]
    label2 = os.path.splitext(os.path.basename(file2))[0]

    print("Reading Excel files locally...")
    df1 = pd.read_excel(file1, engine='openpyxl').fillna('')
    df2 = pd.read_excel(file2, engine='openpyxl').fillna('')

    # Compare tables side-by-side using the file names as headers
    comparison = df1.compare(df2, result_names=(label1, label2))
    
    if comparison.empty:
        print("\nSuccess: The files are completely identical!")
    else:
        # Save to local directory with an auto-width adjustment engine
        with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
            comparison.to_excel(writer, sheet_name='Diff_Report')
            worksheet = writer.sheets['Diff_Report']
            
            # Auto-adjust column sizes on generation (using the index tracker)
            for col_idx, col in enumerate(worksheet.columns, start=1):
                max_len = max(len(str(cell.value or '')) for cell in col)
                col_letter = get_column_letter(col_idx) # Native safe conversion
                worksheet.column_dimensions[col_letter].width = max(max_len + 3, 12)

        print(f"\n--- Done! ---")
        print(f"Differences saved to current folder as:")
        print(filename)

except Exception as e:
    print(f"\n!!! ERROR !!!\n{e}")

print("\n-------------------------")
input("Press Enter to close this window...")
