#!/usr/bin/env python3
"""
Financial Back-Office Automation Script
This script processes Excel files containing financial transactions,
cleans and validates the data, and generates a comprehensive report.
"""

import pandas as pd
import tkinter as tk
from tkinter import filedialog
from datetime import datetime
import os

def select_excel_file():
    """
    Opens a file picker dialog to select an Excel file.
    Returns the file path if selected, None if cancelled.
    """
    # Create a root window (hidden) for the file dialog
    root = tk.Tk()
    root.withdraw()  # Hide the main window
    
    # Open file picker dialog
    file_path = filedialog.askopenfilename(
        title="Select Excel File",
        filetypes=[
            ("Excel files", "*.xlsx *.xls"),
            ("All files", "*.*")
        ]
    )
    
    root.destroy()  # Close the root window
    return file_path

def clean_data(df):
    """
    Cleans the input DataFrame by:
    - Removing empty rows
    - Converting 'Amount' to numeric
    - Converting 'Date' to datetime format
    - Filling empty 'Category' cells with "Uncategorized"
    
    Returns the cleaned DataFrame.
    """
    print("Cleaning data...")
    
    # Create a copy to avoid modifying the original
    cleaned_df = df.copy()
    
    # Remove completely empty rows (rows where all values are NaN)
    cleaned_df = cleaned_df.dropna(how='all')
    
    # Convert 'Amount' column to numeric, replacing non-numeric values with NaN
    cleaned_df['Amount'] = pd.to_numeric(cleaned_df['Amount'], errors='coerce')
    
    # Convert 'Date' column to datetime format
    # Try multiple date formats to handle different input formats
    cleaned_df['Date'] = pd.to_datetime(cleaned_df['Date'], errors='coerce', infer_datetime_format=True)
    
    # Fill empty 'Category' cells with "Uncategorized"
    cleaned_df['Category'] = cleaned_df['Category'].fillna("Uncategorized")
    
    # Reset index after cleaning
    cleaned_df = cleaned_df.reset_index(drop=True)
    
    print(f"  - Removed empty rows")
    print(f"  - Converted {len(cleaned_df)} rows")
    
    return cleaned_df

def validate_data(df):
    """
    Validates the data and flags problematic rows.
    Returns a DataFrame containing only the rows with issues.
    """
    print("Validating data...")
    
    issues = []
    
    # Check each row for issues
    for index, row in df.iterrows():
        issue_reasons = []
        
        # Flag if Amount is missing or <= 0
        if pd.isna(row['Amount']) or row['Amount'] <= 0:
            issue_reasons.append("Amount is missing or <= 0")
        
        # Flag if Date is missing
        if pd.isna(row['Date']):
            issue_reasons.append("Date is missing")
        
        # If there are any issues, add to issues list
        if issue_reasons:
            issue_row = row.copy()
            issue_row['Issue'] = "; ".join(issue_reasons)
            issues.append(issue_row)
    
    # Create DataFrame from issues list
    if issues:
        issues_df = pd.DataFrame(issues)
        print(f"  - Found {len(issues_df)} rows with issues")
    else:
        issues_df = pd.DataFrame(columns=df.columns.tolist() + ['Issue'])
        print("  - No issues found!")
    
    return issues_df

def summarize_data(df):
    """
    Groups the data by 'Category' and sums the 'Amount' for each category.
    Returns a DataFrame with the summary.
    """
    print("Summarizing data...")
    
    # Group by Category and sum the Amount
    summary = df.groupby('Category', as_index=False)['Amount'].sum()
    
    # Sort by Amount in descending order
    summary = summary.sort_values('Amount', ascending=False)
    
    # Add a total row
    total_row = pd.DataFrame({
        'Category': ['TOTAL'],
        'Amount': [summary['Amount'].sum()]
    })
    summary = pd.concat([summary, total_row], ignore_index=True)
    
    print(f"  - Found {len(summary) - 1} categories")
    
    return summary

def export_to_excel(cleaned_df, summary_df, issues_df, output_path):
    """
    Exports the cleaned data, summary, and issues to an Excel file
    with three separate sheets.
    """
    print("Exporting to Excel...")
    
    # Create an Excel writer object
    with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
        # Write each DataFrame to a separate sheet
        cleaned_df.to_excel(writer, sheet_name='Cleaned_Data', index=False)
        summary_df.to_excel(writer, sheet_name='Summary', index=False)
        issues_df.to_excel(writer, sheet_name='Issues', index=False)
    
    print(f"  - File saved to: {output_path}")

def main():
    """
    Main function that orchestrates the entire process.
    """
    print("=" * 60)
    print("Financial Back-Office Automation")
    print("=" * 60)
    print()
    
    # Step 1: Select Excel file
    print("Step 1: Selecting Excel file...")
    file_path = select_excel_file()
    
    if not file_path:
        print("No file selected. Exiting.")
        return
    
    print(f"Selected file: {os.path.basename(file_path)}")
    print()
    
    # Step 2: Read the Excel file
    print("Step 2: Reading Excel file...")
    try:
        df = pd.read_excel(file_path)
        print(f"  - Read {len(df)} rows from file")
        
        # Check if required columns exist
        required_columns = ['Date', 'Description', 'Amount', 'Category']
        missing_columns = [col for col in required_columns if col not in df.columns]
        
        if missing_columns:
            print(f"ERROR: Missing required columns: {', '.join(missing_columns)}")
            print(f"Available columns: {', '.join(df.columns.tolist())}")
            return
        
        print(f"  - Found columns: {', '.join(df.columns.tolist())}")
    except Exception as e:
        print(f"ERROR: Could not read file: {e}")
        return
    
    print()
    
    # Step 3: Clean the data
    print("Step 3: Cleaning data...")
    cleaned_df = clean_data(df)
    print()
    
    # Step 4: Validate the data
    print("Step 4: Validating data...")
    issues_df = validate_data(cleaned_df)
    print()
    
    # Step 5: Summarize the data
    print("Step 5: Summarizing data...")
    summary_df = summarize_data(cleaned_df)
    print()
    
    # Step 6: Generate output filename with timestamp
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    output_filename = f"financial_report_{timestamp}.xlsx"
    output_path = os.path.join(os.path.dirname(file_path), output_filename)
    
    # Step 7: Export to Excel
    print("Step 6: Exporting report...")
    try:
        export_to_excel(cleaned_df, summary_df, issues_df, output_path)
        print()
        print("=" * 60)
        print("Report generated successfully!")
        print(f"Output file: {output_filename}")
        print("=" * 60)
    except Exception as e:
        print(f"ERROR: Could not export file: {e}")
        return

if __name__ == "__main__":
    main()

