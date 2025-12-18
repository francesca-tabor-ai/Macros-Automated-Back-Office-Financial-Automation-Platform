#!/usr/bin/env python3
"""
Financial Back-Office Automation Web Application
This web app processes Excel files containing financial transactions,
cleans and validates the data, and generates a comprehensive report.
"""

import pandas as pd
import streamlit as st
from datetime import datetime
import io

# Page configuration
st.set_page_config(
    page_title="Financial Automation",
    page_icon="💰",
    layout="wide"
)

# Custom CSS for better styling
st.markdown("""
    <style>
    .main-header {
        font-size: 2.5rem;
        font-weight: bold;
        color: #1f77b4;
        text-align: center;
        padding: 1rem 0;
    }
    .success-box {
        padding: 1rem;
        border-radius: 0.5rem;
        background-color: #d4edda;
        border: 1px solid #c3e6cb;
        color: #155724;
        margin: 1rem 0;
    }
    </style>
""", unsafe_allow_html=True)

def detect_column_types(df):
    """
    Automatically detects column types in the DataFrame.
    Returns dictionaries with column names by type.
    """
    numeric_cols = []
    date_cols = []
    text_cols = []
    
    for col in df.columns:
        # Skip if already detected as datetime
        if pd.api.types.is_datetime64_any_dtype(df[col]):
            date_cols.append(col)
        # Check if numeric
        elif pd.api.types.is_numeric_dtype(df[col]):
            numeric_cols.append(col)
        else:
            # Try converting a sample to see if it's a date
            sample = df[col].dropna().head(20)  # Check more samples
            if len(sample) > 0:
                # Check if it looks like a date (has date-like patterns)
                sample_str = sample.astype(str)
                date_like = sample_str.str.contains(r'\d{1,4}[-/]\d{1,2}[-/]\d{1,4}', regex=True, na=False)
                
                if date_like.sum() > len(sample) * 0.5:  # More than 50% look like dates
                    try:
                        pd.to_datetime(sample, errors='raise', infer_datetime_format=True)
                        date_cols.append(col)
                    except:
                        text_cols.append(col)
                else:
                    text_cols.append(col)
            else:
                text_cols.append(col)
    
    return {
        'numeric': numeric_cols,
        'date': date_cols,
        'text': text_cols
    }

def find_section_header(df, keywords, start_row=0, max_rows=None, label_col=0):
    """
    Finds a section header row by searching for keywords in the specified column.
    Returns the row index if found, None otherwise.
    """
    if max_rows is None:
        max_rows = len(df)
    
    for idx in range(start_row, min(start_row + max_rows, len(df))):
        cell_value = str(df.iloc[idx, label_col]).lower() if pd.notna(df.iloc[idx, label_col]) else ""
        for keyword in keywords:
            if keyword.lower() in cell_value:
                return idx
    return None

def extract_section_data(df, start_row, end_row=None, label_col=0):
    """
    Extracts data from a section. Assumes labels are in the first column (label_col)
    and data is in subsequent columns.
    Returns a dictionary with label -> data mapping.
    """
    if end_row is None:
        end_row = len(df)
    
    section_data = {}
    first_col = df.columns[label_col]
    
    for idx in range(start_row, min(end_row, len(df))):
        label = df.iloc[idx, label_col]
        
        # Skip if label is empty or NaN
        if pd.isna(label) or str(label).strip() == "":
            continue
        
        label_str = str(label).strip()
        
        # Extract data from this row (skip the label column)
        row_data = []
        for col_idx in range(label_col + 1, len(df.columns)):
            value = df.iloc[idx, col_idx]
            # Try to convert to numeric
            try:
                if pd.notna(value):
                    numeric_value = pd.to_numeric(value, errors='coerce')
                    row_data.append(numeric_value if pd.notna(numeric_value) else value)
                else:
                    row_data.append(None)
            except:
                row_data.append(value)
        
        if row_data:
            section_data[label_str] = row_data
    
    return section_data

def parse_unstructured_sheet(df, sheet_name):
    """
    Parses an unstructured Excel sheet with sections, labels, and scattered data.
    Handles sheets like "Projected Bookings and Revenues" with multiple sections.
    Returns structured data organized by sections.
    """
    sheet_name_lower = sheet_name.lower()
    parsed_data = {
        'sections': {},
        'raw_data': df,
        'structure_type': 'unstructured'
    }
    
    # For "Projected Bookings and Revenues" type sheets
    if 'booking' in sheet_name_lower and 'revenue' in sheet_name_lower:
        # Find section headers - search more broadly
        bookings_row = find_section_header(df, ['bookings'], start_row=0)
        unit_pricing_row = find_section_header(df, ['unit pricing', 'unit price', 'pricing'], start_row=0)
        revenues_one_time_row = find_section_header(df, ['revenues (one time)', 'revenue (one time)', 'revenues one time', 'revenue one time', 'one time'], start_row=0)
        revenues_mrr_row = find_section_header(df, ['revenues (mrr)', 'revenue (mrr)', 'revenues mrr', 'revenue mrr', 'mrr'], start_row=0)
        
        # Extract Bookings section
        if bookings_row is not None:
            end_row = unit_pricing_row if unit_pricing_row else revenues_one_time_row if revenues_one_time_row else len(df)
            bookings_data = extract_section_data(df, bookings_row + 1, end_row)
            parsed_data['sections']['Bookings'] = bookings_data
        
        # Extract Unit Pricing section
        if unit_pricing_row is not None:
            end_row = revenues_one_time_row if revenues_one_time_row else revenues_mrr_row if revenues_mrr_row else len(df)
            unit_pricing_data = extract_section_data(df, unit_pricing_row + 1, end_row)
            parsed_data['sections']['Unit Pricing'] = unit_pricing_data
        
        # Extract Revenues (One Time) section
        if revenues_one_time_row is not None:
            end_row = revenues_mrr_row if revenues_mrr_row else len(df)
            revenues_one_time_data = extract_section_data(df, revenues_one_time_row + 1, end_row)
            parsed_data['sections']['Revenues (One Time)'] = revenues_one_time_data
        
        # Extract Revenues (MRR) section
        if revenues_mrr_row is not None:
            revenues_mrr_data = extract_section_data(df, revenues_mrr_row + 1, len(df))
            parsed_data['sections']['Revenues (MRR)'] = revenues_mrr_data
    
    return parsed_data

def clean_data(df, sheet_name=None):
    """
    Cleans the input DataFrame, handling both structured and unstructured formats.
    For unstructured sheets, parses sections separately.
    
    Returns the cleaned DataFrame and column type information, or parsed section data.
    """
    # Check if this looks like an unstructured sheet
    sheet_name_lower = (sheet_name or "").lower()
    
    # For unstructured financial sheets (like "Projected Bookings and Revenues")
    if sheet_name and ('booking' in sheet_name_lower and 'revenue' in sheet_name_lower):
        # Parse as unstructured sheet
        parsed_data = parse_unstructured_sheet(df, sheet_name)
        return parsed_data, None  # Return parsed data instead of cleaned df
    
    # For structured sheets, use original cleaning logic
    cleaned_df = df.copy()
    
    # Remove completely empty rows (rows where all values are NaN)
    cleaned_df = cleaned_df.dropna(how='all')
    
    # Remove completely empty columns
    cleaned_df = cleaned_df.dropna(axis=1, how='all')
    
    # Detect column types
    col_types = detect_column_types(cleaned_df)
    
    # Convert numeric columns
    for col in col_types['numeric']:
        cleaned_df[col] = pd.to_numeric(cleaned_df[col], errors='coerce')
    
    # Convert date columns
    for col in col_types['date']:
        if col not in col_types['numeric']:
            cleaned_df[col] = pd.to_datetime(cleaned_df[col], errors='coerce', infer_datetime_format=True)
    
    # Fill empty text columns
    for col in col_types['text']:
        if cleaned_df[col].isna().any():
            cleaned_df[col] = cleaned_df[col].fillna("Uncategorized")
    
    # Reset index after cleaning
    cleaned_df = cleaned_df.reset_index(drop=True)
    
    return cleaned_df, col_types

def validate_data(df, col_types):
    """
    Validates the data and flags problematic rows.
    Checks for missing values in important columns and invalid numeric values.
    Returns a DataFrame containing only the rows with issues.
    """
    issues = []
    
    # Get numeric columns for validation
    numeric_cols = col_types.get('numeric', [])
    date_cols = col_types.get('date', [])
    
    # Check each row for issues
    for index, row in df.iterrows():
        issue_reasons = []
        
        # Flag missing or invalid numeric values (<= 0 or NaN)
        for col in numeric_cols:
            if pd.isna(row[col]) or (pd.notna(row[col]) and row[col] <= 0):
                issue_reasons.append(f"{col} is missing or <= 0")
        
        # Flag missing date values
        for col in date_cols:
            if pd.isna(row[col]):
                issue_reasons.append(f"{col} is missing")
        
        # Flag completely empty rows (all values NaN)
        if row.isna().all():
            issue_reasons.append("Row is completely empty")
        
        # If there are any issues, add to issues list
        if issue_reasons:
            issue_row = row.copy()
            issue_row['Issue'] = "; ".join(issue_reasons)
            issues.append(issue_row)
    
    # Create DataFrame from issues list
    if issues:
        issues_df = pd.DataFrame(issues)
    else:
        issues_df = pd.DataFrame(columns=df.columns.tolist() + ['Issue'])
    
    return issues_df

def find_column_by_keywords(df, keywords, case_sensitive=False):
    """
    Finds a column in the DataFrame that contains any of the given keywords.
    Returns the column name if found, None otherwise.
    """
    if not case_sensitive:
        keywords = [k.lower() for k in keywords]
    
    for col in df.columns:
        col_lower = col.lower() if not case_sensitive else col
        for keyword in keywords:
            if keyword in col_lower:
                return col
    return None

def process_unstructured_summary(parsed_data, sheet_name):
    """
    Processes unstructured sheet data and creates financial summaries.
    Works with parsed section data from unstructured Excel sheets.
    """
    summary_rows = []
    sections = parsed_data.get('sections', {})
    
    # Process Bookings section
    if 'Bookings' in sections:
        bookings = sections['Bookings']
        
        # Find "Direct" bookings
        direct_data = None
        for key in bookings.keys():
            if 'direct' in key.lower():
                direct_data = bookings[key]
                break
        
        if direct_data:
            # Sum all numeric values in direct_data
            direct_total = sum([v for v in direct_data if isinstance(v, (int, float)) and pd.notna(v)])
            summary_rows.append({
                'Metric': 'Total Bookings (Direct)',
                'Value': direct_total
            })
    
    # Process Unit Pricing section
    if 'Unit Pricing' in sections:
        unit_pricing = sections['Unit Pricing']
        
        for key, values in unit_pricing.items():
            if values:
                numeric_values = [v for v in values if isinstance(v, (int, float)) and pd.notna(v)]
                if numeric_values:
                    total = sum(numeric_values)
                    avg = total / len(numeric_values) if numeric_values else 0
                    summary_rows.append({
                        'Metric': f'Unit Pricing - {key} (Total)',
                        'Value': total
                    })
                    summary_rows.append({
                        'Metric': f'Unit Pricing - {key} (Average)',
                        'Value': avg
                    })
    
    # Process Revenues (One Time) section
    if 'Revenues (One Time)' in sections:
        revenues_one_time = sections['Revenues (One Time)']
        
        for key, values in revenues_one_time.items():
            if values:
                numeric_values = [v for v in values if isinstance(v, (int, float)) and pd.notna(v)]
                if numeric_values:
                    total = sum(numeric_values)
                    summary_rows.append({
                        'Metric': f'Revenues (One Time) - {key}',
                        'Value': total
                    })
                    
                    # Special handling for "Total Revenue (One Time)" or similar
                    if 'total' in key.lower():
                        summary_rows.append({
                            'Metric': 'Total Revenue (One Time) - Grand Total',
                            'Value': total
                        })
    
    # Process Revenues (MRR) section
    if 'Revenues (MRR)' in sections:
        revenues_mrr = sections['Revenues (MRR)']
        
        for key, values in revenues_mrr.items():
            if values:
                numeric_values = [v for v in values if isinstance(v, (int, float)) and pd.notna(v)]
                if numeric_values:
                    total = sum(numeric_values)
                    summary_rows.append({
                        'Metric': f'Revenues (MRR) - {key}',
                        'Value': total
                    })
    
    if summary_rows:
        return pd.DataFrame(summary_rows)
    return None

def process_sheet_specific_summary(sheet_name, df, col_types, parsed_data=None):
    """
    Processes sheet-specific financial calculations based on sheet name.
    Handles both structured and unstructured sheets.
    Returns a DataFrame with sheet-specific summary or None.
    """
    sheet_name_lower = sheet_name.lower()
    
    # Handle unstructured sheets
    if parsed_data and parsed_data.get('structure_type') == 'unstructured':
        return process_unstructured_summary(parsed_data, sheet_name)
    
    # Handle structured sheets (original logic)
    numeric_cols = col_types.get('numeric', []) if col_types else []
    text_cols = col_types.get('text', []) if col_types else []
    
    # Process "Projected Bookings and Revenues" or similar sheets (structured version)
    if 'booking' in sheet_name_lower and 'revenue' in sheet_name_lower:
        summary_rows = []
        
        # Find relevant columns
        direct_col = find_column_by_keywords(df, ['direct', 'direct cost', 'direct booking'], case_sensitive=False)
        one_time_col = find_column_by_keywords(df, ['one time', 'one-time', 'one time ending', 'one-time ending'], case_sensitive=False)
        unit_pricing_col = find_column_by_keywords(df, ['unit pricing', 'unit price', 'pricing'], case_sensitive=False)
        revenue_one_time_col = find_column_by_keywords(df, ['revenue one time', 'revenues one time', 'revenue (one time)', 'revenues (one time)'], case_sensitive=False)
        
        # Calculate Total Booking Cost (Direct + One Time)
        if direct_col and one_time_col:
            total_booking_cost = (df[direct_col].fillna(0).sum() + df[one_time_col].fillna(0).sum())
            summary_rows.append({
                'Metric': 'Total Booking Cost (Direct + One Time)',
                'Value': total_booking_cost
            })
        elif direct_col:
            total_booking_cost = df[direct_col].fillna(0).sum()
            summary_rows.append({
                'Metric': 'Total Booking Cost (Direct)',
                'Value': total_booking_cost
            })
        elif one_time_col:
            total_booking_cost = df[one_time_col].fillna(0).sum()
            summary_rows.append({
                'Metric': 'Total Booking Cost (One Time)',
                'Value': total_booking_cost
            })
        
        # Unit Pricing summary
        if unit_pricing_col:
            unit_pricing_total = df[unit_pricing_col].fillna(0).sum()
            unit_pricing_avg = df[unit_pricing_col].fillna(0).mean()
            unit_pricing_count = (df[unit_pricing_col].fillna(0) > 0).sum()
            summary_rows.append({
                'Metric': 'Unit Pricing - Total',
                'Value': unit_pricing_total
            })
            summary_rows.append({
                'Metric': 'Unit Pricing - Average',
                'Value': unit_pricing_avg
            })
            summary_rows.append({
                'Metric': 'Unit Pricing - Count (Non-Zero)',
                'Value': unit_pricing_count
            })
        
        # Revenues (One Time) summary
        if revenue_one_time_col:
            revenue_one_time_total = df[revenue_one_time_col].fillna(0).sum()
            summary_rows.append({
                'Metric': 'Revenues (One Time) - Total',
                'Value': revenue_one_time_total
            })
        
        # Add all numeric column summaries
        for col in numeric_cols:
            if col not in [direct_col, one_time_col, unit_pricing_col, revenue_one_time_col]:
                col_total = df[col].fillna(0).sum()
                summary_rows.append({
                    'Metric': f'{col} - Total',
                    'Value': col_total
                })
        
        if summary_rows:
            return pd.DataFrame(summary_rows)
    
    # Process sheets with "Revenue" in name
    elif 'revenue' in sheet_name_lower:
        summary_rows = []
        
        # Find revenue-related columns
        revenue_cols = [col for col in numeric_cols if 'revenue' in col.lower()]
        
        for col in revenue_cols:
            total = df[col].fillna(0).sum()
            avg = df[col].fillna(0).mean()
            summary_rows.append({
                'Metric': f'{col} - Total',
                'Value': total
            })
            summary_rows.append({
                'Metric': f'{col} - Average',
                'Value': avg
            })
        
        if summary_rows:
            return pd.DataFrame(summary_rows)
    
    # Process sheets with "Cost" or "Expense" in name
    elif 'cost' in sheet_name_lower or 'expense' in sheet_name_lower:
        summary_rows = []
        
        cost_cols = [col for col in numeric_cols if any(word in col.lower() for word in ['cost', 'expense', 'spend'])]
        
        for col in cost_cols:
            total = df[col].fillna(0).sum()
            summary_rows.append({
                'Metric': f'{col} - Total',
                'Value': total
            })
        
        if summary_rows:
            return pd.DataFrame(summary_rows)
    
    # Default: Try to find a category-like column and group
    category_col = None
    for col in text_cols:
        unique_count = df[col].nunique()
        if unique_count > 1 and unique_count < len(df) * 0.8:
            category_col = col
            break
    
    if category_col is None and text_cols:
        category_col = text_cols[0]
    
    # If we have a category column and numeric columns, create grouped summary
    if category_col and numeric_cols:
        summaries = []
        for num_col in numeric_cols:
            summary = df.groupby(category_col, as_index=False)[num_col].sum()
            summary = summary.sort_values(num_col, ascending=False)
            
            # Add total row
            total_row = pd.DataFrame({
                category_col: ['TOTAL'],
                num_col: [summary[num_col].sum()]
            })
            summary = pd.concat([summary, total_row], ignore_index=True)
            summaries.append(summary)
        
        # Merge all numeric summaries
        if len(summaries) > 1:
            result = summaries[0]
            for s in summaries[1:]:
                result = result.merge(s, on=category_col, how='outer')
            return result
        else:
            return summaries[0] if summaries else None
    elif numeric_cols:
        # Create basic statistics for numeric columns
        summary_data = {}
        for col in numeric_cols:
            summary_data[f'{col}_Sum'] = [df[col].fillna(0).sum()]
            summary_data[f'{col}_Mean'] = [df[col].fillna(0).mean()]
            summary_data[f'{col}_Min'] = [df[col].fillna(0).min()]
            summary_data[f'{col}_Max'] = [df[col].fillna(0).max()]
        return pd.DataFrame(summary_data)
    else:
        return None

def summarize_data(df, col_types, sheet_name=None, parsed_data=None):
    """
    Creates summaries based on available columns and sheet name.
    Applies sheet-specific logic for financial calculations.
    Handles both structured and unstructured sheets.
    Returns a DataFrame with the summary or None if no suitable columns.
    """
    # If sheet name is provided, use sheet-specific processing
    if sheet_name:
        return process_sheet_specific_summary(sheet_name, df, col_types, parsed_data)
    
    # Fallback to generic processing
    numeric_cols = col_types.get('numeric', [])
    text_cols = col_types.get('text', [])
    
    # Try to find a category-like column
    category_col = None
    for col in text_cols:
        unique_count = df[col].nunique()
        if unique_count > 1 and unique_count < len(df) * 0.8:
            category_col = col
            break
    
    if category_col is None and text_cols:
        category_col = text_cols[0]
    
    if category_col and numeric_cols:
        summaries = []
        for num_col in numeric_cols:
            summary = df.groupby(category_col, as_index=False)[num_col].sum()
            summary = summary.sort_values(num_col, ascending=False)
            
            total_row = pd.DataFrame({
                category_col: ['TOTAL'],
                num_col: [summary[num_col].sum()]
            })
            summary = pd.concat([summary, total_row], ignore_index=True)
            summaries.append(summary)
        
        if len(summaries) > 1:
            result = summaries[0]
            for s in summaries[1:]:
                result = result.merge(s, on=category_col, how='outer')
            return result
        else:
            return summaries[0] if summaries else None
    elif numeric_cols:
        summary_data = {}
        for col in numeric_cols:
            summary_data[f'{col}_Sum'] = [df[col].fillna(0).sum()]
            summary_data[f'{col}_Mean'] = [df[col].fillna(0).mean()]
            summary_data[f'{col}_Min'] = [df[col].fillna(0).min()]
            summary_data[f'{col}_Max'] = [df[col].fillna(0).max()]
        return pd.DataFrame(summary_data)
    else:
        return None

def create_excel_file(sheets_data):
    """
    Creates an Excel file in memory with processed data from all sheets.
    sheets_data: Dictionary with sheet names as keys and dicts containing 
                 'cleaned', 'summary', 'issues' DataFrames as values.
    Handles both structured and unstructured sheets.
    Returns the Excel file as bytes.
    """
    # Create an in-memory buffer
    output = io.BytesIO()
    
    # Create an Excel writer object
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        for sheet_name, data in sheets_data.items():
            parsed_data = data.get('parsed_data')
            
            if parsed_data and parsed_data.get('structure_type') == 'unstructured':
                # Handle unstructured sheets - write sections separately
                sections = parsed_data.get('sections', {})
                
                for section_name, section_data in sections.items():
                    if section_data:
                        # Convert section data to DataFrame
                        max_len = max([len(v) if isinstance(v, list) else 1 for v in section_data.values()])
                        display_data = {}
                        for key, values in section_data.items():
                            if isinstance(values, list):
                                padded_values = values[:max_len] + [None] * (max_len - len(values))
                                display_data[key] = padded_values
                            else:
                                display_data[key] = [values]
                        
                        section_df = pd.DataFrame(display_data)
                        section_sheet_name = f"{sheet_name}_{section_name}"[:31]
                        section_df.to_excel(writer, sheet_name=section_sheet_name, index=False)
                
                # Write summary if available
                if data['summary'] is not None and len(data['summary']) > 0:
                    summary_sheet_name = f"{sheet_name}_Summary"[:31]
                    data['summary'].to_excel(writer, sheet_name=summary_sheet_name, index=False)
            else:
                # Handle structured sheets (original logic)
                cleaned_df = data['cleaned']
                cleaned_sheet_name = f"{sheet_name}_Cleaned"[:31]
                cleaned_df.to_excel(writer, sheet_name=cleaned_sheet_name, index=False)
                
                # Write summary if available
                if data['summary'] is not None and len(data['summary']) > 0:
                    summary_sheet_name = f"{sheet_name}_Summary"[:31]
                    data['summary'].to_excel(writer, sheet_name=summary_sheet_name, index=False)
                
                # Write issues for this sheet
                issues_df = data['issues']
                if len(issues_df) > 0:
                    issues_sheet_name = f"{sheet_name}_Issues"[:31]
                    issues_df.to_excel(writer, sheet_name=issues_sheet_name, index=False)
    
    # Get the Excel file from the buffer
    output.seek(0)
    return output.getvalue()

def main():
    """
    Main function that creates the Streamlit web interface.
    """
    # Header
    st.markdown('<div class="main-header">💰 Financial Back-Office Automation</div>', unsafe_allow_html=True)
    st.markdown("---")
    
    # Sidebar with instructions
    with st.sidebar:
        st.header("📋 Instructions")
        st.markdown("""
        1. **Upload** an Excel file (.xlsx or .xls)
        2. The app works with **any columns**!
        3. It will automatically:
           - Detect column types (numbers, dates, text)
           - Clean the data
           - Validate for issues
           - Generate summaries (if applicable)
        4. **Download** the processed report
        """)
        st.markdown("---")
        st.markdown("**Note:** Works with any Excel file structure. The app automatically detects and processes columns.")
    
    # File uploader
    st.header("📤 Upload Excel File")
    uploaded_file = st.file_uploader(
        "Choose an Excel file",
        type=['xlsx', 'xls'],
        help="Upload your financial transaction Excel file"
    )
    
    # Process the file if uploaded
    if uploaded_file is not None:
        try:
            # Show progress
            with st.spinner("Reading file..."):
                # Read all sheets from the Excel file
                excel_file = pd.ExcelFile(uploaded_file)
                sheet_names = excel_file.sheet_names
                
                # Read all sheets into a dictionary
                all_sheets = {}
                for sheet_name in sheet_names:
                    all_sheets[sheet_name] = pd.read_excel(uploaded_file, sheet_name=sheet_name)
                
                st.success(f"✅ Successfully read {len(sheet_names)} sheet(s): {', '.join(sheet_names)}")
            
            # Process each sheet separately
            sheets_data = {}  # Store processed data for each sheet
            
            # Process each sheet
            for sheet_name, df in all_sheets.items():
                with st.expander(f"📄 Sheet: {sheet_name} ({len(df)} rows)", expanded=(len(sheet_names) == 1)):
                    # Display original data info
                    st.info(f"📊 Columns: {', '.join(df.columns.tolist())}")
                    
                    # Show sheet-specific detected columns
                    sheet_name_lower = sheet_name.lower()
                    detected_info = []
                    
                    if 'booking' in sheet_name_lower and 'revenue' in sheet_name_lower:
                        # Check what columns we can find for bookings & revenues
                        direct_col = find_column_by_keywords(df, ['direct', 'direct cost', 'direct booking'], case_sensitive=False)
                        one_time_col = find_column_by_keywords(df, ['one time', 'one-time', 'one time ending', 'one-time ending'], case_sensitive=False)
                        unit_pricing_col = find_column_by_keywords(df, ['unit pricing', 'unit price', 'pricing'], case_sensitive=False)
                        revenue_one_time_col = find_column_by_keywords(df, ['revenue one time', 'revenues one time', 'revenue (one time)', 'revenues (one time)'], case_sensitive=False)
                        
                        detected_info.append("🔍 **Bookings & Revenues Sheet Detected**")
                        if direct_col:
                            detected_info.append(f"  ✓ Direct Cost: {direct_col}")
                        if one_time_col:
                            detected_info.append(f"  ✓ One Time Ending: {one_time_col}")
                        if unit_pricing_col:
                            detected_info.append(f"  ✓ Unit Pricing: {unit_pricing_col}")
                        if revenue_one_time_col:
                            detected_info.append(f"  ✓ Revenues (One Time): {revenue_one_time_col}")
                    
                    elif 'revenue' in sheet_name_lower:
                        detected_info.append("🔍 **Revenue Sheet Detected** - Will calculate revenue totals")
                    elif 'cost' in sheet_name_lower or 'expense' in sheet_name_lower:
                        detected_info.append("🔍 **Cost/Expense Sheet Detected** - Will calculate cost totals")
                    
                    if detected_info:
                        st.markdown("\n".join(detected_info))
                    
                    # Process the data for this sheet
                    col1, col2, col3 = st.columns(3)
                    
                    cleaned_result = None
                    col_types = None
                    parsed_data = None
                    issues_df = pd.DataFrame()
                    
                    with col1:
                        with st.spinner("Parsing..."):
                            cleaned_result, col_types = clean_data(df, sheet_name=sheet_name)
                            
                            # Check if this is unstructured data
                            if isinstance(cleaned_result, dict) and cleaned_result.get('structure_type') == 'unstructured':
                                parsed_data = cleaned_result
                                sections_found = list(parsed_data.get('sections', {}).keys())
                                st.success(f"✅ Parsed {len(sections_found)} sections")
                                if sections_found:
                                    st.caption(f"Sections: {', '.join(sections_found)}")
                            else:
                                cleaned_df = cleaned_result
                                st.success(f"✅ Cleaned {len(cleaned_df)} rows")
                                if col_types:
                                    if col_types['numeric']:
                                        st.caption(f"📊 Numeric: {', '.join(col_types['numeric'])}")
                                    if col_types['date']:
                                        st.caption(f"📅 Dates: {', '.join(col_types['date'])}")
                                    if col_types['text']:
                                        st.caption(f"📝 Text: {', '.join(col_types['text'])}")
                    
                    with col2:
                        with st.spinner("Validating..."):
                            if parsed_data:
                                # For unstructured sheets, validation is minimal
                                st.info("ℹ️ Unstructured sheet - validation skipped")
                            else:
                                issues_df = validate_data(cleaned_df, col_types)
                                issue_count = len(issues_df)
                                if issue_count > 0:
                                    st.warning(f"⚠️ {issue_count} issues found")
                                else:
                                    st.success("✅ No issues")
                    
                    with col3:
                        with st.spinner("Summarizing..."):
                            if parsed_data:
                                summary_df = summarize_data(None, None, sheet_name=sheet_name, parsed_data=parsed_data)
                            else:
                                summary_df = summarize_data(cleaned_df, col_types, sheet_name=sheet_name)
                            
                            if summary_df is not None and len(summary_df) > 0:
                                st.success(f"✅ Summary generated")
                            else:
                                st.info("ℹ️ No summary")
                    
                    # Store processed data for this sheet
                    if parsed_data:
                        # For unstructured sheets, use the raw data and parsed sections
                        sheets_data[sheet_name] = {
                            'cleaned': parsed_data.get('raw_data', df),
                            'summary': summary_df,
                            'issues': issues_df,
                            'col_types': None,
                            'parsed_data': parsed_data
                        }
                    else:
                        sheets_data[sheet_name] = {
                            'cleaned': cleaned_df,
                            'summary': summary_df,
                            'issues': issues_df,
                            'col_types': col_types,
                            'parsed_data': None
                        }
                    
                    # Display results in tabs for this sheet
                    if parsed_data:
                        # Unstructured sheet display
                        tab1, tab2, tab3 = st.tabs(["📊 Sections Data", "📈 Summary", "📋 Raw Data"])
                        
                        with tab1:
                            st.subheader("Parsed Sections")
                            sections = parsed_data.get('sections', {})
                            
                            for section_name, section_data in sections.items():
                                with st.expander(f"📑 {section_name}"):
                                    # Convert section data to DataFrame for display
                                    if section_data:
                                        # Find the maximum length of data arrays
                                        max_len = max([len(v) if isinstance(v, list) else 1 for v in section_data.values()])
                                        
                                        # Create a DataFrame
                                        display_data = {}
                                        for key, values in section_data.items():
                                            if isinstance(values, list):
                                                # Pad or truncate to max_len
                                                padded_values = values[:max_len] + [None] * (max_len - len(values))
                                                display_data[key] = padded_values
                                            else:
                                                display_data[key] = [values]
                                        
                                        section_df = pd.DataFrame(display_data)
                                        st.dataframe(section_df, use_container_width=True)
                                        
                                        # Show totals for numeric data
                                        st.caption(f"Section: {section_name}")
                        
                        with tab2:
                            if summary_df is not None and len(summary_df) > 0:
                                st.dataframe(summary_df, use_container_width=True)
                                
                                # Display key metrics
                                if 'Metric' in summary_df.columns and 'Value' in summary_df.columns:
                                    st.subheader("Key Financial Metrics")
                                    num_metrics = len(summary_df)
                                    num_cols = min(num_metrics, 3)
                                    metric_cols = st.columns(num_cols)
                                    
                                    for idx, row in summary_df.iterrows():
                                        col_idx = idx % num_cols
                                        with metric_cols[col_idx]:
                                            value = row['Value']
                                            if pd.notna(value):
                                                if isinstance(value, (int, float)):
                                                    display_value = f"${value:,.2f}" if abs(value) >= 0.01 else f"${value:.4f}"
                                                else:
                                                    display_value = str(value)
                                            else:
                                                display_value = "N/A"
                                            
                                            st.metric(
                                                row['Metric'],
                                                display_value
                                            )
                            else:
                                st.info("No summary available.")
                        
                        with tab3:
                            st.subheader("Raw Excel Data")
                            st.dataframe(df, use_container_width=True)
                            st.caption("This is the raw data from the Excel sheet. Sections have been parsed above.")
                    else:
                        # Structured sheet display (original)
                        tab1, tab2, tab3 = st.tabs(["📊 Cleaned Data", "📈 Summary", "⚠️ Issues"])
                        
                        with tab1:
                            st.dataframe(cleaned_df, use_container_width=True)
                            st.info(f"Total rows: {len(cleaned_df)}")
                        
                        with tab2:
                            if summary_df is not None and len(summary_df) > 0:
                                st.dataframe(summary_df, use_container_width=True)
                                
                                # Display key metrics based on sheet type
                                if 'Metric' in summary_df.columns and 'Value' in summary_df.columns:
                                    # Sheet-specific summary format
                                    st.subheader("Key Financial Metrics")
                                    num_metrics = len(summary_df)
                                    num_cols = min(num_metrics, 3)
                                    metric_cols = st.columns(num_cols)
                                    
                                    for idx, row in summary_df.iterrows():
                                        col_idx = idx % num_cols
                                        with metric_cols[col_idx]:
                                            value = row['Value']
                                            if pd.notna(value):
                                                # Format based on value type
                                                if isinstance(value, (int, float)):
                                                    display_value = f"{value:,.2f}" if abs(value) >= 0.01 else f"{value:.4f}"
                                                else:
                                                    display_value = str(value)
                                            else:
                                                display_value = "N/A"
                                            
                                            st.metric(
                                                row['Metric'],
                                                display_value
                                            )
                                else:
                                    # Standard grouped summary format
                                    numeric_cols = col_types.get('numeric', [])
                                    if numeric_cols:
                                        st.subheader("Column Totals")
                                        cols = st.columns(min(len(numeric_cols), 4))
                                        for idx, num_col in enumerate(numeric_cols[:4]):
                                            with cols[idx]:
                                                total = cleaned_df[num_col].fillna(0).sum()
                                                st.metric(f"Total {num_col}", f"{total:,.2f}")
                            else:
                                st.info("No summary available. Summary requires numeric columns and/or categorical columns.")
                        
                        with tab3:
                            if len(issues_df) > 0:
                                st.dataframe(issues_df, use_container_width=True)
                                st.warning(f"⚠️ {len(issues_df)} rows need attention")
                            else:
                                st.success("✅ No issues found! All data is valid.")
            
            st.markdown("---")
            
            # Download section
            st.header("💾 Download Report")
            
            # Generate Excel file with all sheets
            with st.spinner("Generating Excel file with all processed sheets..."):
                excel_data = create_excel_file(sheets_data)
                
                # Generate filename with timestamp
                timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                filename = f"financial_report_{timestamp}.xlsx"
                
                # Download button
                st.download_button(
                    label="📥 Download Complete Excel Report",
                    data=excel_data,
                    file_name=filename,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    help="Download the complete report with all processed sheets"
                )
                
                st.markdown('<div class="success-box">✅ Report generated successfully!</div>', unsafe_allow_html=True)
                st.info(f"📄 File will be saved as: {filename}")
                
                # Show what's in the file
                sheet_info = []
                for sheet_name in sheet_names:
                    sheet_info.append(f"**{sheet_name}**: Cleaned data, Summary, and Issues sheets")
                
                st.markdown(f"""
                **The Excel file contains processed data for each sheet:**
                - {chr(10).join(f'- {info}' for info in sheet_info)}
                
                Each original sheet has been processed separately with its own cleaned data, summary, and issues.
                """)
        
        except Exception as e:
            st.error(f"❌ Error processing file: {str(e)}")
            st.info("Please make sure the file is a valid Excel file with the required columns.")
    
    else:
        # Show placeholder when no file is uploaded
        st.info("👆 Please upload an Excel file to get started")
        
        # Show example of expected format
        with st.expander("📝 How It Works"):
            st.markdown("""
            **This app works with any Excel file structure!**
            
            The app automatically:
            - **Detects** column types (numbers, dates, text)
            - **Cleans** numeric and date columns
            - **Validates** data quality
            - **Summarizes** data when possible (groups by categories if available)
            
            **Example file formats that work:**
            - Financial transactions (Date, Amount, Category, etc.)
            - Sales data (Product, Quantity, Price, etc.)
            - Inventory (Item, Stock, Cost, etc.)
            - Any structured data with columns!
            """)

if __name__ == "__main__":
    main()

