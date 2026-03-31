"""
Data Cleaning Script for LMS Dataset.
Handles: duplicates, missing values, date parsing, text standardization,
and derived column creation.
"""

import pandas as pd
import numpy as np
from datetime import datetime
import os
import warnings
warnings.filterwarnings('ignore')


def load_raw_data(filepath):
    """Load the raw CSV dataset."""
    print("[1/6] Loading raw data...")
    df = pd.read_csv(filepath)
    print(f"  ✓ Loaded {len(df)} rows, {len(df.columns)} columns")
    print(f"  ✓ Columns: {list(df.columns)}")
    return df


def remove_duplicates(df):
    """Remove exact duplicate rows."""
    print("\n[2/6] Removing duplicates...")
    initial = len(df)
    df = df.drop_duplicates()
    removed = initial - len(df)
    print(f"  ✓ Removed {removed} duplicate rows ({removed/initial*100:.1f}%)")
    print(f"  ✓ Remaining: {len(df)} rows")
    return df


def standardize_text(df):
    """Standardize text columns for consistency."""
    print("\n[3/6] Standardizing text values...")

    # --- Status column ---
    # Map all variations to standard values
    status_mapping = {
        "completed": "Completed",
        "COMPLETED": "Completed",
        "Completd": "Completed",
        "Completed": "Completed",
        "pending": "Pending",
        "PENDING": "Pending",
        "Pendig": "Pending",
        "Pending": "Pending",
    }
    original_unique = df['status'].unique()
    df['status'] = df['status'].map(status_mapping).fillna(df['status'])
    print(f"  ✓ Status: {list(original_unique)} → {list(df['status'].unique())}")

    # --- Department column ---
    dept_mapping = {
        "I.T.": "IT",
        "Human Resources": "HR",
        "sales": "Sales",
        "marketing": "Marketing",
    }
    # Apply mapping, keep original if not in mapping
    df['department'] = df['department'].replace(dept_mapping)
    print(f"  ✓ Department standardized: {sorted(df['department'].dropna().unique())}")

    # Strip extra whitespace from all string columns
    for col in df.select_dtypes(include='object').columns:
        df[col] = df[col].str.strip()

    return df


def handle_missing_values(df):
    """Handle missing values with appropriate strategies."""
    print("\n[4/6] Handling missing values...")

    missing_before = df.isnull().sum()
    print(f"  Missing values before:")
    for col, count in missing_before[missing_before > 0].items():
        print(f"    - {col}: {count} ({count/len(df)*100:.1f}%)")

    # Department: fill with mode (most common department)
    if df['department'].isnull().any():
        mode_dept = df['department'].mode()[0]
        df['department'] = df['department'].fillna(mode_dept)
        print(f"  ✓ Filled missing departments with mode: '{mode_dept}'")

    # completion_date: leave empty for Pending status (this is expected)
    # For Completed status with missing completion_date, mark as Pending
    mask = (df['status'] == 'Completed') & (df['completion_date'].isna() | (df['completion_date'] == ''))
    if mask.sum() > 0:
        df.loc[mask, 'status'] = 'Pending'
        print(f"  ✓ Fixed {mask.sum()} records: Completed without completion_date → Pending")

    missing_after = df.isnull().sum().sum()
    print(f"  ✓ Remaining missing values: {missing_after}")

    return df


def parse_dates(df):
    """Parse and standardize date columns."""
    print("\n[5/6] Parsing date columns...")

    # Parse assigned_date (handles mixed formats: YYYY-MM-DD and DD/MM/YYYY)
    df['assigned_date'] = pd.to_datetime(df['assigned_date'], format='mixed', dayfirst=True)
    print(f"  ✓ assigned_date parsed: {df['assigned_date'].dtype}")

    # Parse due_date
    df['due_date'] = pd.to_datetime(df['due_date'], format='mixed', dayfirst=True)
    print(f"  ✓ due_date parsed: {df['due_date'].dtype}")

    # Parse completion_date (may have empty strings)
    df['completion_date'] = df['completion_date'].replace('', np.nan)
    df['completion_date'] = pd.to_datetime(df['completion_date'], format='mixed', dayfirst=True, errors='coerce')
    print(f"  ✓ completion_date parsed: {df['completion_date'].dtype}")

    return df


def add_derived_columns(df):
    """Create derived columns for analysis."""
    print("\n[6/6] Creating derived columns...")

    # 1. Completion Time (in days) — only for completed records
    df['completion_time_days'] = np.nan
    completed_mask = df['status'] == 'Completed'
    df.loc[completed_mask, 'completion_time_days'] = (
        df.loc[completed_mask, 'completion_date'] - df.loc[completed_mask, 'assigned_date']
    ).dt.days
    avg_time = df['completion_time_days'].mean()
    print(f"  ✓ completion_time_days: avg = {avg_time:.1f} days")

    # 2. Is Overdue — training past due_date and still pending
    reference_date = pd.Timestamp('2026-03-31')
    df['is_overdue'] = False
    overdue_mask = (df['status'] == 'Pending') & (df['due_date'] < reference_date)
    df.loc[overdue_mask, 'is_overdue'] = True
    overdue_count = df['is_overdue'].sum()
    print(f"  ✓ is_overdue: {overdue_count} overdue records ({overdue_count/len(df)*100:.1f}%)")

    # 3. Month-Year of assignment (for trend analysis)
    df['assigned_month'] = df['assigned_date'].dt.to_period('M').astype(str)
    print(f"  ✓ assigned_month: {df['assigned_month'].nunique()} unique months")

    # 4. Days until due (from assignment)
    df['days_allowed'] = (df['due_date'] - df['assigned_date']).dt.days
    print(f"  ✓ days_allowed: avg = {df['days_allowed'].mean():.1f} days")

    # 5. On-time completion flag
    df['completed_on_time'] = False
    on_time_mask = (df['status'] == 'Completed') & (df['completion_date'] <= df['due_date'])
    df.loc[on_time_mask, 'completed_on_time'] = True
    on_time = df['completed_on_time'].sum()
    total_completed = completed_mask.sum()
    if total_completed > 0:
        print(f"  ✓ completed_on_time: {on_time}/{total_completed} ({on_time/total_completed*100:.1f}%)")

    return df


def generate_cleaning_report(raw_df, clean_df):
    """Generate a before-vs-after cleaning comparison."""
    print("\n" + "=" * 60)
    print("  DATA CLEANING REPORT")
    print("=" * 60)

    print(f"\n  {'Metric':<30} {'Before':<15} {'After':<15}")
    print(f"  {'-'*60}")
    print(f"  {'Total Rows':<30} {len(raw_df):<15} {len(clean_df):<15}")
    print(f"  {'Duplicate Rows':<30} {len(raw_df) - len(raw_df.drop_duplicates()):<15} {'0':<15}")
    print(f"  {'Missing Values':<30} {raw_df.isnull().sum().sum():<15} {clean_df.isnull().sum().sum():<15}")
    print(f"  {'Unique Status Values':<30} {raw_df['status'].nunique():<15} {clean_df['status'].nunique():<15}")
    print(f"  {'Unique Departments':<30} {raw_df['department'].nunique():<15} {clean_df['department'].nunique():<15}")
    print(f"  {'Date Columns Parsed':<30} {'No':<15} {'Yes':<15}")
    print(f"  {'Derived Columns Added':<30} {'0':<15} {'5':<15}")

    print(f"\n  New Columns Added:")
    new_cols = [c for c in clean_df.columns if c not in raw_df.columns]
    for col in new_cols:
        print(f"    + {col}")

    print("=" * 60)


def main():
    print("=" * 60)
    print("  LMS Data Cleaning Pipeline")
    print("=" * 60)

    # Paths
    base_dir = os.path.dirname(os.path.dirname(__file__))
    raw_path = os.path.join(base_dir, "data", "lms_raw.csv")
    clean_path = os.path.join(base_dir, "data", "lms_cleaned.csv")

    # Load
    raw_df = pd.read_csv(raw_path)
    df = load_raw_data(raw_path)

    # Clean
    df = remove_duplicates(df)
    df = standardize_text(df)
    df = handle_missing_values(df)
    df = parse_dates(df)
    df = add_derived_columns(df)

    # Save
    df.to_csv(clean_path, index=False)
    print(f"\n✅ Cleaned data saved to: {clean_path}")

    # Report
    generate_cleaning_report(raw_df, df)


if __name__ == "__main__":
    main()
