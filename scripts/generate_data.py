"""
Script to generate synthetic LMS (Learning Management System) dataset.
Simulates realistic employee training data with varied completion patterns.
"""

import pandas as pd
import numpy as np
from datetime import datetime, timedelta
import random
import os

# Set seed for reproducibility
np.random.seed(42)
random.seed(42)

# ─── Configuration ───
N_EMPLOYEES = 350
N_RECORDS = 5000

# Realistic department distribution (not uniform)
DEPARTMENTS = {
    "IT": 0.25,
    "HR": 0.12,
    "Sales": 0.22,
    "Finance": 0.13,
    "Operations": 0.18,
    "Marketing": 0.10,
}

# Training courses mapped to departments (some are mandatory for all)
COURSES = {
    "Data Security Awareness": {"mandatory": True, "avg_days": 5},
    "Workplace Safety": {"mandatory": True, "avg_days": 3},
    "Anti-Harassment Training": {"mandatory": True, "avg_days": 2},
    "Excel for Business": {"mandatory": False, "avg_days": 10},
    "Python Basics": {"mandatory": False, "avg_days": 14},
    "Project Management": {"mandatory": False, "avg_days": 12},
    "Communication Skills": {"mandatory": False, "avg_days": 7},
    "Leadership Fundamentals": {"mandatory": False, "avg_days": 8},
    "Customer Service Excellence": {"mandatory": False, "avg_days": 6},
    "SQL Fundamentals": {"mandatory": False, "avg_days": 10},
}

# First and last names for realistic employee names
FIRST_NAMES = [
    "Aarav", "Vivaan", "Aditya", "Vihaan", "Arjun", "Sai", "Reyansh",
    "Ayaan", "Krishna", "Ishaan", "Ananya", "Diya", "Myra", "Sara",
    "Aadhya", "Priya", "Sneha", "Kavya", "Riya", "Pooja", "Rahul",
    "Amit", "Vikram", "Suresh", "Neha", "Deepak", "Mohan", "Lakshmi",
    "Rajesh", "Sunita", "Kiran", "Anjali", "Rohit", "Meera", "Arun",
    "Divya", "Manish", "Rekha", "Gaurav", "Swati", "Nitin", "Pallavi",
    "Sachin", "Nandini", "Varun", "Shruti", "Pranav", "Jyoti", "Tanya",
    "Harish"
]

LAST_NAMES = [
    "Sharma", "Patel", "Reddy", "Kumar", "Singh", "Gupta", "Nair",
    "Rao", "Joshi", "Iyer", "Verma", "Mehta", "Das", "Pillai",
    "Mishra", "Chopra", "Sinha", "Bhat", "Kulkarni", "Deshpande",
    "Malhotra", "Saxena", "Tiwari", "Pandey", "Mukherjee", "Sen",
    "Ghosh", "Banerjee", "Desai", "Shah", "Modi", "Trivedi"
]


def generate_employees(n):
    """Generate a realistic employee roster."""
    employees = []
    emp_id_start = 1001

    dept_names = list(DEPARTMENTS.keys())
    dept_probs = list(DEPARTMENTS.values())

    for i in range(n):
        emp = {
            "employee_id": emp_id_start + i,
            "employee_name": f"{random.choice(FIRST_NAMES)} {random.choice(LAST_NAMES)}",
            "department": np.random.choice(dept_names, p=dept_probs),
        }
        employees.append(emp)

    return pd.DataFrame(employees)


def generate_training_records(employees_df, n_records):
    """Generate training assignment records with realistic patterns."""
    records = []
    course_names = list(COURSES.keys())

    # Date range: last 12 months
    end_date = datetime(2026, 3, 31)
    start_date = end_date - timedelta(days=365)

    for _ in range(n_records):
        emp = employees_df.sample(1).iloc[0]
        course = random.choice(course_names)
        course_info = COURSES[course]

        # Random assigned date within the range
        days_offset = random.randint(0, 330)
        assigned_date = start_date + timedelta(days=days_offset)

        # Due date: assigned + buffer (14-45 days depending on course)
        due_buffer = course_info["avg_days"] + random.randint(7, 30)
        due_date = assigned_date + timedelta(days=due_buffer)

        # Determine completion status with realistic probabilities
        # Mandatory courses have higher completion rates
        if course_info["mandatory"]:
            completion_prob = 0.82
        else:
            # Vary by department
            dept_completion = {
                "IT": 0.72, "HR": 0.78, "Sales": 0.58,
                "Finance": 0.68, "Operations": 0.62, "Marketing": 0.65
            }
            completion_prob = dept_completion.get(emp["department"], 0.65)

        if random.random() < completion_prob:
            status = "Completed"
            # Completion time varies
            comp_days = max(1, int(np.random.normal(course_info["avg_days"], 3)))
            completion_date = assigned_date + timedelta(days=comp_days)
        else:
            status = "Pending"
            completion_date = None

        record = {
            "employee_id": emp["employee_id"],
            "employee_name": emp["employee_name"],
            "department": emp["department"],
            "course_name": course,
            "assigned_date": assigned_date.strftime("%Y-%m-%d"),
            "due_date": due_date.strftime("%Y-%m-%d"),
            "completion_date": completion_date.strftime("%Y-%m-%d") if completion_date else "",
            "status": status,
        }
        records.append(record)

    return pd.DataFrame(records)


def introduce_data_issues(df):
    """
    Introduce realistic data quality issues for the cleaning phase.
    This makes the project more realistic and demonstrates cleaning skills.
    """
    df = df.copy()
    n = len(df)

    # 1. Add some duplicate rows (~2%)
    n_dupes = int(n * 0.02)
    dupes = df.sample(n_dupes, random_state=42)
    df = pd.concat([df, dupes], ignore_index=True)

    # 2. Introduce inconsistent status values (~3%)
    idx = df.sample(int(n * 0.03), random_state=10).index
    status_variants = ["completed", "COMPLETED", "Completd", "pending", "PENDING", "Pendig"]
    for i in idx:
        df.at[i, "status"] = random.choice(status_variants)

    # 3. Add some missing values in non-critical columns (~1.5%)
    missing_idx = df.sample(int(n * 0.015), random_state=20).index
    df.loc[missing_idx, "department"] = np.nan

    # 4. Some inconsistent department names
    dept_typos = {"IT": "I.T.", "HR": "Human Resources", "Sales": "sales", "Marketing": "marketing"}
    typo_idx = df.sample(int(n * 0.02), random_state=30).index
    for i in typo_idx:
        dept = df.at[i, "department"]
        if dept in dept_typos:
            df.at[i, "department"] = dept_typos[dept]

    # 5. Some date format inconsistencies
    date_idx = df.sample(int(n * 0.02), random_state=40).index
    for i in date_idx:
        if df.at[i, "assigned_date"]:
            try:
                d = datetime.strptime(str(df.at[i, "assigned_date"]), "%Y-%m-%d")
                df.at[i, "assigned_date"] = d.strftime("%d/%m/%Y")
            except (ValueError, TypeError):
                pass

    return df


def main():
    print("=" * 60)
    print("  LMS Dataset Generator")
    print("=" * 60)

    # Step 1: Generate employees
    print("\n[1/4] Generating employee roster...")
    employees = generate_employees(N_EMPLOYEES)
    print(f"  ✓ Created {len(employees)} employees across {employees['department'].nunique()} departments")

    # Step 2: Generate training records
    print("\n[2/4] Generating training records...")
    records = generate_training_records(employees, N_RECORDS)
    print(f"  ✓ Created {len(records)} training records")

    # Step 3: Introduce data issues for realistic cleaning
    print("\n[3/4] Introducing realistic data quality issues...")
    raw_data = introduce_data_issues(records)
    print(f"  ✓ Added duplicates, missing values, and inconsistencies")
    print(f"  ✓ Final raw dataset: {len(raw_data)} rows")

    # Step 4: Save
    output_dir = os.path.join(os.path.dirname(os.path.dirname(__file__)), "data")
    os.makedirs(output_dir, exist_ok=True)

    output_path = os.path.join(output_dir, "lms_raw.csv")
    raw_data.to_csv(output_path, index=False)
    print(f"\n[4/4] Saved to: {output_path}")

    # Quick summary
    print("\n" + "=" * 60)
    print("  Quick Data Summary")
    print("=" * 60)
    print(f"  Total Records    : {len(raw_data)}")
    print(f"  Unique Employees : {raw_data['employee_id'].nunique()}")
    print(f"  Departments      : {raw_data['department'].nunique()}")
    print(f"  Courses          : {raw_data['course_name'].nunique()}")
    print(f"  Status Values    : {raw_data['status'].unique().tolist()}")
    print(f"  Missing Values   : {raw_data.isnull().sum().sum()}")
    print("=" * 60)


if __name__ == "__main__":
    main()
