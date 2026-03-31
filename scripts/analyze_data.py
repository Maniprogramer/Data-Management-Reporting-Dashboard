"""
Data Analysis & Report Generation Script for LMS Dataset.
Generates summary statistics, department/course analysis, and exports to Excel.
"""

import pandas as pd
import numpy as np
import os
import warnings
warnings.filterwarnings('ignore')


def load_cleaned_data(filepath):
    """Load the cleaned dataset."""
    df = pd.read_csv(filepath, parse_dates=['assigned_date', 'due_date', 'completion_date'])
    return df


def overall_kpis(df):
    """Calculate top-level KPIs."""
    total_records = len(df)
    total_employees = df['employee_id'].nunique()
    total_courses = df['course_name'].nunique()

    completed = (df['status'] == 'Completed').sum()
    pending = (df['status'] == 'Pending').sum()
    overdue = df['is_overdue'].sum()

    completion_rate = (completed / total_records) * 100
    overdue_rate = (overdue / pending) * 100 if pending > 0 else 0

    avg_completion_time = df.loc[df['status'] == 'Completed', 'completion_time_days'].mean()
    on_time_rate = (df['completed_on_time'].sum() / completed) * 100 if completed > 0 else 0

    kpis = pd.DataFrame({
        'KPI': [
            'Total Training Records',
            'Unique Employees',
            'Total Courses',
            'Completed Trainings',
            'Pending Trainings',
            'Overdue Trainings',
            'Completion Rate (%)',
            'Overdue Rate (% of Pending)',
            'Avg Completion Time (Days)',
            'On-Time Completion Rate (%)',
        ],
        'Value': [
            total_records,
            total_employees,
            total_courses,
            completed,
            pending,
            overdue,
            round(completion_rate, 1),
            round(overdue_rate, 1),
            round(avg_completion_time, 1),
            round(on_time_rate, 1),
        ]
    })

    print("\n📊 Overall KPIs:")
    print(kpis.to_string(index=False))
    return kpis


def department_analysis(df):
    """Analyze performance by department."""
    dept = df.groupby('department').agg(
        total_trainings=('status', 'count'),
        completed=('status', lambda x: (x == 'Completed').sum()),
        pending=('status', lambda x: (x == 'Pending').sum()),
        overdue=('is_overdue', 'sum'),
        avg_completion_days=('completion_time_days', 'mean'),
        on_time_completions=('completed_on_time', 'sum'),
    ).reset_index()

    dept['completion_rate'] = round((dept['completed'] / dept['total_trainings']) * 100, 1)
    dept['overdue_rate'] = round((dept['overdue'] / dept['pending']) * 100, 1)
    dept['on_time_rate'] = round((dept['on_time_completions'] / dept['completed']) * 100, 1)
    dept['avg_completion_days'] = round(dept['avg_completion_days'], 1)

    # Sort by completion rate
    dept = dept.sort_values('completion_rate', ascending=False)

    print("\n📊 Department-wise Performance:")
    print(dept[['department', 'total_trainings', 'completed', 'pending',
                'overdue', 'completion_rate', 'avg_completion_days']].to_string(index=False))
    return dept


def course_analysis(df):
    """Analyze performance by course."""
    course = df.groupby('course_name').agg(
        total_assigned=('status', 'count'),
        completed=('status', lambda x: (x == 'Completed').sum()),
        pending=('status', lambda x: (x == 'Pending').sum()),
        overdue=('is_overdue', 'sum'),
        avg_completion_days=('completion_time_days', 'mean'),
        on_time_completions=('completed_on_time', 'sum'),
    ).reset_index()

    course['completion_rate'] = round((course['completed'] / course['total_assigned']) * 100, 1)
    course['avg_completion_days'] = round(course['avg_completion_days'], 1)

    # Sort by completion rate
    course = course.sort_values('completion_rate', ascending=False)

    print("\n📊 Course-wise Performance:")
    print(course[['course_name', 'total_assigned', 'completed', 'pending',
                  'completion_rate', 'avg_completion_days']].to_string(index=False))
    return course


def monthly_trends(df):
    """Analyze monthly assignment and completion trends."""
    monthly = df.groupby('assigned_month').agg(
        total_assigned=('status', 'count'),
        completed=('status', lambda x: (x == 'Completed').sum()),
        pending=('status', lambda x: (x == 'Pending').sum()),
    ).reset_index()

    monthly['completion_rate'] = round((monthly['completed'] / monthly['total_assigned']) * 100, 1)
    monthly = monthly.sort_values('assigned_month')

    print("\n📊 Monthly Trends:")
    print(monthly.to_string(index=False))
    return monthly


def overdue_employees(df):
    """Identify employees with overdue trainings."""
    overdue = df[df['is_overdue'] == True].copy()

    if len(overdue) == 0:
        print("\n✅ No overdue trainings found!")
        return pd.DataFrame()

    # Calculate days overdue
    reference_date = pd.Timestamp('2026-03-31')
    overdue['days_overdue'] = (reference_date - overdue['due_date']).dt.days

    # Summary per employee
    emp_overdue = overdue.groupby(['employee_id', 'employee_name', 'department']).agg(
        overdue_courses=('course_name', 'count'),
        courses_list=('course_name', lambda x: ', '.join(x)),
        max_days_overdue=('days_overdue', 'max'),
        avg_days_overdue=('days_overdue', 'mean'),
    ).reset_index()

    emp_overdue['avg_days_overdue'] = round(emp_overdue['avg_days_overdue'], 0).astype(int)
    emp_overdue = emp_overdue.sort_values('overdue_courses', ascending=False)

    print(f"\n📊 Top Overdue Employees (showing top 15 of {len(emp_overdue)}):")
    print(emp_overdue.head(15)[['employee_id', 'employee_name', 'department',
                                 'overdue_courses', 'max_days_overdue']].to_string(index=False))
    return emp_overdue


def department_course_matrix(df):
    """Create a department × course completion rate matrix."""
    matrix = df.pivot_table(
        values='status',
        index='department',
        columns='course_name',
        aggfunc=lambda x: round((x == 'Completed').sum() / len(x) * 100, 1)
    )
    matrix = matrix.fillna(0)

    print("\n📊 Department × Course Completion Rate Matrix (%):")
    print(matrix.to_string())
    return matrix


def generate_insights(df, dept_df, course_df):
    """Generate key insights for the README and presentation."""
    insights = []

    # 1. Overall completion
    comp_rate = round((df['status'] == 'Completed').sum() / len(df) * 100, 1)
    insights.append(f"Overall training completion rate is {comp_rate}%")

    # 2. Best department
    best_dept = dept_df.iloc[0]
    insights.append(f"{best_dept['department']} department leads with {best_dept['completion_rate']}% completion rate")

    # 3. Worst department
    worst_dept = dept_df.iloc[-1]
    insights.append(f"{worst_dept['department']} department needs improvement at {worst_dept['completion_rate']}% completion rate")

    # 4. Course insights
    best_course = course_df.iloc[0]
    worst_course = course_df.iloc[-1]
    insights.append(f"'{best_course['course_name']}' has the highest completion ({best_course['completion_rate']}%)")
    insights.append(f"'{worst_course['course_name']}' has the lowest completion ({worst_course['completion_rate']}%)")

    # 5. Overdue
    overdue_count = df['is_overdue'].sum()
    overdue_pct = round(overdue_count / len(df) * 100, 1)
    insights.append(f"{overdue_count} trainings are overdue ({overdue_pct}% of total)")

    # 6. Average completion time
    avg_time = round(df.loc[df['status'] == 'Completed', 'completion_time_days'].mean(), 1)
    insights.append(f"Average completion time is {avg_time} days")

    # 7. On-time performance
    completed = df[df['status'] == 'Completed']
    on_time = completed['completed_on_time'].sum()
    on_time_pct = round(on_time / len(completed) * 100, 1) if len(completed) > 0 else 0
    insights.append(f"{on_time_pct}% of completed trainings were finished on time")

    print("\n" + "=" * 60)
    print("  KEY INSIGHTS")
    print("=" * 60)
    for i, insight in enumerate(insights, 1):
        print(f"  {i}. {insight}")
    print("=" * 60)

    return insights


def export_to_excel(kpis, dept_df, course_df, monthly_df, overdue_df, matrix_df, insights, output_path):
    """Export all analysis results to a formatted Excel workbook."""
    print(f"\n📁 Exporting to Excel: {output_path}")

    with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
        # Sheet 1: KPIs
        kpis.to_excel(writer, sheet_name='KPIs', index=False)

        # Sheet 2: Department Analysis
        dept_df.to_excel(writer, sheet_name='Department Analysis', index=False)

        # Sheet 3: Course Analysis
        course_df.to_excel(writer, sheet_name='Course Analysis', index=False)

        # Sheet 4: Monthly Trends
        monthly_df.to_excel(writer, sheet_name='Monthly Trends', index=False)

        # Sheet 5: Overdue Employees
        if len(overdue_df) > 0:
            overdue_df.to_excel(writer, sheet_name='Overdue Employees', index=False)

        # Sheet 6: Department-Course Matrix
        matrix_df.to_excel(writer, sheet_name='Dept-Course Matrix')

        # Sheet 7: Key Insights
        insights_df = pd.DataFrame({'#': range(1, len(insights)+1), 'Insight': insights})
        insights_df.to_excel(writer, sheet_name='Key Insights', index=False)

    print(f"  ✓ Exported 7 sheets successfully")


def main():
    print("=" * 60)
    print("  LMS Data Analysis & Report Generation")
    print("=" * 60)

    # Paths
    base_dir = os.path.dirname(os.path.dirname(__file__))
    clean_path = os.path.join(base_dir, "data", "lms_cleaned.csv")
    report_path = os.path.join(base_dir, "reports", "summary_report.xlsx")
    os.makedirs(os.path.dirname(report_path), exist_ok=True)

    # Load data
    df = load_cleaned_data(clean_path)
    print(f"\n✓ Loaded {len(df)} cleaned records")

    # Run analyses
    kpis = overall_kpis(df)
    dept_df = department_analysis(df)
    course_df = course_analysis(df)
    monthly_df = monthly_trends(df)
    overdue_df = overdue_employees(df)
    matrix_df = department_course_matrix(df)
    insights = generate_insights(df, dept_df, course_df)

    # Export
    export_to_excel(kpis, dept_df, course_df, monthly_df, overdue_df, matrix_df, insights, report_path)

    print(f"\n✅ Analysis complete! Report saved to: {report_path}")


if __name__ == "__main__":
    main()
