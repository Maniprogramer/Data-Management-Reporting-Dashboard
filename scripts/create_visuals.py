"""
Visualization Script for LMS Dataset.
Creates professional charts using Matplotlib and Seaborn.
"""

import pandas as pd
import numpy as np
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import matplotlib.ticker as mtick
import seaborn as sns
from matplotlib.gridspec import GridSpec
import os
import warnings
warnings.filterwarnings('ignore')

# ─── Style Configuration ───
sns.set_theme(style="whitegrid")
plt.rcParams.update({
    'figure.facecolor': '#FAFAFA',
    'axes.facecolor': '#FAFAFA',
    'font.family': 'sans-serif',
    'font.size': 11,
    'axes.titlesize': 14,
    'axes.titleweight': 'bold',
    'axes.labelsize': 12,
})

# Color palette
COLORS = {
    'primary': '#2563EB',
    'success': '#16A34A',
    'warning': '#F59E0B',
    'danger': '#DC2626',
    'info': '#0891B2',
    'muted': '#94A3B8',
    'completed': '#16A34A',
    'pending': '#F59E0B',
    'overdue': '#DC2626',
}

DEPT_PALETTE = ['#2563EB', '#7C3AED', '#0891B2', '#16A34A', '#F59E0B', '#DC2626']


def load_data(filepath):
    """Load cleaned dataset."""
    df = pd.read_csv(filepath, parse_dates=['assigned_date', 'due_date', 'completion_date'])
    return df


def chart_completion_overview(df, save_dir):
    """Pie/donut chart: Overall completion status breakdown."""
    fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(14, 6))

    # Left: Donut chart for status
    status_counts = df['status'].value_counts()
    colors = [COLORS['completed'], COLORS['pending']]
    wedges, texts, autotexts = ax1.pie(
        status_counts.values,
        labels=status_counts.index,
        colors=colors,
        autopct='%1.1f%%',
        startangle=90,
        pctdistance=0.75,
        wedgeprops=dict(width=0.4, edgecolor='white', linewidth=2),
        textprops={'fontsize': 13, 'fontweight': 'bold'},
    )
    for autotext in autotexts:
        autotext.set_fontsize(14)
        autotext.set_fontweight('bold')
    ax1.set_title('Training Completion Status', pad=20, fontsize=15)

    # Center circle text
    total = len(df)
    ax1.text(0, 0, f'{total}\nTotal', ha='center', va='center',
             fontsize=16, fontweight='bold', color='#1E293B')

    # Right: Completed on-time vs late
    completed = df[df['status'] == 'Completed']
    on_time = completed['completed_on_time'].sum()
    late = len(completed) - on_time
    bars = ax2.bar(['On Time', 'Late'], [on_time, late],
                    color=[COLORS['success'], COLORS['warning']],
                    width=0.5, edgecolor='white', linewidth=2)
    for bar, val in zip(bars, [on_time, late]):
        ax2.text(bar.get_x() + bar.get_width()/2, bar.get_height() + 15,
                 f'{val}', ha='center', va='bottom', fontweight='bold', fontsize=14)
    ax2.set_title('Completed Trainings: On Time vs Late', fontsize=15)
    ax2.set_ylabel('Count')
    ax2.spines['top'].set_visible(False)
    ax2.spines['right'].set_visible(False)

    plt.tight_layout()
    path = os.path.join(save_dir, 'completion_overview.png')
    fig.savefig(path, dpi=150, bbox_inches='tight')
    plt.close()
    print(f"  ✓ Saved: completion_overview.png")


def chart_department_performance(df, save_dir):
    """Bar chart: Completion rate by department."""
    fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(15, 6))

    # Left: Stacked bar chart
    dept_status = df.groupby('department')['status'].value_counts().unstack(fill_value=0)
    dept_status = dept_status.sort_values('Completed', ascending=True)

    dept_status.plot(
        kind='barh',
        stacked=True,
        ax=ax1,
        color=[COLORS['completed'], COLORS['pending']],
        edgecolor='white',
        linewidth=1.5,
    )
    ax1.set_title('Trainings by Department & Status', fontsize=14)
    ax1.set_xlabel('Number of Trainings')
    ax1.set_ylabel('')
    ax1.legend(title='Status', loc='lower right')
    ax1.spines['top'].set_visible(False)
    ax1.spines['right'].set_visible(False)

    # Right: Completion rate bars
    dept_rate = df.groupby('department').apply(
        lambda x: round((x['status'] == 'Completed').sum() / len(x) * 100, 1)
    ).sort_values(ascending=True)

    bars = ax2.barh(dept_rate.index, dept_rate.values, color=DEPT_PALETTE[:len(dept_rate)],
                     edgecolor='white', linewidth=1.5, height=0.6)
    for bar, val in zip(bars, dept_rate.values):
        ax2.text(bar.get_width() + 0.5, bar.get_y() + bar.get_height()/2,
                 f'{val}%', ha='left', va='center', fontweight='bold', fontsize=12)

    ax2.set_title('Completion Rate by Department', fontsize=14)
    ax2.set_xlabel('Completion Rate (%)')
    ax2.set_xlim(0, 100)
    ax2.spines['top'].set_visible(False)
    ax2.spines['right'].set_visible(False)

    plt.tight_layout()
    path = os.path.join(save_dir, 'department_performance.png')
    fig.savefig(path, dpi=150, bbox_inches='tight')
    plt.close()
    print(f"  ✓ Saved: department_performance.png")


def chart_course_analysis(df, save_dir):
    """Bar chart: Course-wise completion rates and training count."""
    fig, ax = plt.subplots(figsize=(14, 7))

    course_data = df.groupby('course_name').agg(
        total=('status', 'count'),
        completed=('status', lambda x: (x == 'Completed').sum()),
    ).reset_index()
    course_data['rate'] = round(course_data['completed'] / course_data['total'] * 100, 1)
    course_data = course_data.sort_values('rate', ascending=True)

    # Color bars based on rate (gradient)
    colors = []
    for rate in course_data['rate']:
        if rate >= 75:
            colors.append(COLORS['success'])
        elif rate >= 60:
            colors.append(COLORS['info'])
        elif rate >= 50:
            colors.append(COLORS['warning'])
        else:
            colors.append(COLORS['danger'])

    bars = ax.barh(course_data['course_name'], course_data['rate'],
                    color=colors, edgecolor='white', linewidth=1.5, height=0.6)

    for bar, val, total in zip(bars, course_data['rate'], course_data['total']):
        ax.text(bar.get_width() + 0.8, bar.get_y() + bar.get_height()/2,
                f'{val}%  (n={total})', ha='left', va='center', fontsize=11, fontweight='bold')

    ax.set_title('Course Completion Rates', fontsize=16, pad=15)
    ax.set_xlabel('Completion Rate (%)')
    ax.set_xlim(0, 105)
    ax.axvline(x=70, color=COLORS['muted'], linestyle='--', alpha=0.5, label='Target (70%)')
    ax.legend(fontsize=10)
    ax.spines['top'].set_visible(False)
    ax.spines['right'].set_visible(False)

    plt.tight_layout()
    path = os.path.join(save_dir, 'course_analysis.png')
    fig.savefig(path, dpi=150, bbox_inches='tight')
    plt.close()
    print(f"  ✓ Saved: course_analysis.png")


def chart_monthly_trends(df, save_dir):
    """Line chart: Monthly training trends."""
    fig, ax = plt.subplots(figsize=(14, 6))

    monthly = df.groupby('assigned_month').agg(
        total=('status', 'count'),
        completed=('status', lambda x: (x == 'Completed').sum()),
        pending=('status', lambda x: (x == 'Pending').sum()),
    ).reset_index()
    monthly = monthly.sort_values('assigned_month')

    x = range(len(monthly))
    ax.plot(x, monthly['total'], marker='o', color=COLORS['primary'],
            linewidth=2.5, markersize=8, label='Total Assigned', zorder=3)
    ax.plot(x, monthly['completed'], marker='s', color=COLORS['success'],
            linewidth=2.5, markersize=8, label='Completed', zorder=3)
    ax.plot(x, monthly['pending'], marker='^', color=COLORS['warning'],
            linewidth=2.5, markersize=8, label='Pending', zorder=3)

    ax.fill_between(x, monthly['completed'], alpha=0.1, color=COLORS['success'])
    ax.fill_between(x, monthly['pending'], alpha=0.1, color=COLORS['warning'])

    ax.set_xticks(x)
    ax.set_xticklabels(monthly['assigned_month'], rotation=45, ha='right')
    ax.set_title('Monthly Training Assignment Trends', fontsize=16, pad=15)
    ax.set_xlabel('Month')
    ax.set_ylabel('Number of Trainings')
    ax.legend(fontsize=11)
    ax.spines['top'].set_visible(False)
    ax.spines['right'].set_visible(False)

    plt.tight_layout()
    path = os.path.join(save_dir, 'monthly_trends.png')
    fig.savefig(path, dpi=150, bbox_inches='tight')
    plt.close()
    print(f"  ✓ Saved: monthly_trends.png")


def chart_overdue_analysis(df, save_dir):
    """Charts for overdue training analysis."""
    fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(14, 6))

    # Left: Overdue by department
    overdue_dept = df[df['is_overdue'] == True].groupby('department').size().sort_values(ascending=True)

    if len(overdue_dept) > 0:
        bars = ax1.barh(overdue_dept.index, overdue_dept.values,
                         color=COLORS['danger'], edgecolor='white', linewidth=1.5, height=0.5,
                         alpha=0.85)
        for bar, val in zip(bars, overdue_dept.values):
            ax1.text(bar.get_width() + 0.3, bar.get_y() + bar.get_height()/2,
                     str(val), ha='left', va='center', fontweight='bold', fontsize=12)

    ax1.set_title('Overdue Trainings by Department', fontsize=14)
    ax1.set_xlabel('Number of Overdue Trainings')
    ax1.spines['top'].set_visible(False)
    ax1.spines['right'].set_visible(False)

    # Right: Overdue by course
    overdue_course = df[df['is_overdue'] == True].groupby('course_name').size().sort_values(ascending=True)

    if len(overdue_course) > 0:
        bars = ax2.barh(overdue_course.index, overdue_course.values,
                         color=COLORS['warning'], edgecolor='white', linewidth=1.5, height=0.5,
                         alpha=0.85)
        for bar, val in zip(bars, overdue_course.values):
            ax2.text(bar.get_width() + 0.3, bar.get_y() + bar.get_height()/2,
                     str(val), ha='left', va='center', fontweight='bold', fontsize=12)

    ax2.set_title('Overdue Trainings by Course', fontsize=14)
    ax2.set_xlabel('Number of Overdue Trainings')
    ax2.spines['top'].set_visible(False)
    ax2.spines['right'].set_visible(False)

    plt.tight_layout()
    path = os.path.join(save_dir, 'overdue_analysis.png')
    fig.savefig(path, dpi=150, bbox_inches='tight')
    plt.close()
    print(f"  ✓ Saved: overdue_analysis.png")


def chart_completion_time_distribution(df, save_dir):
    """Histogram/box: Completion time distribution by department."""
    fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(14, 6))

    completed = df[df['status'] == 'Completed'].copy()

    # Left: Histogram
    ax1.hist(completed['completion_time_days'].dropna(), bins=25,
             color=COLORS['primary'], edgecolor='white', linewidth=1.2, alpha=0.8)
    mean_time = completed['completion_time_days'].mean()
    ax1.axvline(x=mean_time, color=COLORS['danger'], linestyle='--', linewidth=2,
                label=f'Mean: {mean_time:.1f} days')
    ax1.set_title('Completion Time Distribution', fontsize=14)
    ax1.set_xlabel('Days to Complete')
    ax1.set_ylabel('Frequency')
    ax1.legend(fontsize=11)
    ax1.spines['top'].set_visible(False)
    ax1.spines['right'].set_visible(False)

    # Right: Box plot by department
    dept_order = completed.groupby('department')['completion_time_days'].median().sort_values().index
    sns.boxplot(data=completed, x='department', y='completion_time_days',
                order=dept_order, palette=DEPT_PALETTE[:len(dept_order)], ax=ax2,
                linewidth=1.5, fliersize=4)
    ax2.set_title('Completion Time by Department', fontsize=14)
    ax2.set_xlabel('')
    ax2.set_ylabel('Days to Complete')
    ax2.tick_params(axis='x', rotation=30)
    ax2.spines['top'].set_visible(False)
    ax2.spines['right'].set_visible(False)

    plt.tight_layout()
    path = os.path.join(save_dir, 'completion_time.png')
    fig.savefig(path, dpi=150, bbox_inches='tight')
    plt.close()
    print(f"  ✓ Saved: completion_time.png")


def chart_heatmap(df, save_dir):
    """Heatmap: Department × Course completion rates."""
    fig, ax = plt.subplots(figsize=(14, 7))

    matrix = df.pivot_table(
        values='status',
        index='department',
        columns='course_name',
        aggfunc=lambda x: round((x == 'Completed').sum() / len(x) * 100, 1)
    ).fillna(0)

    sns.heatmap(
        matrix, annot=True, fmt='.0f', cmap='RdYlGn',
        linewidths=2, linecolor='white',
        cbar_kws={'label': 'Completion Rate (%)', 'shrink': 0.8},
        ax=ax, vmin=0, vmax=100,
        annot_kws={'fontsize': 11, 'fontweight': 'bold'},
    )

    ax.set_title('Completion Rate: Department × Course (%)', fontsize=16, pad=15)
    ax.set_xlabel('')
    ax.set_ylabel('')
    ax.tick_params(axis='x', rotation=35)

    plt.tight_layout()
    path = os.path.join(save_dir, 'heatmap_dept_course.png')
    fig.savefig(path, dpi=150, bbox_inches='tight')
    plt.close()
    print(f"  ✓ Saved: heatmap_dept_course.png")


def main():
    print("=" * 60)
    print("  LMS Data Visualization")
    print("=" * 60)

    base_dir = os.path.dirname(os.path.dirname(__file__))
    clean_path = os.path.join(base_dir, "data", "lms_cleaned.csv")
    visuals_dir = os.path.join(base_dir, "visuals")
    os.makedirs(visuals_dir, exist_ok=True)

    df = load_data(clean_path)
    print(f"\n✓ Loaded {len(df)} records")
    print(f"\nGenerating charts...")

    chart_completion_overview(df, visuals_dir)
    chart_department_performance(df, visuals_dir)
    chart_course_analysis(df, visuals_dir)
    chart_monthly_trends(df, visuals_dir)
    chart_overdue_analysis(df, visuals_dir)
    chart_completion_time_distribution(df, visuals_dir)
    chart_heatmap(df, visuals_dir)

    print(f"\n✅ All 7 charts saved to: {visuals_dir}/")


if __name__ == "__main__":
    main()
