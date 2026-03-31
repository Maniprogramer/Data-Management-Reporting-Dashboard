"""
Master Pipeline Script - Runs the entire LMS data pipeline.
Usage: python3 run_pipeline.py
"""

import subprocess
import os
import time


def run_step(step_num, total, description, script_name, base_dir):
    """Run a pipeline step and handle errors."""
    print(f"\n{'='*60}")
    print(f"  STEP {step_num}/{total}: {description}")
    print(f"{'='*60}")

    script_path = os.path.join(base_dir, "scripts", script_name)

    start = time.time()
    result = subprocess.run(
        ["python3", script_path],
        capture_output=True, text=True,
        cwd=base_dir
    )

    elapsed = time.time() - start

    if result.stdout:
        print(result.stdout)

    if result.returncode != 0:
        print(f"\n❌ Step {step_num} FAILED ({elapsed:.1f}s)")
        if result.stderr:
            print(f"Error:\n{result.stderr}")
        return False

    print(f"\n✅ Step {step_num} completed in {elapsed:.1f}s")
    return True


def main():
    base_dir = os.path.dirname(os.path.abspath(__file__))

    print("\n" + "=" * 60)
    print("  🚀 LMS DATA MANAGEMENT & REPORTING PIPELINE")
    print("  " + "=" * 56)
    print(f"  Started at: {time.strftime('%Y-%m-%d %H:%M:%S')}")
    print("=" * 60)

    pipeline_start = time.time()

    steps = [
        (1, "Generate Synthetic Dataset", "generate_data.py"),
        (2, "Clean & Transform Data", "clean_data.py"),
        (3, "Analyze Data & Generate Report", "analyze_data.py"),
        (4, "Create Visualizations", "create_visuals.py"),
        (5, "Build Excel Dashboard (VLOOKUP, Pivots, Filters)", "create_excel_dashboard.py"),
    ]

    total = len(steps)

    for step_num, description, script in steps:
        success = run_step(step_num, total, description, script, base_dir)
        if not success:
            print(f"\n💥 Pipeline failed at Step {step_num}. Fix the error and re-run.")
            exit(1)

    total_time = time.time() - pipeline_start

    print("\n" + "=" * 60)
    print("  ✅ PIPELINE COMPLETE")
    print("=" * 60)
    print(f"  Total time: {total_time:.1f}s")
    print(f"\n  📦 Outputs:")
    print(f"    📄 data/lms_raw.csv           → Raw dataset (5,000+ records)")
    print(f"    📄 data/lms_cleaned.csv        → Cleaned dataset")
    print(f"    📊 reports/summary_report.xlsx  → Multi-sheet analysis report")
    print(f"    📊 reports/dashboard.xlsx       → Excel dashboard (VLOOKUP, pivots, filters)")
    print(f"    📈 visuals/*.png               → 7 professional charts")
    print(f"\n  👉 Open reports/dashboard.xlsx in Excel to explore!")
    print("=" * 60)


if __name__ == "__main__":
    main()
