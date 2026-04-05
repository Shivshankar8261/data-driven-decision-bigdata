"""
Dataset Generator for Big Data Capstone Project
-------------------------------------------------
Generates ~123,847 rows of realistic organizational data simulating
data-driven decision making across departments and regions.

All columns have realistic correlations and business-logic-driven distributions.
"""

import numpy as np
import pandas as pd
from datetime import datetime, timedelta
import warnings
warnings.filterwarnings('ignore')

np.random.seed(42)

# ---------------------------------------------------------------------------
# Configuration
# ---------------------------------------------------------------------------
NUM_EMPLOYEES = 2500
START_DATE = datetime(2024, 1, 1)
END_DATE = datetime(2025, 12, 31)
TARGET_ROWS = 123847

DEPARTMENTS = {
    'Sales':       {'size_weight': 0.25, 'training_mean': 45, 'training_std': 15, 'sales_mean': 52000, 'sales_std': 18000, 'perf_bonus': 0.08},
    'Marketing':   {'size_weight': 0.18, 'training_mean': 40, 'training_std': 12, 'sales_mean': 38000, 'sales_std': 14000, 'perf_bonus': 0.05},
    'Engineering': {'size_weight': 0.22, 'training_mean': 50, 'training_std': 18, 'sales_mean': 12000, 'sales_std': 6000,  'perf_bonus': 0.10},
    'HR':          {'size_weight': 0.10, 'training_mean': 30, 'training_std': 10, 'sales_mean': 5000,  'sales_std': 3000,  'perf_bonus': -0.02},
    'Finance':     {'size_weight': 0.12, 'training_mean': 35, 'training_std': 11, 'sales_mean': 8000,  'sales_std': 4000,  'perf_bonus': 0.02},
    'Operations':  {'size_weight': 0.13, 'training_mean': 38, 'training_std': 13, 'sales_mean': 15000, 'sales_std': 7000,  'perf_bonus': 0.03},
}

REGIONS = {'North': 0.30, 'South': 0.28, 'East': 0.22, 'West': 0.20}

DECISION_TYPES = ['Strategic', 'Operational', 'Tactical', 'Analytical']
DECISION_WEIGHTS_BY_DEPT = {
    'Sales':       [0.15, 0.40, 0.30, 0.15],
    'Marketing':   [0.25, 0.25, 0.25, 0.25],
    'Engineering': [0.20, 0.20, 0.20, 0.40],
    'HR':          [0.30, 0.35, 0.25, 0.10],
    'Finance':     [0.25, 0.20, 0.15, 0.40],
    'Operations':  [0.10, 0.50, 0.30, 0.10],
}


def generate_timestamps(n: int) -> np.ndarray:
    """Generate non-uniform timestamps weighted toward weekdays with seasonal patterns."""
    total_days = (END_DATE - START_DATE).days + 1
    dates = [START_DATE + timedelta(days=i) for i in range(total_days)]

    weights = []
    for d in dates:
        w = 1.0
        if d.weekday() >= 5:  # weekends get less weight
            w *= 0.3
        month = d.month
        if month in (1, 2, 3, 7, 8, 9):  # Q1 & Q3 hiring spikes
            w *= 1.25
        if month == 12:  # year-end slowdown
            w *= 0.7
        weights.append(w)

    weights = np.array(weights)
    weights /= weights.sum()

    chosen_indices = np.random.choice(len(dates), size=n, p=weights)
    chosen_dates = np.array(dates)[chosen_indices]

    hours = np.random.choice(range(8, 19), size=n)
    minutes = np.random.randint(0, 60, size=n)
    seconds = np.random.randint(0, 60, size=n)

    timestamps = []
    for d, h, m, s in zip(chosen_dates, hours, minutes, seconds):
        timestamps.append(d.replace(hour=int(h), minute=int(m), second=int(s)))

    return np.array(timestamps)


def generate_employee_profiles(n_employees: int) -> pd.DataFrame:
    """Create stable employee profiles with department, region, and base experience."""
    dept_names = list(DEPARTMENTS.keys())
    dept_weights = [DEPARTMENTS[d]['size_weight'] for d in dept_names]
    dept_weights = np.array(dept_weights) / sum(dept_weights)

    region_names = list(REGIONS.keys())
    region_weights = list(REGIONS.values())

    departments = np.random.choice(dept_names, size=n_employees, p=dept_weights)
    regions = np.random.choice(region_names, size=n_employees, p=region_weights)

    base_experience = np.random.lognormal(mean=1.6, sigma=0.6, size=n_employees)
    base_experience = np.clip(base_experience, 0.5, 35).round(1)

    employee_ids = [f"EMP{str(i+1).zfill(5)}" for i in range(n_employees)]

    return pd.DataFrame({
        'Employee_ID': employee_ids,
        'Department': departments,
        'Region': regions,
        'Base_Experience': base_experience,
    })


def build_dataset() -> pd.DataFrame:
    print("Step 1/7: Generating employee profiles...")
    profiles = generate_employee_profiles(NUM_EMPLOYEES)

    print("Step 2/7: Generating timestamps and assigning employees...")
    timestamps = generate_timestamps(TARGET_ROWS)
    employee_indices = np.random.choice(NUM_EMPLOYEES, size=TARGET_ROWS, replace=True)

    df = pd.DataFrame({
        'Timestamp': timestamps,
        'Employee_ID': profiles['Employee_ID'].values[employee_indices],
        'Department': profiles['Department'].values[employee_indices],
        'Region': profiles['Region'].values[employee_indices],
        'Base_Experience': profiles['Base_Experience'].values[employee_indices],
    })
    df = df.sort_values('Timestamp').reset_index(drop=True)

    print("Step 3/7: Generating Experience_Years with time progression...")
    months_elapsed = (df['Timestamp'].dt.year - 2024) * 12 + df['Timestamp'].dt.month
    experience_growth = months_elapsed / 12.0 * np.random.uniform(0.7, 1.0, size=TARGET_ROWS)
    df['Experience_Years'] = (df['Base_Experience'] + experience_growth).round(1)
    df.drop(columns=['Base_Experience'], inplace=True)

    print("Step 4/7: Generating correlated features...")
    dept_config = df['Department'].map(DEPARTMENTS)

    training_means = df['Department'].map(lambda d: DEPARTMENTS[d]['training_mean'])
    training_stds = df['Department'].map(lambda d: DEPARTMENTS[d]['training_std'])
    df['Training_Hours'] = np.random.normal(training_means, training_stds).clip(0, 120).round(1)

    # Performance_Score: correlated with experience + training + department bonus
    exp_norm = (df['Experience_Years'] - df['Experience_Years'].mean()) / (df['Experience_Years'].std() + 1e-9)
    train_norm = (df['Training_Hours'] - df['Training_Hours'].mean()) / (df['Training_Hours'].std() + 1e-9)
    dept_bonus = df['Department'].map(lambda d: DEPARTMENTS[d]['perf_bonus'])
    noise_perf = np.random.normal(0, 1, size=TARGET_ROWS)
    raw_perf = 0.30 * exp_norm + 0.35 * train_norm + 0.15 * (dept_bonus * 10) + 0.20 * noise_perf
    df['Performance_Score'] = ((raw_perf - raw_perf.min()) / (raw_perf.max() - raw_perf.min()) * 9 + 1).round(2)

    # Monthly_Sales: department-driven, log-normal, with seasonal variation
    sales_means = df['Department'].map(lambda d: DEPARTMENTS[d]['sales_mean'])
    sales_stds = df['Department'].map(lambda d: DEPARTMENTS[d]['sales_std'])
    seasonal_factor = 1 + 0.15 * np.sin(2 * np.pi * df['Timestamp'].dt.month / 12 - np.pi / 3)
    base_sales = np.random.lognormal(
        mean=np.log(sales_means) - 0.5 * (sales_stds / sales_means) ** 2,
        sigma=sales_stds / sales_means
    )
    df['Monthly_Sales'] = (base_sales * seasonal_factor).round(2)

    # Customer_Satisfaction: correlated with performance
    perf_norm = (df['Performance_Score'] - df['Performance_Score'].mean()) / (df['Performance_Score'].std() + 1e-9)
    train_effect = (df['Training_Hours'] / df['Training_Hours'].max())
    noise_sat = np.random.normal(0, 0.3, size=TARGET_ROWS)
    raw_sat = 0.50 * perf_norm + 0.20 * train_effect + 0.30 * noise_sat
    df['Customer_Satisfaction'] = ((raw_sat - raw_sat.min()) / (raw_sat.max() - raw_sat.min()) * 4 + 1).round(2)

    # Attrition_Risk: inversely correlated with performance + satisfaction
    perf_inv = 1 - (df['Performance_Score'] - 1) / 9
    sat_inv = 1 - (df['Customer_Satisfaction'] - 1) / 4
    noise_attr = np.random.beta(2, 5, size=TARGET_ROWS)
    df['Attrition_Risk'] = (0.40 * perf_inv + 0.35 * sat_inv + 0.25 * noise_attr).clip(0, 1).round(3)

    # Decision_Type: department-weighted categorical
    decision_types = []
    for dept in df['Department']:
        w = DECISION_WEIGHTS_BY_DEPT[dept]
        decision_types.append(np.random.choice(DECISION_TYPES, p=w))
    df['Decision_Type'] = decision_types

    # Data_Quality_Score: beta distribution, department-influenced
    alpha_base = df['Department'].map({
        'Engineering': 7, 'Finance': 6, 'Sales': 4, 'Marketing': 5, 'HR': 4, 'Operations': 5
    })
    beta_base = 2
    df['Data_Quality_Score'] = np.array([
        np.random.beta(a, beta_base) for a in alpha_base
    ])
    df['Data_Quality_Score'] = (df['Data_Quality_Score'] * 100).round(1)

    # Processing_Time_sec: log-normal, somewhat correlated with data complexity
    complexity_proxy = df['Data_Quality_Score'] / 100
    df['Processing_Time_sec'] = (
        np.random.lognormal(mean=2.5, sigma=0.8, size=TARGET_ROWS) *
        (1.5 - 0.5 * complexity_proxy)
    ).round(2)

    # Error_Rate: inversely correlated with experience and data quality
    exp_norm2 = (df['Experience_Years'] - df['Experience_Years'].min()) / (df['Experience_Years'].max() - df['Experience_Years'].min() + 1e-9)
    dq_norm = df['Data_Quality_Score'] / 100
    noise_err = np.random.beta(2, 8, size=TARGET_ROWS)
    df['Error_Rate'] = (0.40 * (1 - exp_norm2) + 0.35 * (1 - dq_norm) + 0.25 * noise_err).clip(0, 1).round(4)

    print("Step 5/7: Computing Decision_Impact_Score (target)...")
    perf_scaled = (df['Performance_Score'] - 1) / 9
    train_scaled = df['Training_Hours'] / df['Training_Hours'].max()
    sat_scaled = (df['Customer_Satisfaction'] - 1) / 4
    dq_scaled = df['Data_Quality_Score'] / 100
    err_scaled = df['Error_Rate']
    noise_target = np.random.normal(0, 0.05, size=TARGET_ROWS)

    df['Decision_Impact_Score'] = (
        0.25 * dq_scaled +
        0.20 * perf_scaled -
        0.20 * err_scaled +
        0.15 * train_scaled +
        0.10 * sat_scaled +
        0.10 * noise_target
    )
    df['Decision_Impact_Score'] = (
        (df['Decision_Impact_Score'] - df['Decision_Impact_Score'].min()) /
        (df['Decision_Impact_Score'].max() - df['Decision_Impact_Score'].min()) * 100
    ).round(2)

    print("Step 6/7: Injecting realism (NULLs, outliers, inconsistencies)...")
    # Inject NULLs (5-8% in selected columns)
    null_config = {
        'Training_Hours': 0.06,
        'Customer_Satisfaction': 0.07,
        'Error_Rate': 0.05,
        'Data_Quality_Score': 0.04,
        'Processing_Time_sec': 0.03,
        'Monthly_Sales': 0.02,
    }
    for col, frac in null_config.items():
        null_mask = np.random.random(TARGET_ROWS) < frac
        df.loc[null_mask, col] = np.nan

    # Inject outliers (~0.5% of rows get extreme sales or processing times)
    outlier_idx = np.random.choice(TARGET_ROWS, size=int(TARGET_ROWS * 0.005), replace=False)
    df.loc[outlier_idx[:len(outlier_idx)//2], 'Monthly_Sales'] = np.random.uniform(200000, 500000, size=len(outlier_idx)//2).round(2)
    df.loc[outlier_idx[len(outlier_idx)//2:], 'Processing_Time_sec'] = np.random.uniform(500, 2000, size=len(outlier_idx) - len(outlier_idx)//2).round(2)

    # Minor inconsistencies: a few negative experience years
    neg_exp_idx = np.random.choice(TARGET_ROWS, size=15, replace=False)
    df.loc[neg_exp_idx, 'Experience_Years'] = np.random.uniform(-2, -0.1, size=15).round(1)

    print("Step 7/7: Finalizing and saving...")
    column_order = [
        'Timestamp', 'Employee_ID', 'Department', 'Region', 'Experience_Years',
        'Training_Hours', 'Performance_Score', 'Monthly_Sales', 'Customer_Satisfaction',
        'Attrition_Risk', 'Decision_Type', 'Data_Quality_Score', 'Processing_Time_sec',
        'Error_Rate', 'Decision_Impact_Score'
    ]
    df = df[column_order]

    return df


if __name__ == '__main__':
    print("=" * 60)
    print("  Generating Realistic Organizational Dataset")
    print("=" * 60)

    df = build_dataset()

    output_path = 'data_driven_decision_realistic.csv'
    df.to_csv(output_path, index=False)
    print(f"\nDataset saved to: {output_path}")
    print(f"Shape: {df.shape}")
    print(f"Unique Employees: {df['Employee_ID'].nunique()}")
    print(f"\nNull counts:\n{df.isnull().sum()}")
    print(f"\nSample rows:\n{df.head()}")
    print(f"\nDescriptive stats:\n{df.describe()}")
    print(f"\nCorrelation of key variables with Decision_Impact_Score:")
    numeric_cols = df.select_dtypes(include=[np.number]).columns
    corr = df[numeric_cols].corr()['Decision_Impact_Score'].sort_values(ascending=False)
    print(corr)
