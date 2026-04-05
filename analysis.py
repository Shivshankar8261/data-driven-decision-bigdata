"""
Data Analysis & Visualization
------------------------------
Generates 7 publication-quality charts and derives 6 business insights
from the organizational decision-making dataset.
"""

import pandas as pd
import numpy as np
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import matplotlib.ticker as mticker
import seaborn as sns
import os
import warnings
warnings.filterwarnings('ignore')

sns.set_theme(style="whitegrid", font_scale=1.1)
plt.rcParams.update({
    'figure.dpi': 150,
    'savefig.bbox': 'tight',
    'savefig.pad_inches': 0.3,
    'font.family': 'sans-serif',
})

CHARTS_DIR = "charts"
os.makedirs(CHARTS_DIR, exist_ok=True)

print("Loading dataset...")
df = pd.read_csv("data_driven_decision_realistic.csv", parse_dates=["Timestamp"])
print(f"Dataset shape: {df.shape}")

df_clean = df.dropna()
print(f"Clean rows (no NaN): {len(df_clean)}")

# ============================================================================
# CHART 1: Correlation Heatmap
# ============================================================================
print("\nGenerating Chart 1: Correlation Heatmap...")
numeric_cols = [
    'Experience_Years', 'Training_Hours', 'Performance_Score', 'Monthly_Sales',
    'Customer_Satisfaction', 'Attrition_Risk', 'Data_Quality_Score',
    'Processing_Time_sec', 'Error_Rate', 'Decision_Impact_Score'
]
corr_matrix = df_clean[numeric_cols].corr()

fig, ax = plt.subplots(figsize=(12, 9))
mask = np.triu(np.ones_like(corr_matrix, dtype=bool))
cmap = sns.diverging_palette(220, 20, as_cmap=True)
sns.heatmap(corr_matrix, mask=mask, cmap=cmap, center=0, annot=True, fmt='.2f',
            square=True, linewidths=0.8, ax=ax,
            cbar_kws={"shrink": 0.8, "label": "Correlation Coefficient"})
ax.set_title("Correlation Matrix of Key Organizational Metrics", fontsize=14, fontweight='bold', pad=15)
plt.savefig(os.path.join(CHARTS_DIR, "01_correlation_heatmap.png"))
plt.close()

# ============================================================================
# CHART 2: Department-wise Average Decision Impact Score
# ============================================================================
print("Generating Chart 2: Department-wise Decision Impact...")
dept_impact = df_clean.groupby('Department')['Decision_Impact_Score'].mean().sort_values(ascending=True)

fig, ax = plt.subplots(figsize=(10, 6))
colors = sns.color_palette("viridis", len(dept_impact))
bars = ax.barh(dept_impact.index, dept_impact.values, color=colors, edgecolor='white', height=0.6)
for bar, val in zip(bars, dept_impact.values):
    ax.text(val + 0.3, bar.get_y() + bar.get_height()/2, f'{val:.1f}',
            va='center', fontweight='bold', fontsize=11)
ax.set_xlabel("Average Decision Impact Score", fontsize=12)
ax.set_title("Department-wise Average Decision Impact Score", fontsize=14, fontweight='bold')
ax.set_xlim(0, dept_impact.max() * 1.12)
ax.spines['top'].set_visible(False)
ax.spines['right'].set_visible(False)
plt.savefig(os.path.join(CHARTS_DIR, "02_dept_decision_impact.png"))
plt.close()

# ============================================================================
# CHART 3: Data Quality Score vs Decision Impact Score (Scatter + Regression)
# ============================================================================
print("Generating Chart 3: Data Quality vs Decision Impact...")
sample = df_clean.sample(n=min(5000, len(df_clean)), random_state=42)

fig, ax = plt.subplots(figsize=(10, 7))
scatter = ax.scatter(sample['Data_Quality_Score'], sample['Decision_Impact_Score'],
                     c=sample['Error_Rate'], cmap='RdYlGn_r', alpha=0.5, s=15, edgecolors='none')
z = np.polyfit(sample['Data_Quality_Score'], sample['Decision_Impact_Score'], 1)
p = np.poly1d(z)
x_line = np.linspace(sample['Data_Quality_Score'].min(), sample['Data_Quality_Score'].max(), 100)
ax.plot(x_line, p(x_line), color='#e74c3c', linewidth=2.5, linestyle='--', label=f'Trend (slope={z[0]:.2f})')
cbar = plt.colorbar(scatter, ax=ax, shrink=0.8)
cbar.set_label("Error Rate", fontsize=11)
r_val = sample['Data_Quality_Score'].corr(sample['Decision_Impact_Score'])
ax.text(0.05, 0.95, f'r = {r_val:.3f}', transform=ax.transAxes,
        fontsize=13, fontweight='bold', va='top',
        bbox=dict(boxstyle='round,pad=0.3', facecolor='wheat', alpha=0.8))
ax.set_xlabel("Data Quality Score", fontsize=12)
ax.set_ylabel("Decision Impact Score", fontsize=12)
ax.set_title("Data Quality Score vs Decision Impact Score", fontsize=14, fontweight='bold')
ax.legend(loc='lower right', fontsize=11)
plt.savefig(os.path.join(CHARTS_DIR, "03_quality_vs_impact_scatter.png"))
plt.close()

# ============================================================================
# CHART 4: Monthly Trend of Decision Impact (Time Series)
# ============================================================================
print("Generating Chart 4: Monthly Decision Impact Trend...")
df_clean['YearMonth'] = df_clean['Timestamp'].dt.to_period('M')
monthly_trend = df_clean.groupby('YearMonth').agg(
    Avg_Impact=('Decision_Impact_Score', 'mean'),
    Avg_Error=('Error_Rate', 'mean'),
    Count=('Decision_Impact_Score', 'count')
).reset_index()
monthly_trend['YearMonth'] = monthly_trend['YearMonth'].dt.to_timestamp()

fig, ax1 = plt.subplots(figsize=(14, 6))
ax1.plot(monthly_trend['YearMonth'], monthly_trend['Avg_Impact'],
         color='#2c3e50', linewidth=2.5, marker='o', markersize=5, label='Avg Decision Impact')
ax1.fill_between(monthly_trend['YearMonth'], monthly_trend['Avg_Impact'],
                 alpha=0.15, color='#3498db')
ax1.set_xlabel("Month", fontsize=12)
ax1.set_ylabel("Avg Decision Impact Score", fontsize=12, color='#2c3e50')
ax1.tick_params(axis='y', labelcolor='#2c3e50')

ax2 = ax1.twinx()
ax2.plot(monthly_trend['YearMonth'], monthly_trend['Avg_Error'],
         color='#e74c3c', linewidth=2, linestyle='--', marker='s', markersize=4, label='Avg Error Rate')
ax2.set_ylabel("Avg Error Rate", fontsize=12, color='#e74c3c')
ax2.tick_params(axis='y', labelcolor='#e74c3c')

lines1, labels1 = ax1.get_legend_handles_labels()
lines2, labels2 = ax2.get_legend_handles_labels()
ax1.legend(lines1 + lines2, labels1 + labels2, loc='upper left', fontsize=10)
ax1.set_title("Monthly Trend: Decision Impact Score & Error Rate (2024-2025)",
              fontsize=14, fontweight='bold')
fig.autofmt_xdate()
plt.savefig(os.path.join(CHARTS_DIR, "04_monthly_trend.png"))
plt.close()

# ============================================================================
# CHART 5: Error Rate Distribution by Department (Box Plot)
# ============================================================================
print("Generating Chart 5: Error Rate by Department...")
dept_order = df_clean.groupby('Department')['Error_Rate'].median().sort_values().index

fig, ax = plt.subplots(figsize=(10, 6))
bp = sns.boxplot(data=df_clean, x='Department', y='Error_Rate', order=dept_order,
                 palette='Set2', fliersize=2, linewidth=1.2, ax=ax)
ax.axhline(y=0.3, color='red', linestyle='--', linewidth=1.5, alpha=0.7, label='Threshold (0.3)')
ax.set_xlabel("Department", fontsize=12)
ax.set_ylabel("Error Rate", fontsize=12)
ax.set_title("Error Rate Distribution by Department", fontsize=14, fontweight='bold')
ax.legend(fontsize=11)
plt.savefig(os.path.join(CHARTS_DIR, "05_error_rate_boxplot.png"))
plt.close()

# ============================================================================
# CHART 6: Decision Type Distribution (Pie Chart)
# ============================================================================
print("Generating Chart 6: Decision Type Distribution...")
decision_counts = df['Decision_Type'].value_counts()

fig, ax = plt.subplots(figsize=(8, 8))
colors_pie = ['#3498db', '#e74c3c', '#2ecc71', '#f39c12']
wedges, texts, autotexts = ax.pie(
    decision_counts.values, labels=decision_counts.index,
    autopct='%1.1f%%', colors=colors_pie, startangle=140,
    pctdistance=0.75, wedgeprops=dict(width=0.5, edgecolor='white', linewidth=2)
)
for t in autotexts:
    t.set_fontsize(12)
    t.set_fontweight('bold')
ax.set_title("Distribution of Decision Types Across Organization",
             fontsize=14, fontweight='bold', pad=20)
plt.savefig(os.path.join(CHARTS_DIR, "06_decision_type_pie.png"))
plt.close()

# ============================================================================
# CHART 7: Training Hours vs Performance Score by Department (Violin)
# ============================================================================
print("Generating Chart 7: Training vs Performance (Violin)...")
fig, axes = plt.subplots(1, 2, figsize=(16, 6))

sns.violinplot(data=df_clean, x='Department', y='Training_Hours', palette='muted',
               inner='quartile', ax=axes[0])
axes[0].set_title("Training Hours Distribution by Department", fontsize=13, fontweight='bold')
axes[0].set_xlabel("Department", fontsize=11)
axes[0].set_ylabel("Training Hours", fontsize=11)
axes[0].tick_params(axis='x', rotation=30)

sns.violinplot(data=df_clean, x='Department', y='Performance_Score', palette='muted',
               inner='quartile', ax=axes[1])
axes[1].set_title("Performance Score Distribution by Department", fontsize=13, fontweight='bold')
axes[1].set_xlabel("Department", fontsize=11)
axes[1].set_ylabel("Performance Score", fontsize=11)
axes[1].tick_params(axis='x', rotation=30)

plt.suptitle("Training Investment vs Performance Outcome", fontsize=15, fontweight='bold', y=1.02)
plt.tight_layout()
plt.savefig(os.path.join(CHARTS_DIR, "07_training_performance_violin.png"))
plt.close()

# ============================================================================
# BUSINESS INSIGHTS
# ============================================================================
print("\n" + "=" * 70)
print("  KEY BUSINESS INSIGHTS")
print("=" * 70)

# Insight 1: Department impact comparison
dept_avg = df_clean.groupby('Department')['Decision_Impact_Score'].mean()
top_dept = dept_avg.idxmax()
bot_dept = dept_avg.idxmin()
pct_diff = ((dept_avg[top_dept] - dept_avg[bot_dept]) / dept_avg[bot_dept] * 100)
print(f"\n1. DEPARTMENT PERFORMANCE GAP:")
print(f"   {top_dept} department achieves {pct_diff:.1f}% higher Decision Impact than {bot_dept}.")
print(f"   Top: {dept_avg[top_dept]:.1f} vs Bottom: {dept_avg[bot_dept]:.1f}")

# Insight 2: Training impact
high_train = df_clean[df_clean['Training_Hours'] > 40]['Performance_Score'].mean()
low_train = df_clean[df_clean['Training_Hours'] <= 40]['Performance_Score'].mean()
train_lift = ((high_train - low_train) / low_train * 100)
print(f"\n2. TRAINING ROI:")
print(f"   Employees with >40 training hours achieve {train_lift:.1f}% higher performance.")
print(f"   High training: {high_train:.2f} vs Low training: {low_train:.2f}")

# Insight 3: Error rate threshold
high_err = df_clean[df_clean['Error_Rate'] > 0.15]['Decision_Impact_Score'].mean()
low_err = df_clean[df_clean['Error_Rate'] <= 0.15]['Decision_Impact_Score'].mean()
err_drop = ((low_err - high_err) / low_err * 100)
print(f"\n3. ERROR RATE THRESHOLD EFFECT:")
print(f"   Error rates above 15% cause a {err_drop:.1f}% drop in Decision Impact.")
print(f"   Low error impact: {low_err:.1f} vs High error impact: {high_err:.1f}")

# Insight 4: Data quality as top predictor
r_dq = df_clean['Data_Quality_Score'].corr(df_clean['Decision_Impact_Score'])
r_perf = df_clean['Performance_Score'].corr(df_clean['Decision_Impact_Score'])
print(f"\n4. DATA QUALITY IS THE STRONGEST PREDICTOR:")
print(f"   Data Quality → Decision Impact correlation: r = {r_dq:.3f}")
print(f"   Performance → Decision Impact correlation: r = {r_perf:.3f}")

# Insight 5: Attrition feedback loop
r_attr_sat = df_clean['Attrition_Risk'].corr(df_clean['Customer_Satisfaction'])
high_risk = df_clean[df_clean['Attrition_Risk'] > 0.6]
print(f"\n5. ATTRITION-SATISFACTION FEEDBACK LOOP:")
print(f"   Attrition Risk ↔ Customer Satisfaction correlation: r = {r_attr_sat:.3f}")
print(f"   High-risk employees ({len(high_risk)} records) avg satisfaction: {high_risk['Customer_Satisfaction'].mean():.2f}")

# Insight 6: Quarterly patterns
quarterly = df_clean.copy()
quarterly['Quarter'] = quarterly['Timestamp'].dt.quarter
q_means = quarterly.groupby('Quarter')['Decision_Impact_Score'].mean()
best_q = q_means.idxmax()
worst_q = q_means.idxmin()
print(f"\n6. SEASONAL DECISION EFFECTIVENESS:")
print(f"   Best quarter: Q{best_q} (avg impact: {q_means[best_q]:.1f})")
print(f"   Weakest quarter: Q{worst_q} (avg impact: {q_means[worst_q]:.1f})")
print(f"   Quarterly scores: {', '.join(f'Q{q}={v:.1f}' for q, v in q_means.items())}")

print(f"\n{'=' * 70}")
print(f"  All charts saved to: {CHARTS_DIR}/")
print(f"{'=' * 70}")
