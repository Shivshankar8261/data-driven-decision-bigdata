"""
Generate a professional system architecture diagram.
"""

import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
from matplotlib.patches import FancyBboxPatch, FancyArrowPatch
import numpy as np

fig, ax = plt.subplots(figsize=(14, 10))
ax.set_xlim(0, 14)
ax.set_ylim(0, 10)
ax.axis('off')
fig.patch.set_facecolor('white')

COLORS = {
    'source':  '#4A90D9',
    'ingest':  '#50B5A9',
    'process': '#E8843C',
    'storage': '#9B59B6',
    'viz':     '#27AE60',
    'decide':  '#2C3E50',
    'arrow':   '#7F8C8D',
    'bg':      '#F8F9FA',
}

def draw_box(ax, x, y, w, h, color, title, items, icon=''):
    box = FancyBboxPatch((x, y), w, h, boxstyle="round,pad=0.15",
                          facecolor=color, edgecolor='white', linewidth=2.5, alpha=0.92)
    ax.add_patch(box)
    title_text = f"{icon}  {title}" if icon else title
    ax.text(x + w/2, y + h - 0.32, title_text, ha='center', va='top',
            fontsize=12, fontweight='bold', color='white', family='sans-serif')
    for i, item in enumerate(items):
        ax.text(x + w/2, y + h - 0.72 - i*0.32, item, ha='center', va='top',
                fontsize=8.5, color='#FFFFFFDD', family='sans-serif')

def draw_arrow(ax, x1, y1, x2, y2, label=''):
    ax.annotate('', xy=(x2, y2), xytext=(x1, y1),
                arrowprops=dict(arrowstyle='->', color=COLORS['arrow'],
                                lw=2.5, connectionstyle='arc3,rad=0'))
    if label:
        mx, my = (x1+x2)/2, (y1+y2)/2
        ax.text(mx + 0.15, my, label, fontsize=7.5, color=COLORS['arrow'],
                fontstyle='italic', va='center')

# Title
ax.text(7, 9.6, 'System Architecture — Data-Driven Decision Pipeline',
        ha='center', va='center', fontsize=16, fontweight='bold',
        color=COLORS['decide'], family='sans-serif')
ax.plot([1.5, 12.5], [9.35, 9.35], color='#BDC3C7', linewidth=1.5)

# Layer 1: Data Sources
draw_box(ax, 0.5, 7.2, 3.8, 1.8, COLORS['source'],
         'DATA SOURCES', ['CSV Files (ERP Exports)', 'Databases (PostgreSQL, MySQL)', 'APIs (CRM, HRMS Systems)'])

draw_box(ax, 5.1, 7.2, 3.8, 1.8, COLORS['ingest'],
         'DATA INGESTION', ['GCP Cloud Storage (Data Lake)', 'Python Validation Scripts', 'Schema Enforcement'])

draw_box(ax, 9.7, 7.2, 3.8, 1.8, COLORS['process'],
         'PROCESSING ENGINE', ['Apache Spark (PySpark)', 'GCP Dataproc Cluster', 'Cleaning | FE | Aggregation'])

# Layer 2: Storage + Visualization
draw_box(ax, 1.5, 3.8, 4.5, 2.2, COLORS['storage'],
         'STORAGE LAYER', ['Apache Parquet (Columnar)', 'Partitioned by Department', 'Snappy Compression (60-80%)', 'GCP Cloud Storage / BigQuery'])

draw_box(ax, 8.0, 3.8, 4.5, 2.2, COLORS['viz'],
         'VISUALIZATION LAYER', ['Streamlit Dashboard (Interactive)', 'Plotly / Matplotlib / Seaborn', 'KPI Cards | Charts | Filters', 'Deployed on Streamlit Cloud'])

# Layer 3: Decision Making
draw_box(ax, 3.5, 0.8, 7.0, 2.0, COLORS['decide'],
         'BUSINESS DECISION LAYER', ['Training Budget Allocation', 'Data Quality Monitoring & Thresholds',
                                      'Error Rate Quality Gates', 'Attrition Risk Mitigation'])

# Arrows
draw_arrow(ax, 4.3, 8.1, 5.1, 8.1, 'Raw Data')
draw_arrow(ax, 8.9, 8.1, 9.7, 8.1, 'Validated')
draw_arrow(ax, 11.6, 7.2, 10.25, 6.0, 'Processed')
draw_arrow(ax, 3.75, 7.2, 3.75, 6.0, 'Raw Store')
draw_arrow(ax, 6.0, 4.9, 8.0, 4.9, 'Read')
draw_arrow(ax, 3.75, 3.8, 5.5, 2.8, 'Query')
draw_arrow(ax, 10.25, 3.8, 8.5, 2.8, 'Insights')

# Tool badges at bottom
tools = ['PySpark', 'GCP Dataproc', 'Cloud Storage', 'Parquet', 'Streamlit', 'Plotly', 'Python']
badge_y = 0.15
for i, tool in enumerate(tools):
    bx = 1.0 + i * 1.75
    badge = FancyBboxPatch((bx, badge_y), 1.5, 0.4, boxstyle="round,pad=0.08",
                            facecolor='#ECF0F1', edgecolor='#BDC3C7', linewidth=1)
    ax.add_patch(badge)
    ax.text(bx + 0.75, badge_y + 0.2, tool, ha='center', va='center',
            fontsize=7, color='#2C3E50', fontweight='bold')

plt.tight_layout()
plt.savefig('charts/00_system_architecture.png', dpi=200, bbox_inches='tight',
            facecolor='white', edgecolor='none')
plt.close()
print("Architecture diagram saved to charts/00_system_architecture.png")
