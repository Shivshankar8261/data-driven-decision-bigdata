"""
Streamlit Dashboard for Data-Driven Decision Making
-----------------------------------------------------
Interactive dashboard with KPI cards, time-series, bar, scatter, and pie charts.
Run: streamlit run dashboard.py
"""

import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots

st.set_page_config(
    page_title="Data-Driven Decision Dashboard",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

st.markdown("""
<style>
    .metric-card {
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        padding: 20px; border-radius: 12px; color: white; text-align: center;
    }
    .metric-card h3 { margin: 0; font-size: 14px; opacity: 0.85; }
    .metric-card h1 { margin: 5px 0 0 0; font-size: 32px; }
    .stMetric { background-color: #f8f9fa; border-radius: 10px; padding: 10px; }
</style>
""", unsafe_allow_html=True)


@st.cache_data
def load_data():
    df = pd.read_csv("data_driven_decision_realistic.csv", parse_dates=["Timestamp"])
    df_clean = df.dropna()
    return df, df_clean


df_raw, df = load_data()

# --- SIDEBAR FILTERS ---
st.sidebar.title("🔍 Filters")
st.sidebar.markdown("---")

departments = st.sidebar.multiselect(
    "Department", options=sorted(df['Department'].unique()),
    default=sorted(df['Department'].unique())
)
regions = st.sidebar.multiselect(
    "Region", options=sorted(df['Region'].unique()),
    default=sorted(df['Region'].unique())
)
date_range = st.sidebar.date_input(
    "Date Range",
    value=(df['Timestamp'].min().date(), df['Timestamp'].max().date()),
    min_value=df['Timestamp'].min().date(),
    max_value=df['Timestamp'].max().date()
)
decision_types = st.sidebar.multiselect(
    "Decision Type", options=sorted(df['Decision_Type'].unique()),
    default=sorted(df['Decision_Type'].unique())
)

filtered = df[
    (df['Department'].isin(departments)) &
    (df['Region'].isin(regions)) &
    (df['Decision_Type'].isin(decision_types)) &
    (df['Timestamp'].dt.date >= date_range[0]) &
    (df['Timestamp'].dt.date <= date_range[1])
]

st.sidebar.markdown("---")
st.sidebar.metric("Filtered Records", f"{len(filtered):,}")
st.sidebar.metric("Total Records (Raw)", f"{len(df_raw):,}")

# --- HEADER ---
st.title("📊 Data-Driven Decision Making Dashboard")
st.markdown("*Analyzing organizational decision effectiveness across departments, regions, and time*")
st.markdown("---")

# --- KPI CARDS ---
col1, col2, col3, col4 = st.columns(4)

with col1:
    avg_impact = filtered['Decision_Impact_Score'].mean()
    st.metric("Avg Decision Impact", f"{avg_impact:.1f}", delta=f"{avg_impact - df['Decision_Impact_Score'].mean():.1f} vs overall")

with col2:
    avg_perf = filtered['Performance_Score'].mean()
    st.metric("Avg Performance", f"{avg_perf:.2f}", delta=f"{avg_perf - df['Performance_Score'].mean():.2f} vs overall")

with col3:
    avg_quality = filtered['Data_Quality_Score'].mean()
    st.metric("Avg Data Quality", f"{avg_quality:.1f}", delta=f"{avg_quality - df['Data_Quality_Score'].mean():.1f} vs overall")

with col4:
    avg_error = filtered['Error_Rate'].mean()
    st.metric("Avg Error Rate", f"{avg_error:.3f}", delta=f"{avg_error - df['Error_Rate'].mean():.3f} vs overall", delta_color="inverse")

st.markdown("---")

# --- ROW 1: Time Series + Department Bar ---
r1c1, r1c2 = st.columns(2)

with r1c1:
    st.subheader("Monthly Decision Impact Trend")
    monthly = filtered.copy()
    monthly['YearMonth'] = monthly['Timestamp'].dt.to_period('M').dt.to_timestamp()
    monthly_agg = monthly.groupby('YearMonth').agg(
        Avg_Impact=('Decision_Impact_Score', 'mean'),
        Avg_Error=('Error_Rate', 'mean')
    ).reset_index()

    fig = make_subplots(specs=[[{"secondary_y": True}]])
    fig.add_trace(
        go.Scatter(x=monthly_agg['YearMonth'], y=monthly_agg['Avg_Impact'],
                   mode='lines+markers', name='Decision Impact',
                   line=dict(color='#3498db', width=3), marker=dict(size=6)),
        secondary_y=False
    )
    fig.add_trace(
        go.Scatter(x=monthly_agg['YearMonth'], y=monthly_agg['Avg_Error'],
                   mode='lines', name='Error Rate',
                   line=dict(color='#e74c3c', width=2, dash='dash')),
        secondary_y=True
    )
    fig.update_layout(height=400, margin=dict(t=30, b=30), legend=dict(orientation="h", y=1.12))
    fig.update_yaxes(title_text="Decision Impact Score", secondary_y=False)
    fig.update_yaxes(title_text="Error Rate", secondary_y=True)
    st.plotly_chart(fig, use_container_width=True)

with r1c2:
    st.subheader("Department-wise Decision Impact")
    dept_agg = filtered.groupby('Department')['Decision_Impact_Score'].mean().sort_values(ascending=True).reset_index()
    fig = px.bar(dept_agg, x='Decision_Impact_Score', y='Department', orientation='h',
                 color='Decision_Impact_Score', color_continuous_scale='viridis',
                 text=dept_agg['Decision_Impact_Score'].round(1))
    fig.update_layout(height=400, margin=dict(t=30, b=30), showlegend=False,
                      coloraxis_showscale=False)
    fig.update_traces(textposition='outside')
    st.plotly_chart(fig, use_container_width=True)

# --- ROW 2: Scatter + Pie ---
r2c1, r2c2 = st.columns(2)

with r2c1:
    st.subheader("Data Quality vs Decision Impact")
    sample = filtered.sample(n=min(3000, len(filtered)), random_state=42)
    fig = px.scatter(sample, x='Data_Quality_Score', y='Decision_Impact_Score',
                     color='Department', opacity=0.5, size_max=8,
                     color_discrete_sequence=px.colors.qualitative.Set2)
    x_all = sample['Data_Quality_Score'].dropna().values
    y_all = sample['Decision_Impact_Score'].dropna().values
    if len(x_all) > 2:
        z = np.polyfit(x_all, y_all, 1)
        x_line = np.linspace(x_all.min(), x_all.max(), 100)
        y_line = np.polyval(z, x_line)
        fig.add_trace(go.Scatter(x=x_line, y=y_line, mode='lines',
                                 name=f'Trend (slope={z[0]:.2f})',
                                 line=dict(color='#e74c3c', width=3, dash='dash')))
    fig.update_layout(height=420, margin=dict(t=30, b=30),
                      legend=dict(orientation="h", y=-0.15))
    st.plotly_chart(fig, use_container_width=True)

with r2c2:
    st.subheader("Decision Type Distribution")
    dt_counts = filtered['Decision_Type'].value_counts().reset_index()
    dt_counts.columns = ['Decision_Type', 'Count']
    fig = px.pie(dt_counts, values='Count', names='Decision_Type',
                 color_discrete_sequence=['#3498db', '#e74c3c', '#2ecc71', '#f39c12'],
                 hole=0.45)
    fig.update_traces(textposition='outside', textinfo='percent+label')
    fig.update_layout(height=420, margin=dict(t=30, b=30), showlegend=False)
    st.plotly_chart(fig, use_container_width=True)

# --- ROW 3: Region Analysis + Error Distribution ---
r3c1, r3c2 = st.columns(2)

with r3c1:
    st.subheader("Regional Performance Overview")
    region_agg = filtered.groupby('Region').agg(
        Avg_Impact=('Decision_Impact_Score', 'mean'),
        Avg_Satisfaction=('Customer_Satisfaction', 'mean'),
        Avg_Attrition=('Attrition_Risk', 'mean'),
        Count=('Decision_Impact_Score', 'count')
    ).reset_index()
    fig = px.bar(region_agg, x='Region', y=['Avg_Impact', 'Avg_Satisfaction'],
                 barmode='group', color_discrete_sequence=['#3498db', '#2ecc71'],
                 text_auto='.1f')
    fig.update_layout(height=400, margin=dict(t=30, b=30),
                      legend=dict(orientation="h", y=1.12), yaxis_title="Score")
    st.plotly_chart(fig, use_container_width=True)

with r3c2:
    st.subheader("Error Rate Distribution by Department")
    fig = px.box(filtered, x='Department', y='Error_Rate',
                 color='Department', color_discrete_sequence=px.colors.qualitative.Set2)
    fig.add_hline(y=0.3, line_dash="dash", line_color="red",
                  annotation_text="Threshold (0.3)")
    fig.update_layout(height=400, margin=dict(t=30, b=30), showlegend=False)
    st.plotly_chart(fig, use_container_width=True)

# --- INSIGHTS SECTION ---
st.markdown("---")
st.subheader("📈 Key Business Insights")

i1, i2, i3 = st.columns(3)
with i1:
    r_val = filtered['Data_Quality_Score'].corr(filtered['Decision_Impact_Score'])
    st.info(f"**Data Quality is King**: Correlation with Decision Impact: r = {r_val:.3f}. "
            f"Investing in data quality yields the highest returns.")
with i2:
    high_t = filtered[filtered['Training_Hours'] > 40]['Performance_Score'].mean()
    low_t = filtered[filtered['Training_Hours'] <= 40]['Performance_Score'].mean()
    lift = (high_t - low_t) / low_t * 100
    st.success(f"**Training Pays Off**: Employees with >40 hrs training show {lift:.0f}% higher performance "
               f"({high_t:.1f} vs {low_t:.1f}).")
with i3:
    high_e = filtered[filtered['Error_Rate'] > 0.15]['Decision_Impact_Score'].mean()
    low_e = filtered[filtered['Error_Rate'] <= 0.15]['Decision_Impact_Score'].mean()
    drop = (low_e - high_e) / low_e * 100
    st.error(f"**Error Threshold**: Error rates >15% cause {drop:.0f}% drop in decision impact. "
             f"Implement quality gates at 15%.")

st.markdown("---")
st.caption("Dashboard powered by Streamlit | Data: 123,847 organizational records (2024-2025)")
