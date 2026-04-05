"""
PPT Generator - Populates the BDCCT PPT Template
--------------------------------------------------
Fills all 6 slides following the template structure.
"""

from pptx import Presentation
from pptx.util import Pt, Emu
from pptx.dml.color import RGBColor

TEMPLATE = "BDCCT_PPT_Template (1).pptx"
OUTPUT = "BDCCT_PPT.pptx"

prs = Presentation(TEMPLATE)


def set_text_frame(text_frame, lines, font_size=14, bold_first=False):
    """Replace all text in a text frame with new lines, preserving formatting."""
    text_frame.clear()
    for i, line in enumerate(lines):
        if i == 0:
            para = text_frame.paragraphs[0]
        else:
            para = text_frame.add_paragraph()

        is_bullet = line.strip().startswith('•')
        if is_bullet:
            para.level = 1

        run = para.add_run()
        run.text = line
        run.font.size = Pt(font_size)

        if bold_first and i == 0:
            run.font.bold = True

        if i == 0 and not bold_first:
            run.font.size = Pt(font_size)


# ============================================================================
# SLIDE 1: Project Title & Problem Context
# ============================================================================
slide1 = prs.slides[0]
title1 = slide1.placeholders[0]
content1 = slide1.placeholders[1]

title1.text_frame.paragraphs[0].text = "Project Title & Problem Context"
for run in title1.text_frame.paragraphs[0].runs:
    run.font.size = Pt(28)

set_text_frame(content1.text_frame, [
    "Project Title: Data-Driven Decision Making in an Organization Using Big Data Technologies",
    "Name & USN: ____________________ | ____________________",
    "Domain: Big Data Analytics & Organizational Decision Intelligence",
    "Problem Statement: Organizations generate vast operational data but lack integrated pipelines to convert it into high-quality decisions. Poor data quality, high error rates, and fragmented analytics reduce decision effectiveness by up to 40%.",
    "Target Users: C-suite executives, Department heads, Data analysts, HR managers",
], font_size=14, bold_first=True)

# ============================================================================
# SLIDE 2: System Architecture
# ============================================================================
slide2 = prs.slides[1]
title2 = slide2.placeholders[0]
content2 = slide2.placeholders[1]

set_text_frame(content2.text_frame, [
    "Data Source (CSV/ERP)  →  Ingestion (GCP Cloud Storage)  →  Processing (PySpark on GCP Dataproc)  →  Storage (Parquet)  →  Dashboard (Streamlit)",
    "",
    "Tools & Technologies:",
    "• Apache Spark (PySpark) — Distributed data processing engine",
    "• GCP Dataproc — Managed Spark cluster with auto-scaling",
    "• GCP Cloud Storage — Durable raw & processed data lake",
    "• Apache Parquet — Columnar storage with Snappy compression",
    "• Streamlit — Interactive Python dashboard framework",
    "• Plotly / Matplotlib / Seaborn — Visualization libraries",
    "",
    "Workflow: Raw CSV → Schema Validation → Cleaning (Nulls, Duplicates, Outliers) → Feature Engineering (9 derived features) → Aggregations → Parquet Storage → Interactive Dashboard → Business Insights",
], font_size=13)

# ============================================================================
# SLIDE 3: PySpark Implementation
# ============================================================================
slide3 = prs.slides[2]
title3 = slide3.placeholders[0]
content3 = slide3.placeholders[1]

set_text_frame(content3.text_frame, [
    "Dataset: data_driven_decision_realistic.csv — 123,847 rows × 15 columns (2024–2025)",
    "",
    "Key Operations:",
    "• Data Cleaning: Null imputation (median), duplicate removal, outlier capping (IQR method), negative value correction",
    "• Feature Engineering: experience_training_ratio, efficiency_score, quality_error_interaction, temporal features (month, quarter, day_of_week)",
    "• Aggregations: Department-wise (avg impact, performance, sales), Region-wise (satisfaction, attrition), Quarterly trends",
    "• Filtering: High error rate cases (>0.3), High attrition risk employees (≥0.7)",
    "",
    "Code Snippet:",
    "  df = spark.read.csv(path, header=True, schema=schema)",
    "  df = df.dropDuplicates()",
    "  df = df.withColumn('Efficiency_Score', col('Performance_Score')/(col('Processing_Time_sec')+1))",
    "  df.write.partitionBy('Department').parquet('processed_data/')",
    "",
    "Output: 24 columns (9 engineered), Parquet format, partitioned by Department",
], font_size=12)

# ============================================================================
# SLIDE 4: Dashboard & Insights
# ============================================================================
slide4 = prs.slides[3]
title4 = slide4.placeholders[0]
content4 = slide4.placeholders[1]

set_text_frame(content4.text_frame, [
    "Dashboard: Streamlit interactive dashboard with KPI cards, 6 chart types, sidebar filters",
    "",
    "Key Insight 1: Data Quality Score is the strongest predictor of Decision Impact (r = 0.736). Organizations should prioritize data quality programs to maximize decision effectiveness.",
    "",
    "Key Insight 2: Employees with >40 training hours/year achieve 27% higher performance scores (4.76 vs 3.76). Clear ROI on training investment.",
    "",
    "Key Insight 3: Error rates above 15% cause a 43% drop in Decision Impact Score (77.7 → 44.6). Implement automated quality gates at this threshold.",
    "",
    "Additional: Engineering department leads in Decision Impact (+40% vs HR) due to superior data quality and training investment. Attrition Risk and Customer Satisfaction show strong inverse correlation (r = -0.875).",
], font_size=13)

# ============================================================================
# SLIDE 5: Real-World Application
# ============================================================================
slide5 = prs.slides[4]
title5 = slide5.placeholders[0]
content5 = slide5.placeholders[1]

set_text_frame(content5.text_frame, [
    "Real-World Applications:",
    "• Management Consulting — Analyze client decision quality, identify bottlenecks, prescribe interventions",
    "• Healthcare Administration — Monitor clinical decision errors, track staff training compliance, predict attrition",
    "• Retail & E-Commerce — Optimize pricing/marketing decisions across regions and product categories",
    "• Financial Services — Quality-gate credit decisions, flag high-error-rate processing pipelines",
    "",
    "Scalability:",
    "• Horizontal Scaling: GCP Dataproc clusters (10–100 nodes) for petabyte-scale data",
    "• Partition Strategy: Composite partitioning (Department + Region + Year) for efficient queries",
    "• Storage: Parquet + Snappy compression (60–80% reduction), Delta Lake for ACID transactions",
    "",
    "Future Improvements:",
    "• Real-time processing via Spark Structured Streaming",
    "• ML models (Random Forest, XGBoost) to predict Decision Impact proactively",
    "• Automated alerts when error rates exceed thresholds",
], font_size=12)

# ============================================================================
# SLIDE 6: Conclusion & Questions
# ============================================================================
slide6 = prs.slides[5]
title6 = slide6.placeholders[0]
content6 = slide6.placeholders[1]

set_text_frame(content6.text_frame, [
    "Summary:",
    "",
    "Problem: Organizations fail to leverage data for high-quality decisions due to poor data quality, fragmented pipelines, and lack of actionable insights.",
    "",
    "Solution: Built a complete big data pipeline using PySpark, GCP, and Streamlit that ingests 123,847 records, cleans and transforms the data, and delivers interactive dashboards with business insights.",
    "",
    "Outcome: Identified Data Quality (r=0.74) and Error Rate (r=-0.77) as the top drivers of decision effectiveness. Derived 5 actionable decisions projected to improve decision quality by 18–25%.",
    "",
    "",
    "Thank You",
], font_size=14)

# ============================================================================
# SAVE
# ============================================================================
prs.save(OUTPUT)
print(f"Presentation saved to: {OUTPUT}")
print(f"Total slides: {len(prs.slides)}")
