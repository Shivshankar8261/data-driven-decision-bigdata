"""
Report Generator v3 - Complete rebuild
-----------------------------------------
Builds report from template. Finds sections by heading text (not index).
Replaces ALL section content. Embeds 8 chart images inline.
Zero leftover template placeholder text.
"""

from docx import Document
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
import os

TEMPLATE = "BDCCT_Report_Template (1).docx"
OUTPUT = "BDCCT_Report.docx"
CHARTS = "charts"

doc = Document(TEMPLATE)

# ============================================================================
# HELPERS
# ============================================================================

def find_content_para(doc, heading_text):
    """Find the content paragraph immediately after a heading by text match."""
    for i, para in enumerate(doc.paragraphs):
        if para.style.name == 'Heading 1' and heading_text in para.text:
            if i + 1 < len(doc.paragraphs):
                return doc.paragraphs[i + 1]
    return None


def write_rich(para, parts):
    """Write mixed-format content. parts = [(text, {opts}), ...]"""
    para.clear()
    for text, opts in parts:
        run = para.add_run(text)
        run.font.size = Pt(opts.get('size', 11))
        run.font.name = 'Times New Roman'
        run.bold = opts.get('bold', False)
        run.italic = opts.get('italic', False)
        if 'color' in opts:
            run.font.color.rgb = RGBColor(*opts['color'])
    para.paragraph_format.line_spacing = 1.5
    para.paragraph_format.space_after = Pt(6)


def add_img(doc, after_elem, path, width=5.5, caption=None):
    """Insert image + caption after an element. Returns last inserted element."""
    if not os.path.exists(path):
        return after_elem
    ip = doc.add_paragraph()
    ip.alignment = WD_ALIGN_PARAGRAPH.CENTER
    ip.add_run().add_picture(path, width=Inches(width))
    ip.paragraph_format.space_before = Pt(8)
    ip.paragraph_format.space_after = Pt(2)
    after_elem.addnext(ip._element)
    if caption:
        cp = doc.add_paragraph()
        cp.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = cp.add_run(caption)
        r.font.size = Pt(9); r.font.name = 'Times New Roman'; r.italic = True
        r.font.color.rgb = RGBColor(100, 100, 100)
        cp.paragraph_format.space_after = Pt(10)
        ip._element.addnext(cp._element)
        return cp
    return ip


def add_text_block(doc, after_elem, parts):
    """Add a new paragraph with mixed formatting after an element."""
    p = doc.add_paragraph()
    for text, opts in parts:
        run = p.add_run(text)
        run.font.size = Pt(opts.get('size', 11))
        run.font.name = 'Times New Roman'
        run.bold = opts.get('bold', False)
        run.italic = opts.get('italic', False)
        if 'color' in opts:
            run.font.color.rgb = RGBColor(*opts['color'])
    p.paragraph_format.line_spacing = 1.5
    p.paragraph_format.space_after = Pt(6)
    after_elem.addnext(p._element)
    return p


# ============================================================================
# TITLE PAGE
# ============================================================================
doc.paragraphs[4].clear()
r = doc.paragraphs[4].add_run("Project Title: Data-Driven Decision Making in an Organization Using Big Data Technologies")
r.font.size = Pt(12); r.font.name = 'Times New Roman'; r.bold = True

# ============================================================================
# PHASE 1: Replace ALL section text content FIRST (before inserting images)
#           This avoids index shifting problems.
# ============================================================================

# --- SECTION 1: Problem Statement ---
para = find_content_para(doc, '1. Problem Statement')
write_rich(para, [
    ("Domain & Context\n", {'bold': True, 'size': 12, 'color': (44, 62, 80)}),
    ("\nIn the contemporary business landscape, organizations generate vast volumes of data across departments, "
     "yet a significant majority struggle to translate this data into actionable, high-quality decisions. "
     "According to a 2024 survey by NewVantage Partners, only 26.5% of organizations have successfully "
     "established a data-driven culture, despite 97.2% investing in big data and AI initiatives. "
     "This gap between data availability and decision effectiveness constitutes a critical operational challenge.\n", {}),
    ("\nThe domain selected for this project is organizational performance analytics, focusing on how data quality, "
     "employee competency, and processing efficiency collectively influence the impact of business decisions.\n", {}),
    ("\nProblem Definition\n", {'bold': True, 'size': 12, 'color': (44, 62, 80)}),
    ("\nOrganizations lack an integrated, scalable pipeline that ingests raw operational data, cleanses and transforms it, "
     "and produces actionable insights that directly improve decision-making quality. Poor decision-making\u2014driven by "
     "low data quality, high error rates, and fragmented analytics\u2014directly translates to revenue loss, "
     "employee attrition, and competitive disadvantage.\n", {}),
    ("\nTarget Users\n", {'bold': True, 'size': 12, 'color': (44, 62, 80)}),
    ("\n\u2022 C-suite executives \u2014 dashboards to monitor decision effectiveness across departments and regions\n", {}),
    ("\u2022 Department heads \u2014 insights into team performance, training ROI, and error rate trends\n", {}),
    ("\u2022 Data analysts \u2014 clean, transformed data in Parquet format for advanced analytics and ML\n", {}),
    ("\u2022 HR managers \u2014 attrition risk patterns and their correlation with satisfaction metrics\n", {}),
    ("\nBusiness Significance\n", {'bold': True, 'size': 12, 'color': (44, 62, 80)}),
    ("\nMcKinsey estimates data-driven organizations are 23\u00d7 more likely to acquire customers and 19\u00d7 more likely to be profitable. "
     "This project demonstrates how a robust big data pipeline can systematically improve decision-making capability.\n", {}),
    ("\nDataset Overview\n", {'bold': True, 'size': 12, 'color': (44, 62, 80)}),
    ("\nThe dataset comprises 123,847 records spanning 15 variables across 6 departments (Sales, Marketing, Engineering, "
     "HR, Finance, Operations), 4 regions (North, South, East, West), and a 2-year period (Jan 2024 \u2013 Dec 2025). "
     "Includes correlated variables, 5\u20138% missing values, 0.5% outliers, and seasonal patterns. "
     "Target variable: Decision Impact Score (composite of data quality, performance, error rate, training, satisfaction).", {}),
])

# --- SECTION 2: Pipeline Design ---
para = find_content_para(doc, '2. Pipeline Design')
write_rich(para, [
    ("The pipeline follows a five-stage architecture transforming raw data into actionable insights.\n", {}),
    ("\nStage 1 \u2014 Data Source & Ingestion\n", {'bold': True, 'color': (74, 144, 217)}),
    ("Raw dataset (123,847 rows \u00d7 15 columns) stored as CSV. In production, data resides in GCP Cloud Storage for automated ingestion.\n", {}),
    ("\nStage 2 \u2014 Data Processing (PySpark)\n", {'bold': True, 'color': (232, 132, 60)}),
    ("Apache Spark as core engine: (1) distributed computing for horizontal scaling, (2) Catalyst optimizer for query optimization, "
     "(3) DataFrame API for structured data. Processing: schema enforcement, null imputation, dedup, outlier capping, feature engineering.\n", {}),
    ("\nStage 3 \u2014 Data Transformation\n", {'bold': True, 'color': (232, 132, 60)}),
    ("Feature engineering: experience_training_ratio, efficiency_score, quality_error_interaction, temporal features (month, quarter, day_of_week).\n", {}),
    ("\nStage 4 \u2014 Data Storage (Parquet)\n", {'bold': True, 'color': (155, 89, 182)}),
    ("Parquet format, partitioned by Department. Benefits: columnar I/O reduction, Snappy compression (60\u201380%), predicate pushdown.\n", {}),
    ("\nStage 5 \u2014 Visualization (Streamlit)\n", {'bold': True, 'color': (39, 174, 96)}),
    ("Interactive dashboard with KPI cards, time-series, bar charts, scatter plots, pie charts. Deployed on Streamlit Cloud.\n", {}),
    ("\nDesign Justification\n", {'bold': True, 'size': 12, 'color': (44, 62, 80)}),
    ("\nReproducibility (coded pipelines), scalability (Spark + GCP), maintainability (modular scripts), open-source (no vendor lock-in).", {}),
])

# --- SECTION 3: System Architecture ---
para = find_content_para(doc, '3. System Architecture')
write_rich(para, [
    ("The system follows a layered design where each component has a single responsibility. The diagram below illustrates the complete data flow.", {}),
])

# --- SECTION 4: PySpark Implementation ---
para = find_content_para(doc, '4. PySpark Implementation')
write_rich(para, [
    ("Dataset Description\n", {'bold': True, 'size': 12, 'color': (44, 62, 80)}),
    ("\n123,847 records \u00d7 15 columns. 6 departments, 4 regions, ~2,500 employees. Jan 2024 \u2013 Dec 2025. "
     "Includes 5\u20138% NULLs, 0.5% outliers, negative experience values, department transfers.\n", {}),
    ("\nStep 1 \u2014 SparkSession & Loading\n", {'bold': True, 'color': (232, 132, 60)}),
    ("SparkSession with 4GB memory, AQE enabled. Loaded with explicit StructType schema:\n", {}),
    ("\n    spark = SparkSession.builder.appName(\"Pipeline\")\n", {'size': 10, 'color': (41, 128, 185)}),
    ("        .config(\"spark.driver.memory\", \"4g\").getOrCreate()\n", {'size': 10, 'color': (41, 128, 185)}),
    ("    schema = StructType([\n", {'size': 10, 'color': (41, 128, 185)}),
    ("        StructField(\"Timestamp\", TimestampType(), True),\n", {'size': 10, 'color': (41, 128, 185)}),
    ("        StructField(\"Employee_ID\", StringType(), True), ...\n", {'size': 10, 'color': (41, 128, 185)}),
    ("        StructField(\"Decision_Impact_Score\", DoubleType(), True)])\n", {'size': 10, 'color': (41, 128, 185)}),
    ("    df = spark.read.csv(path, header=True, schema=schema)\n", {'size': 10, 'color': (41, 128, 185)}),
    ("\nStep 2 \u2014 Data Cleaning\n", {'bold': True, 'color': (232, 132, 60)}),
    ("(a) dropDuplicates() for exact duplicate removal. (b) abs() for negative Experience_Years. "
     "(c) Median imputation for numeric NULLs (robust to outliers). "
     "(d) IQR capping (Q1\u22121.5\u00d7IQR, Q3+1.5\u00d7IQR) for Monthly_Sales, Processing_Time.\n", {}),
    ("\nStep 3 \u2014 Feature Engineering\n", {'bold': True, 'color': (232, 132, 60)}),
    ("\u2022 Experience_Training_Ratio = Experience / (Training + 1)\n", {}),
    ("\u2022 Efficiency_Score = Performance / (Processing_Time + 1)\n", {}),
    ("\u2022 Quality_Error_Interaction = Data_Quality \u00d7 (1 \u2212 Error_Rate)\n", {}),
    ("\u2022 Temporal: Month, Quarter, Day_of_Week, Year from Timestamp\n", {}),
    ("\u2022 Binning: Performance_Category (High/Medium/Low), Risk_Level (Critical/Moderate/Low)\n", {}),
    ("\nStep 4 \u2014 Aggregations\n", {'bold': True, 'color': (232, 132, 60)}),
    ("Department-wise, Region-wise, and Quarterly aggregations of key metrics.\n", {}),
    ("\nStep 5 \u2014 Filtering & Output\n", {'bold': True, 'color': (232, 132, 60)}),
    ("High error (>0.3) and high attrition (\u22650.7) cases isolated. Output: 24 columns, Parquet, partitioned by Department.", {}),
])

# --- SECTION 5: Dashboard & Visualization ---
para = find_content_para(doc, '5. Dashboard & Visualization')
write_rich(para, [
    ("Tool Selection\n", {'bold': True, 'size': 12, 'color': (44, 62, 80)}),
    ("\nStreamlit selected for: (1) native Python integration, (2) zero licensing cost, "
     "(3) interactive widgets, (4) Streamlit Cloud deployment. "
     "Live at: https://data-driven-decision-bigdata.streamlit.app\n", {}),
    ("\nDashboard Components\n", {'bold': True, 'size': 12, 'color': (44, 62, 80)}),
    ("\n1. KPI Metric Cards: ", {'bold': True}), ("Avg Decision Impact, Performance, Data Quality, Error Rate with deltas.\n", {}),
    ("2. Monthly Trend (Line Chart): ", {'bold': True}), ("Dual-axis: Decision Impact vs Error Rate over 24 months.\n", {}),
    ("3. Dept Impact (Bar Chart): ", {'bold': True}), ("Horizontal bar ranking departments by avg Decision Impact.\n", {}),
    ("4. Quality vs Impact (Scatter): ", {'bold': True}), ("Strong r=0.736 correlation with regression trendline.\n", {}),
    ("5. Decision Types (Donut): ", {'bold': True}), ("Strategic/Operational/Tactical/Analytical distribution.\n", {}),
    ("6. Regional Overview (Grouped Bar): ", {'bold': True}), ("Impact and Satisfaction across 4 regions.\n", {}),
    ("7. Error Distribution (Box Plot): ", {'bold': True}), ("Per-department Error Rate with 0.3 threshold line.", {}),
])

# --- SECTION 6: Practical Application & Scalability ---
para = find_content_para(doc, '6. Practical Application & Scalability')
write_rich(para, [
    ("Real-World Use Cases\n", {'bold': True, 'size': 12, 'color': (44, 62, 80)}),
    ("\n1. Management Consulting: ", {'bold': True}),
    ("Analyze client operational data, identify decision quality bottlenecks, prescribe training/quality interventions.\n", {}),
    ("2. Healthcare: ", {'bold': True}),
    ("Clinical decision error monitoring, staff training compliance, nursing attrition prediction.\n", {}),
    ("3. Retail: ", {'bold': True}),
    ("POS and customer data analysis for regional decision optimization and resource reallocation.\n", {}),
    ("4. Finance: ", {'bold': True}),
    ("Quality-gate credit/fraud decisions, error rate compliance for regulatory requirements.\n", {}),
    ("\nScalability\n", {'bold': True, 'size': 12, 'color': (44, 62, 80)}),
    ("\n\u2022 Horizontal: GCP Dataproc clusters (10\u2013100 nodes) for petabyte-scale data\n", {}),
    ("\u2022 Partitioning: Composite (Dept+Region+Year) for efficient pruning\n", {}),
    ("\u2022 Storage: Parquet+Snappy (60\u201380% compression), Delta Lake for ACID\n", {}),
    ("\u2022 Performance: Broadcast joins, df.cache(), AQE optimization\n", {}),
    ("\nFuture Enhancements\n", {'bold': True, 'size': 12, 'color': (44, 62, 80)}),
    ("\n\u2022 Spark Structured Streaming for real-time dashboards\n", {}),
    ("\u2022 ML models (Random Forest, XGBoost) to predict Decision Impact\n", {}),
    ("\u2022 Automated alerts (PagerDuty/Slack) for threshold breaches\n", {}),
    ("\u2022 Data lineage via Apache Atlas or GCP Data Catalog", {}),
])

# --- SECTION 7: Data-Driven Decision Making ---
para = find_content_para(doc, '7. Data-Driven Decision Making')
write_rich(para, [
    ("The analysis enables five concrete, data-backed decisions:\n", {}),
    ("\nDecision 1 \u2014 Training Budget Allocation\n", {'bold': True, 'color': (39, 174, 96)}),
    ("Training >40 hrs \u2192 27% higher performance. Increase budget 20% for HR/Finance/Operations. Expected: 15\u201320% improvement in 6 months.\n", {}),
    ("\nDecision 2 \u2014 Data Quality Monitoring\n", {'bold': True, 'color': (39, 174, 96)}),
    ("Strongest predictor (r=0.736). Three-tier: Green(\u226575), Yellow(60\u201374), Red(<60). Red \u2192 manual review. Expected: 30% fewer low-quality decisions.\n", {}),
    ("\nDecision 3 \u2014 Error Rate Quality Gates\n", {'bold': True, 'color': (39, 174, 96)}),
    ("Error >15% \u2192 43% impact drop. Hard gate at 15%: auto-pause + review. Expected: eliminate lowest-quality 20%.\n", {}),
    ("\nDecision 4 \u2014 Attrition Risk Mitigation\n", {'bold': True, 'color': (39, 174, 96)}),
    ("Attrition\u2194Satisfaction r=\u22120.875. Monthly surveys for high-risk (\u22650.7), targeted retention. Expected: 25% reduction in satisfaction decline.\n", {}),
    ("\nDecision 5 \u2014 Regional Reallocation\n", {'bold': True, 'color': (39, 174, 96)}),
    ("Regional disparities detected. Reallocate resources, standardize best practices. Expected: convergence within 10%.\n", {}),
    ("\nImpact Summary: ", {'bold': True, 'size': 12, 'color': (44, 62, 80)}),
    ("Projected 18\u201325% improvement in Decision Impact over 12 months, translating to 9\u201312% operational efficiency gain.", {}),
])

# --- SECTION 8: Conclusion ---
para = find_content_para(doc, '8. Conclusion')
write_rich(para, [
    ("This project designed, implemented, and demonstrated a complete big data pipeline for data-driven decision making.\n", {}),
    ("\nProblem: ", {'bold': True}), ("Organizations cannot convert operational data into quality decisions due to data quality issues, fragmented pipelines, and lack of insights.\n", {}),
    ("\nApproach: ", {'bold': True}), ("Five-stage pipeline: (1) 123,847-record dataset, (2) PySpark ETL, (3) Parquet storage, (4) Streamlit dashboard, (5) Six insights + five decisions.\n", {}),
    ("\nKey Results:\n", {'bold': True, 'size': 12, 'color': (44, 62, 80)}),
    ("\u2022 Data Quality: strongest predictor (r=0.736)\n", {}),
    ("\u2022 Training >40 hrs: 27% higher performance\n", {}),
    ("\u2022 Error threshold at 15%: 43% quality degradation beyond\n", {}),
    ("\u2022 Engineering leads Impact (best quality + training)\n", {}),
    ("\u2022 Attrition\u2194Satisfaction feedback loop (r=\u22120.875)\n", {}),
    ("\nModern big data technologies, in a well-architected pipeline, transform raw data into actionable intelligence for better decisions.", {}),
])

# --- SECTION 9: References ---
para = find_content_para(doc, '9. References')
write_rich(para, [
    ("[1] Zaharia, M., et al. (2016). Apache Spark: A Unified Engine for Big Data Processing. CACM, 59(11), 56\u201365.\n\n", {}),
    ("[2] Apache Spark Documentation. (2025). https://spark.apache.org/docs/latest/\n\n", {}),
    ("[3] Google Cloud. (2025). Dataproc. https://cloud.google.com/dataproc/docs\n\n", {}),
    ("[4] Apache Parquet. (2025). https://parquet.apache.org/documentation/latest/\n\n", {}),
    ("[5] Streamlit. (2025). https://docs.streamlit.io/\n\n", {}),
    ("[6] NewVantage Partners. (2024). Data & Analytics Leadership Survey.\n\n", {}),
    ("[7] McKinsey. (2023). The Age of Analytics.\n\n", {}),
    ("[8] Provost & Fawcett. (2013). Data Science for Business. O'Reilly.\n\n", {}),
    ("[9] Kimball & Ross. (2013). The Data Warehouse Toolkit. Wiley.\n\n", {}),
    ("[10] Google Cloud. (2025). BigQuery. https://cloud.google.com/bigquery/docs", {}),
])

# --- SECTION 10: Appendix ---
para = find_content_para(doc, '10. Appendix')
write_rich(para, [
    ("Project Repository\n", {'bold': True, 'size': 12, 'color': (44, 62, 80)}),
    ("\nGitHub: https://github.com/Shivshankar8261/data-driven-decision-bigdata\n", {}),
    ("\nKey Files:\n", {'bold': True}),
    ("\u2022 generate_dataset.py \u2014 Dataset generation (123,847 rows, correlated, NULLs, outliers)\n", {}),
    ("\u2022 pyspark_pipeline.py \u2014 PySpark ETL (clean, transform, aggregate, Parquet)\n", {}),
    ("\u2022 pyspark_pipeline.ipynb \u2014 Jupyter notebook with step-by-step processing\n", {}),
    ("\u2022 analysis.py \u2014 7 charts + 6 business insights\n", {}),
    ("\u2022 dashboard.py \u2014 Streamlit dashboard (7 chart types, filters)\n", {}),
    ("\u2022 data_driven_decision_realistic.csv \u2014 Generated dataset\n", {}),
    ("\u2022 charts/ \u2014 Visualization images (PNG)\n", {}),
    ("\u2022 processed_data/ \u2014 Parquet output partitioned by Department", {}),
])


# ============================================================================
# PHASE 2: Insert images AFTER all text has been replaced
# ============================================================================

# Re-find paragraphs (indices are stable now since Phase 1 only replaced text)
sec3_para = find_content_para(doc, '3. System Architecture')
sec4_para = find_content_para(doc, '4. PySpark Implementation')
sec5_para = find_content_para(doc, '5. Dashboard & Visualization')
sec6_para = find_content_para(doc, '6. Practical Application')

# Section 3: Architecture diagram
last = add_img(doc, sec3_para._element, os.path.join(CHARTS, '00_system_architecture.png'),
               width=5.8, caption='Figure 1: System Architecture \u2014 Data-Driven Decision Pipeline')

comp_parts = [
    ("1. Data Source: ", True), ("CSV, databases, APIs \u2192 GCP Cloud Storage.\n", False),
    ("2. Spark Engine: ", True), ("GCP Dataproc, StructType schema, cleaning, FE, aggregations.\n", False),
    ("3. Storage: ", True), ("Parquet on GCS, partitioned by Department, 11-nines durability.\n", False),
    ("4. Visualization: ", True), ("Streamlit dashboard, KPI cards, 7 chart types, sidebar filters.\n", False),
    ("5. Flow: ", True), ("Sources \u2192 Processing \u2192 Storage \u2192 Visualization \u2192 Decisions.", False),
]
cp = doc.add_paragraph()
for text, bold in comp_parts:
    r = cp.add_run(text)
    r.font.size = Pt(11); r.font.name = 'Times New Roman'; r.bold = bold
cp.paragraph_format.line_spacing = 1.5
last._element.addnext(cp._element)

# Section 4: Correlation heatmap
add_img(doc, sec4_para._element, os.path.join(CHARTS, '01_correlation_heatmap.png'),
        width=5.0, caption='Figure 2: Correlation Matrix of Key Organizational Metrics')

# Section 5: Multiple charts
last = add_img(doc, sec5_para._element, os.path.join(CHARTS, '02_dept_decision_impact.png'),
               width=5.0, caption='Figure 3: Department-wise Average Decision Impact Score')
last = add_img(doc, last._element, os.path.join(CHARTS, '03_quality_vs_impact_scatter.png'),
               width=5.0, caption='Figure 4: Data Quality vs Decision Impact (r = 0.736)')
last = add_img(doc, last._element, os.path.join(CHARTS, '04_monthly_trend.png'),
               width=5.2, caption='Figure 5: Monthly Decision Impact & Error Rate (2024\u20132025)')
last = add_img(doc, last._element, os.path.join(CHARTS, '06_decision_type_pie.png'),
               width=4.0, caption='Figure 6: Decision Type Distribution')

# Insights block after section 5 charts
ins = doc.add_paragraph()
ins_parts = [
    ("\nKey Business Insights\n\n", {'bold': True, 'size': 12, 'color': (44, 62, 80)}),
    ("Insight 1: ", {'bold': True}), ("Data Quality is the strongest predictor of Decision Impact (r=0.736).\n\n", {}),
    ("Insight 2: ", {'bold': True}), ("Training >40 hrs/yr yields 27% higher performance (4.76 vs 3.76).\n\n", {}),
    ("Insight 3: ", {'bold': True}), ("Error rates >15% cause 43% decline in Decision Impact (77.7\u219244.6).\n\n", {}),
    ("Insight 4: ", {'bold': True}), ("Engineering leads Impact (+40% vs HR) due to best quality + training.\n\n", {}),
    ("Insight 5: ", {'bold': True}), ("Attrition Risk \u2194 Satisfaction: r=\u22120.875 feedback loop.\n\n", {}),
    ("Insight 6: ", {'bold': True}), ("Quarterly improvement Q1\u2192Q4: organizational learning compounds over fiscal years.", {}),
]
for text, opts in ins_parts:
    r = ins.add_run(text)
    r.font.size = Pt(opts.get('size', 11)); r.font.name = 'Times New Roman'; r.bold = opts.get('bold', False)
    if 'color' in opts: r.font.color.rgb = RGBColor(*opts['color'])
ins.paragraph_format.line_spacing = 1.5
last._element.addnext(ins._element)

# Section 6: Training violin + Error boxplot
last = add_img(doc, sec6_para._element, os.path.join(CHARTS, '07_training_performance_violin.png'),
               width=5.2, caption='Figure 7: Training Investment vs Performance by Department')
add_img(doc, last._element, os.path.join(CHARTS, '05_error_rate_boxplot.png'),
        width=5.0, caption='Figure 8: Error Rate Distribution by Department')

# ============================================================================
# SAVE & VERIFY
# ============================================================================
doc.save(OUTPUT)

v = Document(OUTPUT)
bad = sum(1 for p in v.paragraphs if any(x in p.text for x in ['Insert architecture', 'Insert code snippets', 'Insert dashboard', 'Mention tool used']))
print(f"Report saved to: {OUTPUT}")
print(f"Total paragraphs: {len(v.paragraphs)}")
print(f"Template placeholders remaining: {bad}")
if bad == 0:
    print("ALL sections properly populated.")
