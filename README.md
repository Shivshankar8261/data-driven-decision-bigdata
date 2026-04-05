# Data-Driven Decision Making in an Organization Using Big Data Technologies

**Capstone Project — Big Data & Cloud Computing Technologies**

## Overview

This project implements a complete big data pipeline for analyzing organizational decision-making effectiveness. It processes 123,847+ records across 6 departments and 4 regions, using PySpark for ETL and Streamlit for interactive visualization.

## Project Structure

| File | Description |
|------|-------------|
| `generate_dataset.py` | Generates 123,847 rows of realistic organizational data |
| `data_driven_decision_realistic.csv` | The generated dataset (15 columns) |
| `pyspark_pipeline.py` | PySpark ETL pipeline (clean, transform, aggregate, save as Parquet) |
| `analysis.py` | Statistical analysis + 7 publication-quality charts |
| `dashboard.py` | Streamlit interactive dashboard |
| `charts/` | Generated visualization PNGs |
| `BDCCT_Report.docx` | Full academic report |
| `BDCCT_PPT.pptx` | 6-slide presentation |

## Key Findings

- **Data Quality Score** is the strongest predictor of Decision Impact (r = 0.74)
- **Error rates >15%** cause a 43% drop in decision effectiveness
- **Training >40 hrs/year** yields 27% higher employee performance
- **Engineering** leads in Decision Impact, outperforming HR by ~40%

## Tech Stack

- **Processing**: Apache Spark (PySpark), Python
- **Storage**: Apache Parquet (columnar, compressed)
- **Cloud**: GCP Dataproc, Cloud Storage
- **Visualization**: Streamlit, Plotly, Matplotlib, Seaborn

## Run the Dashboard

```bash
pip install -r requirements.txt
streamlit run dashboard.py
```

## Run the PySpark Pipeline

```bash
pip install pyspark
python pyspark_pipeline.py
```
