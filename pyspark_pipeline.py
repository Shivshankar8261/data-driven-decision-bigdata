"""
PySpark Data Pipeline for Data-Driven Decision Making
------------------------------------------------------
Complete ETL pipeline: Load → Profile → Clean → Transform → Aggregate → Save
"""

from pyspark.sql import SparkSession
from pyspark.sql.types import (
    StructType, StructField, TimestampType, StringType,
    DoubleType, IntegerType
)
from pyspark.sql.functions import (
    col, when, count, isnan, isnull, mean as spark_mean, median as spark_median,
    abs as spark_abs, month, quarter, dayofweek, year,
    round as spark_round, lit, percentile_approx, desc, avg, sum as spark_sum,
    max as spark_max, min as spark_min, countDistinct
)
from pyspark.sql.window import Window
import os

# ============================================================================
# 1. SPARK SESSION SETUP
# ============================================================================
spark = SparkSession.builder \
    .appName("DataDrivenDecisionPipeline") \
    .config("spark.driver.memory", "4g") \
    .config("spark.sql.legacy.timeParserPolicy", "LEGACY") \
    .config("spark.sql.adaptive.enabled", "true") \
    .getOrCreate()

spark.sparkContext.setLogLevel("WARN")
print("=" * 70)
print("  PySpark Data Pipeline - Data-Driven Decision Making")
print("=" * 70)

# ============================================================================
# 2. DEFINE SCHEMA & LOAD DATA
# ============================================================================
schema = StructType([
    StructField("Timestamp",             TimestampType(), True),
    StructField("Employee_ID",           StringType(),    True),
    StructField("Department",            StringType(),    True),
    StructField("Region",                StringType(),    True),
    StructField("Experience_Years",      DoubleType(),    True),
    StructField("Training_Hours",        DoubleType(),    True),
    StructField("Performance_Score",     DoubleType(),    True),
    StructField("Monthly_Sales",         DoubleType(),    True),
    StructField("Customer_Satisfaction", DoubleType(),    True),
    StructField("Attrition_Risk",        DoubleType(),    True),
    StructField("Decision_Type",         StringType(),    True),
    StructField("Data_Quality_Score",    DoubleType(),    True),
    StructField("Processing_Time_sec",   DoubleType(),    True),
    StructField("Error_Rate",            DoubleType(),    True),
    StructField("Decision_Impact_Score", DoubleType(),    True),
])

DATA_PATH = "data_driven_decision_realistic.csv"
OUTPUT_PATH = "processed_data"

print("\n[LOAD] Reading dataset with explicit schema...")
df = spark.read.csv(DATA_PATH, header=True, schema=schema, timestampFormat="yyyy-MM-dd HH:mm:ss")
print(f"  Rows loaded: {df.count()}")
print(f"  Columns: {len(df.columns)}")
df.printSchema()

# ============================================================================
# 3. DATA PROFILING
# ============================================================================
print("\n[PROFILE] Dataset overview")
print("-" * 50)

total_rows = df.count()
print(f"  Total rows: {total_rows}")
print(f"  Total columns: {len(df.columns)}")

print("\n  Null counts per column:")
null_counts = df.select([
    count(when(isnull(c) | isnan(c), c)).alias(c)
    for c in df.columns if df.schema[c].dataType != StringType()
] + [
    count(when(isnull(c), c)).alias(c)
    for c in df.columns if df.schema[c].dataType == StringType()
])
null_counts.show(truncate=False)

print("  Descriptive statistics:")
df.describe().show()

# ============================================================================
# 4. DATA CLEANING
# ============================================================================
print("\n[CLEAN] Step 4.1: Removing exact duplicates...")
before_dedup = df.count()
df = df.dropDuplicates()
after_dedup = df.count()
print(f"  Removed {before_dedup - after_dedup} duplicate rows")

print("[CLEAN] Step 4.2: Fixing negative Experience_Years...")
neg_count = df.filter(col("Experience_Years") < 0).count()
df = df.withColumn("Experience_Years",
                    when(col("Experience_Years") < 0, spark_abs(col("Experience_Years")))
                    .otherwise(col("Experience_Years")))
print(f"  Fixed {neg_count} negative experience values")

print("[CLEAN] Step 4.3: Imputing NULL values...")
numeric_cols_to_impute = [
    "Training_Hours", "Monthly_Sales", "Customer_Satisfaction",
    "Data_Quality_Score", "Processing_Time_sec", "Error_Rate"
]
for c in numeric_cols_to_impute:
    median_val = df.select(spark_median(col(c))).collect()[0][0]
    null_count = df.filter(isnull(col(c))).count()
    if median_val is not None:
        df = df.withColumn(c, when(isnull(col(c)), lit(median_val)).otherwise(col(c)))
        print(f"    {c}: filled {null_count} NULLs with median = {median_val:.2f}")

print("[CLEAN] Step 4.4: Capping outliers using IQR method...")
outlier_cols = ["Monthly_Sales", "Processing_Time_sec"]
for c in outlier_cols:
    quantiles = df.approxQuantile(c, [0.25, 0.75], 0.01)
    q1, q3 = quantiles[0], quantiles[1]
    iqr = q3 - q1
    lower = q1 - 1.5 * iqr
    upper = q3 + 1.5 * iqr
    outlier_count = df.filter((col(c) < lower) | (col(c) > upper)).count()
    df = df.withColumn(c,
                        when(col(c) < lower, lit(lower))
                        .when(col(c) > upper, lit(upper))
                        .otherwise(col(c)))
    print(f"    {c}: capped {outlier_count} outliers (IQR bounds: [{lower:.2f}, {upper:.2f}])")

print(f"\n  Cleaned dataset rows: {df.count()}")

# ============================================================================
# 5. FEATURE ENGINEERING
# ============================================================================
print("\n[TRANSFORM] Feature Engineering...")

df = df.withColumn("Experience_Training_Ratio",
                    spark_round(col("Experience_Years") / (col("Training_Hours") + 1), 4))

df = df.withColumn("Efficiency_Score",
                    spark_round(col("Performance_Score") / (col("Processing_Time_sec") + 1), 4))

df = df.withColumn("Quality_Error_Interaction",
                    spark_round(col("Data_Quality_Score") * (1 - col("Error_Rate")), 2))

df = df.withColumn("Month", month(col("Timestamp")))
df = df.withColumn("Quarter", quarter(col("Timestamp")))
df = df.withColumn("Day_of_Week", dayofweek(col("Timestamp")))
df = df.withColumn("Year", year(col("Timestamp")))

df = df.withColumn("Performance_Category",
                    when(col("Performance_Score") >= 7.5, "High")
                    .when(col("Performance_Score") >= 4.5, "Medium")
                    .otherwise("Low"))

df = df.withColumn("Risk_Level",
                    when(col("Attrition_Risk") >= 0.7, "Critical")
                    .when(col("Attrition_Risk") >= 0.4, "Moderate")
                    .otherwise("Low"))

print("  New features added:")
new_cols = ["Experience_Training_Ratio", "Efficiency_Score", "Quality_Error_Interaction",
            "Month", "Quarter", "Day_of_Week", "Year", "Performance_Category", "Risk_Level"]
for c in new_cols:
    print(f"    - {c}")

# ============================================================================
# 6. AGGREGATIONS
# ============================================================================
print("\n[AGGREGATE] Department-wise Analysis:")
dept_agg = df.groupBy("Department").agg(
    spark_round(avg("Decision_Impact_Score"), 2).alias("Avg_Decision_Impact"),
    spark_round(avg("Performance_Score"), 2).alias("Avg_Performance"),
    spark_round(avg("Monthly_Sales"), 2).alias("Avg_Sales"),
    spark_round(avg("Data_Quality_Score"), 2).alias("Avg_Data_Quality"),
    spark_round(avg("Error_Rate"), 4).alias("Avg_Error_Rate"),
    countDistinct("Employee_ID").alias("Unique_Employees"),
    count("*").alias("Record_Count")
).orderBy(desc("Avg_Decision_Impact"))
dept_agg.show(truncate=False)

print("[AGGREGATE] Region-wise Analysis:")
region_agg = df.groupBy("Region").agg(
    spark_round(avg("Customer_Satisfaction"), 2).alias("Avg_Satisfaction"),
    spark_round(avg("Attrition_Risk"), 3).alias("Avg_Attrition_Risk"),
    spark_round(avg("Decision_Impact_Score"), 2).alias("Avg_Decision_Impact"),
    count("*").alias("Record_Count")
).orderBy(desc("Avg_Decision_Impact"))
region_agg.show(truncate=False)

print("[AGGREGATE] Quarterly Trend Analysis:")
quarterly_agg = df.groupBy("Year", "Quarter").agg(
    spark_round(avg("Decision_Impact_Score"), 2).alias("Avg_Impact"),
    spark_round(avg("Performance_Score"), 2).alias("Avg_Performance"),
    spark_round(avg("Error_Rate"), 4).alias("Avg_Error_Rate"),
    count("*").alias("Record_Count")
).orderBy("Year", "Quarter")
quarterly_agg.show(truncate=False)

# ============================================================================
# 7. FILTERING
# ============================================================================
print("[FILTER] High Error Rate Cases (Error_Rate > 0.3):")
high_error = df.filter(col("Error_Rate") > 0.3)
print(f"  Records with high error rate: {high_error.count()} ({high_error.count()/df.count()*100:.1f}%)")
high_error_dept = high_error.groupBy("Department").agg(
    count("*").alias("High_Error_Count"),
    spark_round(avg("Decision_Impact_Score"), 2).alias("Avg_Impact_When_High_Error")
).orderBy(desc("High_Error_Count"))
high_error_dept.show(truncate=False)

print("[FILTER] High Attrition Risk Employees (Risk >= 0.7):")
high_risk = df.filter(col("Attrition_Risk") >= 0.7)
print(f"  High-risk records: {high_risk.count()} ({high_risk.count()/df.count()*100:.1f}%)")
high_risk_summary = high_risk.groupBy("Department").agg(
    countDistinct("Employee_ID").alias("At_Risk_Employees"),
    spark_round(avg("Performance_Score"), 2).alias("Avg_Performance"),
    spark_round(avg("Customer_Satisfaction"), 2).alias("Avg_Satisfaction")
).orderBy(desc("At_Risk_Employees"))
high_risk_summary.show(truncate=False)

# ============================================================================
# 8. SAVE PROCESSED DATA TO PARQUET
# ============================================================================
print(f"\n[SAVE] Writing processed data to Parquet (partitioned by Department)...")
df.write.mode("overwrite") \
    .partitionBy("Department") \
    .parquet(OUTPUT_PATH)
print(f"  Saved to: {OUTPUT_PATH}/")

verification = spark.read.parquet(OUTPUT_PATH)
print(f"  Verification - Rows in Parquet: {verification.count()}")
print(f"  Verification - Columns: {len(verification.columns)}")

# ============================================================================
# SUMMARY
# ============================================================================
print("\n" + "=" * 70)
print("  PIPELINE SUMMARY")
print("=" * 70)
print(f"  Input rows:          {total_rows}")
print(f"  Duplicates removed:  {before_dedup - after_dedup}")
print(f"  NULLs imputed:       {sum(df.filter(isnull(c)).count() == 0 for c in numeric_cols_to_impute)} columns cleaned")
print(f"  Features added:      {len(new_cols)}")
print(f"  Output format:       Parquet (partitioned by Department)")
print(f"  Output location:     {OUTPUT_PATH}/")
print("=" * 70)

spark.stop()
print("\nSpark session stopped. Pipeline complete.")
