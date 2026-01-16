import pandas as pd
import numpy as np
import plotly.express as px

df = pd.read_excel(r"C:\Users\mamil\Desktop\ola-ride-insights\OLA_July.xlsx", engine="openpyxl")

for col in ["Booking_Value", "Customer_Rating", "Ride_Distance"]:
    if col in df.columns:
        df[col] = pd.to_numeric(df[col], errors="coerce")

print("\n===== DATASET OVERVIEW =====")
print(df.shape)
print(df.columns.tolist())
print(df.head())

print("\n===== MISSING VALUES =====")
mc = df.isnull().sum()
mp = (mc / len(df)) * 100
print(pd.DataFrame({"Missing_Count": mc, "Missing_%": mp.round(2)}))

print("\n===== DUPLICATES =====")
print(df.duplicated().sum())

print("\n===== STATS =====")
print(df.describe().T)

cols = [c for c in ["Booking_Value", "Customer_Rating", "Ride_Distance"] if c in df.columns]
for c in cols:
    q1 = df[c].quantile(0.25)
    q3 = df[c].quantile(0.75)
    iqr = q3 - q1
    lo = q1 - 1.5 * iqr
    hi = q3 + 1.5 * iqr
    print(len(df[(df[c] < lo) | (df[c] > hi)]))

if len(cols) > 1:
    print(df[cols].corr())

for c in cols:
    if not df[c].dropna().empty:
        px.box(df, x=c, title=c).show()

if "Booking_Value" in df.columns and "Ride_Distance" in df.columns:
    px.scatter(df, x="Ride_Distance", y="Booking_Value", title="Revenue").show()

if "Vehicle_Type" in df.columns:
    v = df["Vehicle_Type"].value_counts().reset_index()
    v.columns = ["Vehicle_Type", "Count"]
    px.bar(v, x="Vehicle_Type", y="Count", title="Demand").show()

print("EDA DONE")
