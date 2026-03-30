import pandas as pd
from pathlib import Path
import argparse
from datetime import datetime

EXCEL_FILE = "Book1.xlsx"
OUTPUT_DIR = Path("output")

def load_data():
    df = pd.read_excel(EXCEL_FILE, sheet_name="Raw_Data")
    return df

def process_data(df):
    df['Date'] = pd.to_datetime(df['Date'])
    df['Revenue'] = pd.to_numeric(df['Revenue'], errors='coerce')
    return df

def generate_kpis(df):
    total_revenue = df['Revenue'].sum()
    avg_revenue = df['Revenue'].mean()

    return {
        "Total Revenue": total_revenue,
        "Average Revenue": avg_revenue
    }

def monthly_trend(df):
    df['Month'] = df['Date'].dt.to_period('M')
    trend = df.groupby('Month')['Revenue'].sum().reset_index()
    return trend

def save_outputs(df, trend):
    timestamp = datetime.now().strftime("%Y%m%d_%H%M")

    OUTPUT_DIR.mkdir(exist_ok=True)

    trend.to_csv(OUTPUT_DIR / f"monthly_trend_{timestamp}.csv", index=False)

def update_excel(kpis):
    summary_df = pd.DataFrame(list(kpis.items()), columns=["Metric", "Value"])

    with pd.ExcelWriter(EXCEL_FILE, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
        summary_df.to_excel(writer, sheet_name="Summary", index=False)

def main(report=False):
    print("Running analytics...")

    df = load_data()
    df = process_data(df)

    kpis = generate_kpis(df)
    trend = monthly_trend(df)

    update_excel(kpis)

    if report:
        save_outputs(df, trend)

    print("Done!")

if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("--report", action="store_true")

    args = parser.parse_args()

    main(report=args.report)