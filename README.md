# 🚀 AI-Powered Excel Analytics Engine

A Python-based automation system that transforms raw Excel data into actionable business insights, reports, and KPIs — all with one command.

---

## 📌 Overview

This project combines **Excel + Python + Automation** to:

* 📊 Process raw business data
* 📈 Generate KPIs and trends
* 🧠 Produce structured analytics
* 📁 Export reports automatically

---

## ⚙️ Features

* ✅ Excel data ingestion (Raw_Data sheet)
* ✅ Data cleaning & preprocessing
* ✅ KPI calculation (Total & Average Revenue)
* ✅ Monthly trend analysis
* ✅ Automatic Excel Summary updates
* ✅ CSV report generation
* ✅ One-click automation via `.bat` file

---

## 🧱 Project Structure

```
AI_Analytics_Project/
│
├── analytics_engine.py      # Main Python script
├── Book1.xlsx              # Excel data file
├── requirements.txt        # Dependencies
├── run_windows.bat         # One-click runner
└── output/                 # Generated reports
```

---

## 📦 Installation

### 1. Install Python

Download from: https://www.python.org/

👉 IMPORTANT: Enable **"Add Python to PATH"**

---

### 2. Install Dependencies

```bash
pip install -r requirements.txt
```

OR manually:

```bash
pip install pandas numpy openpyxl
```

---

## 📊 Excel Setup

Your Excel file must contain **3 sheets**:

### 1. Raw_Data

| Date       | Revenue | Product | Region |
| ---------- | ------- | ------- | ------ |
| 2025-01-01 | 1000    | A       | India  |

---

### 2. Mapping_Config

| Field   | Column_Name |
| ------- | ----------- |
| Date    | Date        |
| Revenue | Revenue     |
| Product | Product     |
| Region  | Region      |

---

### 3. Summary

👉 Leave empty (auto-filled by Python)

---

## ▶️ How to Run

### Basic Run

```bash
python analytics_engine.py
```

---

### Generate Reports

```bash
python analytics_engine.py --report
```

---

### One-Click Run (Windows)

Double-click:

```
run_windows.bat
```

---

## 📁 Output Files

Generated inside `/output` folder:

* 📄 `monthly_trend_YYYYMMDD.csv`
* 📊 KPI updates in Excel (Summary sheet)

---

## 🧠 How It Works

1. Reads Excel data
2. Cleans & processes values
3. Calculates KPIs
4. Groups data by month
5. Writes results back to Excel
6. (Optional) Saves CSV reports

---




## 🧪 Testing

✔ Modify data in `Raw_Data`
✔ Run script again
✔ Verify updated results in `Summary`

---

## 🚀 Future Improvements

* 🔍 Anomaly detection (Z-score)
* 🤖 AI-generated insights (OpenAI integration)
* 📧 Email automation (SMTP)
* 🌐 Web dashboard (FastAPI)
* 📊 Interactive visualizations

---

## 💼 Use Cases

* Sales analytics
* Financial reporting
* Business dashboards
* Startup automation tools
* Data analyst portfolios

---

## 👨‍💻 Author

Built by an AI/ML Engineer krishna focused on automation, analytics, and scalable systems.


---

## 📌 License

Free to use for learning and portfolio purposes.
