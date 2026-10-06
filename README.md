# Aviation Market Analysis Automation Tool

An automated, Streamlit-based Minimum Viable Product (MVP) designed for aviation market analysts and data engineers. This tool processes, cleans, and analyzes airline industry data (OAG, Cirium, and IATA) to generate standardized reports, pivot tables, and visualizations.

## 🚀 Features

### 1. Robust Data Ingestion & Validation
* **File Uploads:** Supports `.xlsx` file uploads for three core data sources:
  * OAG Data
  * Cirium Trend Data
  * IATA Data
* **Automated Quality Checks:** Automatically scans the first sheet of uploaded data for missing values (nulls). If missing data is detected, the process halts and alerts the user exactly which column requires cleaning before proceeding.
* **Data Preview:** Successfully uploaded and cleaned data is cached in the background and can be previewed via a collapsible section showing the first 5 rows of the standardized wide table.

### 2. Core Analytical Modules (ETL & BI)
Users can select from four specific analysis modules via a dropdown menu. The UI dynamically generates the required input fields based on the selected module.

* **Module 1: C909 Range Coverage Destination Airports**
  * **Inputs:** `Carrier Name`, `Dep Airport Code`
  * **Logic:** Filters OAG data for the specified airline and departure airport. Extracts destination airport codes (`Arr Airport Code`) where the flight frequency is greater than the median. Outputs the result as an Excel file (useful for range map generation).

* **Module 2: Airline Fleet Size Evolution (2015-2025)**
  * **Input:** `Operator`
  * **Logic:** Filters Cirium data for active and stored aircraft (`Total in Service` & `Total in Storage`). Creates a pivot table by category (`CAT`) and years. Generates a stacked bar chart styled in **COMAC Blue** (PANTONE Reflex Blue C/U Process C100 M70 Y0 K0 R0 G78 B162) and exports the data to Excel.

* **Module 3: Route Distribution based on Daily One-way Passenger Volume**
  * **Input:** `Carrier Name`
  * **Logic:** Merges IATA and OAG data. Calculates new metrics:
    * *Daily One-way Passenger Volume* = `Pax Count` / 365
    * *Passenger Volume Groups* = '>400', '100-400', '50-100', '0-50'
    * *Fare per Pax Kilometer* = `Fare` / `Pax Count` / `Distance`
  * Joins the origin-destination data (`Orig-Dest` / `Dep-Arr`). Outputs a pivot table grouping routes by passenger volume, displaying average fares, average distance, and total frequency.

* **Module 4: National Fleet Size Evolution by Airline (1990-2026)**
  * **Inputs:** `Operator Country/Subregion`, `Time Range` (e.g., 2000-2025)
  * **Logic:** Filters Cirium data for a specific country/region and dynamically subsets the data based on the provided time range. Outputs a pivot table of fleet sizes by airline (`Operator`) and generates a styled line chart (COMAC Blue theme with distinct line colors for different operators).

## 🛠️ Tech Stack
* **Python 3.x**
* **Streamlit:** Interactive UI and web application framework
* **Pandas:** Advanced data manipulation, ETL processes, and aggregations
* **Matplotlib / Plotly / Seaborn:** Data visualization (Charts and graphs)
* **Openpyxl:** Excel file reading and writing

## 💻 Installation and Setup

1. **Clone the repository:**
   ```bash
   git clone https://github.com/yourusername/aviation-market-analysis.git
   cd aviation-market-analysis
   ```

2. **Install dependencies:**
   Ensure you have Python installed, then run:
   ```bash
   pip install pandas streamlit openpyxl matplotlib
   ```

3. **Run the application:**
   The tool is designed as a single-file MVP. Run it directly using Streamlit:
   ```bash
   streamlit run app.py
   ```
   The application will open automatically in your default web browser in wide-page layout.

## 📂 Data Format Requirements
To ensure the application runs smoothly, ensure your uploaded `.xlsx` files contain the required columns:
* **OAG Data:** `Carrier Name`, `Dep Airport Code`, `Arr Airport Code`, `Frequency`, `Time Series`
* **Cirium Data:** `CAT`, `Type`, `Operator`, `Manufacturer`, `Metric`, `Operator Country/Subregion`, and Year columns (1990-2026).
* **IATA Data:** `Seg Orig`, `Seg Dest`, `Fare`, `Pax Count`, `Distance`

## 📄 License
[MIT License](LICENSE)
