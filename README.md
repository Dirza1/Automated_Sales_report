# Excel Cleaner & KPI Report Generator

## 📌 Project Description
This project demonstrates how to **automatically clean incoming Excel data** and **generate KPI reports** using Python and the `openpyxl` library.  
It is designed as a lightweight alternative to manual Excel cleanup and pivot tables, fully automated with logging and visualizations.

The script:
- Cleans raw Excel data (handles missing values, removes invalid rows, applies defaults).
- Generates KPI summaries per **Customer** and **Product**.
- Adds visual **bar charts** to the Excel file for better insight.
- Logs all cleanup actions for transparency.

---

## ⚙️ Features
- ✅ Detects and logs missing data (`OrderID`, counts, prices, customer names).  
- ✅ Applies default values where possible (e.g. missing customer = "Unknown").  
- ✅ Creates **Customer KPI report**: number of orders & total order value.  
- ✅ Creates **Product KPI report**: total items sold & revenue.  
- ✅ Inserts bar charts directly into the Excel output.  
- ✅ Saves results into the same Excel file.  

---

## 🛠️ Requirements
- Python 3.10+  
- Libraries:
  - `openpyxl`
  - `logging` (built-in)

Install dependencies:
```bash
pip install openpyxl
```

---

## ▶️ Usage
1. Place your Excel input file in the project folder and rename it to:
   ```
   example_sales.xlsx
   ```
   *(The default input is `Sheet1` in the workbook.)*

2. Run the script:
   ```bash
   python main.py
   ```

3. Open the same file `example_sales.xlsx` after execution.  
   - Two new sheets will be added:
     - `Customer KPI's`
     - `Product KPI's`
   - Each contains summary tables and bar charts.

---

## 📊 Example Output
**Customer KPI’s:**

| Customer     | Total Orders | Total Value |
|--------------|--------------|-------------|
| Jan de Vries | 2            | 35.0        |
| Maria Smit   | 2            | 37.5        |
| Lisa de Jong | 1            | 21.0        |

**Product KPI’s:**

| Product  | Total Sold | Total Value |
|----------|------------|-------------|
| Pen      | 25         | 37.5        |
| Notebook | 16         | 48.0        |
| Marker   | 10         | 20.0        |

*(Bar charts are added in the Excel file itself.)*

---

## 📝 Logging
All cleanup actions are logged in:
```
cleanup.log
```

Example log:
```
2025-09-24 12:14:53 [ERROR] Missing OrderID in row 4. This row was removed from the summary
2025-09-24 12:14:53 [WARNING] Missing data in row 5. Data was set to a default value
```

---

## 📂 Project Structure
```
.
├── main.py              # Main script
├── example_sales.xlsx   # Input & output file
├── cleanup.log          # Log file with cleanup actions
└── README.md            # Documentation
```


## 📖 Learning Goals
This project was created as part of a portfolio to showcase:
- Data cleaning automation with Python.  
- Practical Excel file manipulation using `openpyxl`.  
- Generating automated KPI reports with visual charts.  
