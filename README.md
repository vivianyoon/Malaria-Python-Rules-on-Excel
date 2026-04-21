# 🦟 Malaria Excel Rule Engine

A **Streamlit web application** for validating malaria data using rules defined in Excel.

This project demonstrates how business users can define validation logic in Excel, while Python dynamically applies those rules to datasets — no code changes required.

---

## 🚀 Live Demo

👉 *(https://malaria-python-rules-on-excel-sq5hw2jeynbuxemoqm8aq8.streamlit.app/)*

---

## 💡 Overview

Data validation in Excel is often manual, error-prone, and hard to scale.

This application solves that by:

* Allowing rules to be defined in Excel
* Automatically applying those rules to datasets
* Generating clear validation outputs and summaries

---
## ⭐ Key Highlights

- Excel-driven rule engine for flexible validation
- Designed for non-technical users (no coding required)
- Supports multi-sheet data validation workflows
- Generates structured error summaries for analysis

## ✨ Features

* 📄 Upload rule workbook (Excel)
* 📊 Upload malaria data workbook
* ⚙️ Apply rule sets dynamically
* 🧩 Supports multiple validation types:

  * Missing values
  * Allowed choices
  * Numeric validation
  * Range checks
  * Date format validation
  * Date range validation
  * Conditional consistency checks
  * Mapping / derived fields
* 📑 Multi-sheet processing
* 🧾 Error summaries (type + detailed)
* 📥 Download processed Excel output

---

## 🧠 How It Works

1. Rules are defined in Excel:

   * `Rules` sheet → validation logic
   * `Mappings` sheet → lookup transformations

2. User uploads:

   * Rule workbook
   * Data workbook

3. Application:

   * Reads rules using Pandas
   * Applies validation logic dynamically
   * Adds a `COMMENT` column with errors

4. Output:

   * Processed dataset
   * Error summary sheets
   * Downloadable Excel file

---

## 📂 Project Structure

```bash
.
├── streamlit_rules_engine_app.py
├── requirements.txt
├── sample_malaria_input.xlsx
├── sample_malaria_rules.xlsx
├── .gitignore
└── README.md
```

---

## ⚙️ Installation & Run

### 1. Clone the repository

```bash
git clone https://github.com/vivianyoon/Malaria-Python-Rules-on-Excel.git
cd Malaria-Python-Rules-on-Excel
```

### 2. Install dependencies

```bash
pip install -r requirements.txt
```

### 3. Run the app

```bash
streamlit run streamlit_rules_engine_app.py
```

---

## 📊 Input Files

### 🔹 Rule Workbook

Must include:

* `Rules` sheet
* `Mappings` sheet

Defines:

* validation conditions
* expected values
* ranges
* mappings

---

### 🔹 Data Workbook

* Excel file (`.xlsx` or `.xls`)
* Can contain one or multiple sheets

---

## 📤 Output

The app generates:

* ✅ Processed Excel file
* ✅ `COMMENT` column for validation errors
* ✅ Error summary (by type)
* ✅ Detailed error breakdown

---

## 🛠 Tech Stack

* Python
* Streamlit
* Pandas
* OpenPyXL

---

## 🌐 Deployment (Streamlit Cloud)

1. Push this repository to GitHub
2. Go to [https://share.streamlit.io](https://malaria-python-rules-on-excel-sq5hw2jeynbuxemoqm8aq8.streamlit.app/)
3. Select:

   * Repo: this repository
   * File: `streamlit_rules_engine_app.py`
4. Deploy

---

## 🎯 Use Cases

* Health data validation
* Excel-based rule engines
* Data quality assurance workflows
* Non-technical rule configuration systems

---

## 🔮 Future Improvements

* UI-based rule creation (no Excel required)
* Database integration
* API version of rule engine
* Logging and audit tracking
* Performance optimization for large datasets

---

## 👤 Author

**Vivian Yoon**

---

## ⭐ Why This Project Matters

This project showcases:

* Real-world data validation challenges
* Excel + Python integration
* Dynamic rule execution
* Building tools for non-technical users
* End-to-end data processing workflows

---
