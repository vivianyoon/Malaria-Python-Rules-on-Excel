# 🦟 Malaria Excel Rule Engine

A **Streamlit-based rule engine** for validating malaria data using rules defined in Excel.

This project allows non-technical users to define validation logic in Excel while Python dynamically applies those rules to real datasets.

---

## 🚀 Demo

Upload:

* 📄 Rule workbook (Excel)
* 📊 Malaria data workbook

Then:

* Run validation
* View errors
* Download processed Excel with summaries

---

## ✨ Features

* Excel-driven validation (no code needed for rules)
* Supports multiple rule types:

  * Missing value checks
  * Choice validation
  * Numeric validation
  * Range checks
  * Date format validation
  * Date range validation
  * Conditional consistency checks
  * Mapping / derived columns
* Multi-sheet processing support
* Automatic error summaries
* Export clean Excel output with comments

---

## 🧠 How It Works

1. Define rules in Excel (`Rules` + `Mappings` sheets)
2. Upload rules + malaria dataset
3. App dynamically applies rules
4. Output includes:

   * `COMMENT` column with validation errors
   * Error summary tables
   * Processed Excel file

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

## ⚙️ Installation

### 1. Clone repository

```bash
git clone https://github.com/your-username/Malaria-Python-Rules-on-Excel.git
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

## 📊 Input Requirements

### 🔹 Rule Workbook

Must contain:

* `Rules` sheet
* `Mappings` sheet

Rules define validation logic such as:

* allowed values
* numeric ranges
* date formats
* conditional checks

---

### 🔹 Data Workbook

* Excel file (`.xlsx` or `.xls`)
* Can contain one or multiple sheets

---

## 📤 Output

The app generates:

* ✅ Processed Excel file
* ✅ `COMMENT` column with validation issues
* ✅ Error summary by type
* ✅ Detailed error breakdown

---

## 🛠 Tech Stack

* Python
* Streamlit
* Pandas
* OpenPyXL

---

## 💡 Use Cases

* Health data validation
* Excel-based rule engines
* Data quality pipelines
* Non-technical rule configuration systems

---

## 🌐 Deployment

You can deploy this app using **Streamlit Cloud**:

1. Push this repo to GitHub
2. Go to https://share.streamlit.io
3. Deploy using:

   * Repo: this repository
   * File: `streamlit_rules_engine_app.py`

---

## 📌 Future Improvements

* UI-based rule builder (no Excel needed)
* Database integration
* API version of rule engine
* Logging & audit tracking
* Performance optimization for large datasets

---

## 👤 Author

**Vivian Yoon**

---

## ⭐ Why This Project Matters

This project demonstrates:

* Real-world data validation workflows
* Excel + Python integration
* Dynamic rule execution
* Building tools for non-technical users

---
