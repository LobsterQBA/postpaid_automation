# Postpaid Marketing Data Automation

Automated pipeline for cleaning campaign data and generating reports (PPT + Pulse Check Word doc).

## Quick Start

```bash
pip install -r requirements.txt

# Report generation (PPT or Pulse Check)
streamlit run auto_load_demo.py

# Data cleaning
streamlit run app.py
```

## Files

| File | Description |
|------|-------------|
| `auto_load_demo.py` | Main app — upload cleaned Excel, generate PPT report or Pulse Check doc |
| `functions.py` | Core PPT manipulation (slide duplication, text replacement) |
| `pulse_check_docx.py` | Pulse Check Word document generation logic |
| `template.pptx` | PPT template (3 slide layouts) |
| `pulse_check_template.docx` | Word template for Pulse Check |
| `app.py` | Data cleaning Streamlit app |
| `EM_cleaning_automation.py` | Email data cleaning (standalone) |
| `EM_clicks_cleaning_automation.py` | Email click data cleaning (standalone) |
| `Sample_Data_Input_Template.xlsx` | Sample input template with instructions |
| `PPT_Sample_Input_Data.xlsx` | Built-in sample data for quick testing |

## Input Format

Excel file with these sheets (auto-detected):

- **EM** — Email data (Deliveries, Unique Opens, Unique Clicks, Touch, Cohort, Audience Details, etc.)
- **RCM SMS** or **SMS** — SMS/RCM data (auto-detects standard vs RCM dual-channel format)
- **SLs** *(optional)* — Subject Line testing data; if absent, derived from EM's `SL Testing Variant` column
