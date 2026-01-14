# Postpaid Marketing Data Automation

Automated tools for processing Postpaid marketing data and generating PowerPoint reports.

## Features

- **Data Cleaning**: Automatically clean Email/SMS marketing data, extract key fields from Delivery Labels
- **PPT Report Generation**: Auto-generate weekly PowerPoint reports from templates
- **Web Apps**: Streamlit interfaces for uploading data and downloading cleaned results or PPT reports

## Quick Start

```bash
# Install dependencies
pip install -r requirements.txt

# Run data cleaning web app
streamlit run app.py

# Run PPT generator web app
streamlit run ppt_generator_app.py

# Generate report via command line
python report_automation.py
```

## Files

| File | Description |
|------|-------------|
| `app.py` | Streamlit data cleaning app - supports Email/SMS data cleaning |
| `ppt_generator_app.py` | Streamlit PPT generator app - upload Excel to generate reports |
| `report_automation.py` | Command line PPT report generation script |
| `functions.py` | Core PPT manipulation functions |
| `EM_cleaning_automation.py` | Email data cleaning script (standalone) |
| `EM_clicks_cleaning_automation.py` | Email click data cleaning script (standalone) |
| `template.pptx` | PowerPoint template file |
| `data_template.xlsx` | Sample data template |

## Data Format Requirements

### PPT Generation (data_template.xlsx)
Excel file must contain the following sheets:
- `EM` - Email data
- `SMS` - SMS data  
- `SL Testing` - Subject Line testing data

### Data Cleaning (app.py)
Supports uploading:
- Raw Email Data (PBI export)
- Email Click Data (AcV8 export)
- SMS Data (PBI or Branch.io export)
- Deploy Document (MD/DD Excel)

## Dependencies

- python-pptx - PPT manipulation
- pandas - Data processing
- streamlit - Web interface
- openpyxl - Excel read/write
- lxml, pillow, numpy - Supporting libraries
