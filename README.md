# PowerPoint Report Automation

Automated PowerPoint report generation using Python.

## Quick Start

```bash
# Install dependencies (first time only)
pip3 install python-pptx pandas lxml pillow

# Generate report
python3 leo_automation.py
```

Output: `output_leo_report.pptx`

## Files

| File | Description |
|------|-------------|
| `leo_automation.py` | Main script - generates weekly report |
| `functions.py` | Core utility functions for PPT manipulation |
| `template.pptx` | PowerPoint template |
| `app.py` | Streamlit automation with all channels |
| `EM_cleaning_automation.py` | EM PBI data automation |
| `EM_clicks_cleaning_automation.py` | EM Clicks data automation |

## How It Works

1. Loads `template.pptx` as base
2. Duplicates slides and replaces placeholder text
3. Fills tables with data using `add_data_table_new()`
4. Removes original template slides
5. Saves final report

## Customization

Edit `leo_automation.py` to:
- Change report title
- Update mock data with real data
- Add/remove slides
