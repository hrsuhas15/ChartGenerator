# Excel Survey Chart Generator

A simple Python script to automatically generate clean, editable charts in Excel from raw survey data.

---

## 1. Requirements

Make sure you have Python installed, then install the required packages:

```bash
pip install pandas openpyxl

```

---

## 2. Quick Start

Save the script as `generator.py` and run the following code:

```python
from generator import generate_editable_report

# Define your charts: (Column Name, Chart Type, Title)
my_survey_inputs = [
    ("Age", "histogram", "Age Distribution"),
    ("Gender", "Pie", "Gender Distribution"),
    ("Overall satisfaction", "horizontal bar", "Satisfaction Score"),
]

# Custom color palette (Hex codes)
custom_colors = ["2491CE", "A5C93D", "B2B2B2", "333333"]

# Run the generator
generate_editable_report(
    file_path="test.xlsx",
    sheet_name="Sheet1",
    chart_configs=my_survey_inputs,
    output_excel_path="report.xlsx",
    colors=custom_colors
)

```

---

## 3. Supported Chart Types

* `"Pie"` — Standard pie chart with percentage labels and legend.
* `"bar"` — Vertical column chart.
* `"horizontal bar"` — Horizontal bar chart (automatically handles comma-separated multi-select answers).
* `"histogram"` — Frequency distribution chart.
