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
    file_path="input.xlsx",
    sheet_name="Sheet1",
    chart_configs=my_survey_inputs,
    output_excel_path="output.xlsx",
    colors=custom_colors
)
