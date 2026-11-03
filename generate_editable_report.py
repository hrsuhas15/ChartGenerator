import pandas as pd
import openpyxl
from openpyxl.chart import BarChart, PieChart, Reference, Series
from openpyxl.chart.label import DataLabelList
from openpyxl.chart.series import DataPoint 
from openpyxl.chart.legend import Legend
from openpyxl.chart.text import RichText
from openpyxl.drawing.text import Paragraph, ParagraphProperties, CharacterProperties, Font
from openpyxl.chart.layout import Layout, ManualLayout
from openpyxl.chart.shapes import GraphicalProperties
from openpyxl.drawing.line import LineProperties

def generate_editable_report(file_path, sheet_name, chart_configs, output_excel_path, colors=None):
    # 1. Read input workbook structures
    df = pd.read_excel(file_path, sheet_name=sheet_name)
    wb = openpyxl.load_workbook(file_path)
    
    # Refresh dashboard clean tab canvas
    report_sheet_name = "Dashboard"
    if report_sheet_name in wb.sheetnames:
        del wb[report_sheet_name]
    ws = wb.create_sheet(title=report_sheet_name)
    
    # Refresh raw storage tab calculation canvas
    data_sheet_name = "Summary_Data_Internal"
    if data_sheet_name in wb.sheetnames:
        del wb[data_sheet_name]
    ws_data = wb.create_sheet(title=data_sheet_name)
    
    data_current_row = 1
    chart_current_row = 2
    
    # Default Color Matrix if none provided (Primary, Secondary, Neutral Light, Dark)
    if colors is None:
        allowed_colors = ["2491CE", "A5C93D", "B2B2B2", "333333"]
    else:
        allowed_colors = colors

    # Font Property Objects for Native Custom Styling (sz is scaled in 1/100ths of a pt)
    font_arial = Font(typeface='Arial')
    cp_title = CharacterProperties(latin=font_arial, sz=1000, b=True) 
    cp_regular = CharacterProperties(latin=font_arial, sz=900, b=False)

    # 2. Iterate dynamically over inputs list
    for column_name, chart_type, chart_title in chart_configs:
        if column_name not in df.columns:
            print(f"Skipping missing column validation: {column_name}")
            continue
            
        data_clean = df[column_name].dropna()
        total_responses = len(data_clean)
        if total_responses == 0:
            continue
            
        c_type = chart_type.lower()
        if 'horizontal bar' in c_type:
            # Splits by commas ONLY if they are outside of parentheses
            series_data = data_clean.astype(str).str.split(r'\s*,\s*(?![^()]*\))').explode()
            total_responses = len(series_data)
        else:
            series_data = data_clean
            
        # Dynamic Title Formatting rule: Append (n = number of rows value)
        formatted_title = f"{chart_title} (n = {total_responses})"
        
        series_data = series_data.astype(str).str.strip()
        counts = series_data.value_counts()
        percentages = (counts / total_responses) * 100
        pivot_df = pd.DataFrame({'Category': percentages.index, 'Percentage': percentages.values})
        pivot_df['Percentage'] = pivot_df['Percentage'].round(0).astype(int)

        c_type = chart_type.lower()  # define early so sort can use it
        if 'horizontal bar' in c_type:
            # Excel reverses horizontal bar order visually, so write 5→1 to display 1→5 on screen
            pivot_df = pivot_df.sort_values(by='Category', ascending=False).reset_index(drop=True)
        else:
            pivot_df = pivot_df.sort_values(by='Percentage', ascending=False).reset_index(drop=True)

        # Track which category has the highest percentage (for color assignment)
        max_pct_category = pivot_df.loc[pivot_df['Percentage'].idxmax(), 'Category']
        
        # 3. Create Pivot Summary
        start_row = data_current_row
        ws_data.cell(row=start_row, column=1, value=column_name) # Title Header
        ws_data.cell(row=start_row, column=2, value="")          # Blank placeholder cell
        
        for idx, row in pivot_df.iterrows():
            data_current_row += 1
            ws_data.cell(row=data_current_row, column=1, value=row['Category'])
            
            # Formatting Switch Rule for Pie vs Bars
            if 'pie' in c_type:
                ws_data.cell(row=data_current_row, column=2, value=float(row['Percentage'] / 100))
                ws_data.cell(row=data_current_row, column=2).number_format = '0%'
            else:
                ws_data.cell(row=data_current_row, column=2, value=int(row['Percentage']))
                ws_data.cell(row=data_current_row, column=2).number_format = '0"%"'
            
        end_row = data_current_row
        data_current_row += 2 
        
        # 4. Initialize Native Charts Layout Maps
        if 'pie' in c_type:
            chart = PieChart()
            chart.type = "standard"
        elif 'horizontal bar' in c_type:
            chart = BarChart()
            chart.type = "bar" 
            chart.grouping = "standard"
        else:
            chart = BarChart()
            chart.type = "col" 
            chart.grouping = "standard"
        
        # Remove chart area outline completely
        chart.graphical_properties = GraphicalProperties(
            ln=LineProperties(noFill=True)
        )
        
        # Select data starting from row+1 to skip the blank row header entirely
        data_ref = Reference(ws_data, min_col=2, min_row=start_row + 1, max_row=end_row)
        cats_ref = Reference(ws_data, min_col=1, min_row=start_row + 1, max_row=end_row)
        
        # 5. Connect data references without pulling default title strings
        chart_series_obj = Series(data_ref, title_from_data=False)
        chart.append(chart_series_obj)
        chart.set_categories(cats_ref)
        
        # Core Title Initialization - Injects title string natively
        chart.title = formatted_title
        
        # Core Title Sizing Fix
        chart.title.tx.rich.p[0].pPr = ParagraphProperties(defRPr=cp_title)
        chart.title.tx.rich.p[0].endParaRPr = cp_title
        
        # Add spacing between title and chart
        chart.layout = Layout(
            manualLayout=ManualLayout(
                x=0.05,
                y=1,   # Moves graph DOWN from title
                h=0.55,   # Reduces chart height slightly
                w=0.70
            )
        )

        # 6. Apply clean labels rule parameters
        chart.dataLabels = DataLabelList()
        chart.dataLabels.showSerName = False
        chart.dataLabels.showLegendKey = False
        
        # Font injection setup for data labels (Arial 11 Regular)
        chart.dataLabels.txPr = RichText(p=[Paragraph(pPr=ParagraphProperties(defRPr=cp_regular), endParaRPr=cp_regular)])
        
        if 'pie' in c_type:
            chart.dataLabels.showPercent = True  
            chart.dataLabels.showVal = False      
            chart.dataLabels.showCatName = False  
            chart.dataLabels.position = "outEnd" 
            
            chart.legend = Legend()
            chart.legend.position = "r"         
            chart.legend.txPr = RichText(p=[Paragraph(pPr=ParagraphProperties(defRPr=cp_regular), endParaRPr=cp_regular)])
            chart_series_obj.title = None
        else:
            chart.dataLabels.showVal = True      
            chart.dataLabels.showPercent = False  
            chart.dataLabels.showCatName = False  
            chart.dataLabels.position = "outEnd"  
            chart.legend = None                  
        
        if hasattr(chart, 'y_axis') and hasattr(chart, 'x_axis'):
          chart.x_axis.majorGridlines = None
          chart.y_axis.majorGridlines = None
      
          chart.x_axis.tickLblPos = "nextTo"
          chart.y_axis.tickLblPos = "nextTo"
      
          chart.x_axis.delete = False
          chart.y_axis.delete = False
      
          if 'horizontal bar' in c_type:
              chart.x_axis.number_format = '0"%"'
              chart.y_axis.number_format = 'General'
          else:
              chart.y_axis.number_format = '0"%"'
              chart.x_axis.number_format = 'General'
      
          chart.x_axis.title = ""
          chart.y_axis.title = ""
          
          cp_axis = CharacterProperties(
              latin=Font(typeface='Arial'),
              sz=900
          )
      
          axis_rt = RichText(
              p=[Paragraph(
                  pPr=ParagraphProperties(defRPr=cp_axis),
                  endParaRPr=cp_axis
              )]
          )
      
          chart.x_axis.txPr = axis_rt
          chart.y_axis.txPr = axis_rt
            
        # 7. Apply solid coloring shades using allowed_colors
        chart.varyColors = False
        num_categories = len(pivot_df)
        
        sorted_categories = pivot_df['Category'].tolist()

        for i in range(num_categories):
            pt = DataPoint(idx=i)
            color_idx = i % len(allowed_colors)
            
            if 'histogram' in c_type:
                pt.graphicalProperties.solidFill = allowed_colors[0]
            else:
                cat = sorted_categories[i]
                if cat == max_pct_category:
                    pt.graphicalProperties.solidFill = allowed_colors[0]  # Primary — highest %
                elif len(allowed_colors) > 1 and pivot_df.loc[pivot_df['Category'] == cat, 'Percentage'].values[0] == \
                     pivot_df['Percentage'].nlargest(2).iloc[-1] and cat != max_pct_category:
                    pt.graphicalProperties.solidFill = allowed_colors[1]  # Secondary — 2nd highest %
                else:
                    pt.graphicalProperties.solidFill = allowed_colors[color_idx]  # Rotating/Fallback color
                
            chart_series_obj.dPt.append(pt) 

        # 8. Append chart back to worksheet layout grid
        ws.add_chart(chart, f"D{chart_current_row}")
        chart_current_row += 20

    wb.save(output_excel_path)
    print(f"Report successfully saved to: {output_excel_path}")


# --- RUN EXECUTION LOOP ---
if __name__ == "__main__":
    my_survey_inputs = [
        ("Age", "histogram", "Age Distribution"),
        ("Gender", "Pie", "Gender Distribution"),
        ("Social Category", "bar", "Social Category"),
    ]
    
    # Custom colors can be passed as a list of Hex strings
    custom_colors = ["2491CE", "A5C93D", "B2B2B2", "333333"]

    generate_editable_report(
        file_path="test.xlsx",
        sheet_name="Sheet1",
        chart_configs=my_survey_inputs,
        output_excel_path="report.xlsx",
        colors=custom_colors
    )
