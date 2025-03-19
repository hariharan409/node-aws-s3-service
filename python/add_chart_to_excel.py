import sys

# %%
# Import Libraries

# %%
# pip install openpyxl

from io import BytesIO

from openpyxl import load_workbook

from openpyxl.chart import LineChart, Reference, Series

if __name__ == "__main__":
    # Step 1: Get log date from sys.argv
    sheet_name = sys.argv[1]  # Extract log date

    # Step 2: Read the Excel file buffer from stdin
    file_bytes = sys.stdin.buffer.read()

# %%
# Step 3: Load the Excel file into memory using openpyxl
wb = load_workbook(BytesIO(file_bytes))

# %%
# Step 4: Access the specific sheet by name
if sheet_name not in wb.sheetnames:
    raise ValueError(f"Sheet '{sheet_name}' not found in the workbook.")
ws = wb[sheet_name]

# %%
def create_line_charts(ws, start_data_row=4, end_data_row=16, start_chart_cell="B25", vertical_offset=15):   
    # Create a reference for the x-axis categories from S3:X3
    x_ref = Reference(ws, min_col=19, max_col=24, min_row=3, max_row=3)
   
    # Determine the starting chart row from start_chart_cell (e.g., "A25" -> 25)
    start_chart_row = int(''.join(filter(str.isdigit, start_chart_cell)))
   
    # Loop over each row with a chart title in column A
    for r in range(start_data_row, end_data_row + 1):
        chart_title = ws.cell(row=r, column=1).value       
        # Create a new LineChart
        chart = LineChart()
        chart.title = chart_title
        chart.style = 2
        # chart.x_axis.title = "Time"
        # chart.y_axis.title = "Value (Deg F)"
        # Force the x-axis to use text formatting
        chart.x_axis.number_format = "@"

        # Ensure axis labels are displayed
        chart.x_axis.delete = False
        chart.y_axis.delete = False

        # Set the x-axis categories to the fixed labels from S3:X3
        chart.set_categories(x_ref)
       
        # Define Y-axis data ranges for each engine using the current row:
        eng1_ref = Reference(ws, min_col=7,  max_col=12, min_row=r, max_row=r)    # Engine 1: G{r}:L{r}
        eng2_ref = Reference(ws, min_col=13, max_col=18, min_row=r, max_row=r)    # Engine 2: M{r}:R{r}
        eng3_ref = Reference(ws, min_col=19, max_col=24, min_row=r, max_row=r)    # Engine 3: S{r}:X{r}
        eng4_ref = Reference(ws, min_col=25, max_col=30, min_row=r, max_row=r)    # Engine 4: Y{r}:AD{r}
        eng5_ref = Reference(ws, min_col=31, max_col=36, min_row=r, max_row=r)    # Engine 5: AE{r}:AJ{r}

         # Create Series and disable smoothing
        for ref, title in zip([eng1_ref, eng2_ref, eng3_ref, eng4_ref, eng5_ref], 
                              ["Engine 1", "Engine 2", "Engine 3", "Engine 4", "Engine 5"]):
            series = Series(ref, title=title)
            series.graphicalProperties.line.smooth = False
            chart.series.append(series)
       
        # Compute the anchor cell for this chart (offset vertically)
        anchor_row = start_chart_row + (r - start_data_row) * vertical_offset
        anchor_cell = "B" + str(anchor_row)
        ws.add_chart(chart, anchor_cell)

# Step 4: Process the Excel file (add charts)
create_line_charts(ws, start_data_row=4, end_data_row=16, start_chart_cell="B25", vertical_offset=15)

# Step 5: Save the modified file to an in-memory buffer
output_buffer = BytesIO()
wb.save(output_buffer)

# Step 6: Write the modified buffer back to stdout
sys.stdout.buffer.write(output_buffer.getvalue())



