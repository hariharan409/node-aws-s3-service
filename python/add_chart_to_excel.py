import sys
from io import BytesIO
from openpyxl import load_workbook
from openpyxl.chart import LineChart, Reference, Series
from openpyxl.utils import get_column_letter, column_index_from_string

def set_uniform_column_widths(ws, start_col="A", end_col="AE", width=15):

    start_idx = column_index_from_string(start_col)
    end_idx = column_index_from_string(end_col)
    for col_idx in range(start_idx, end_idx + 1):
        ws.column_dimensions[get_column_letter(col_idx)].width = width

def create_line_charts_fixed(ws, start_data_row=4, end_data_row=16):

    # 3 rows × 4 columns of charts (12 total):
    anchor_grid = [
        ["A25",  "G25",  "M25",  "S25"],
        ["A40",  "G40",  "M40",  "S40"],
        ["A55",  "G55",  "M55",  "S55"],
    ]
    
    # X-axis categories: cells S3:X3 (columns 19–24, row 3)
    x_ref = Reference(ws, min_col=19, max_col=24, min_row=3, max_row=3)
    
    # Custom colors for each engine (hex codes without '#')
    engine_colors = [
        "FF00FF",
        "3333FF",
        "FF6600",
        "7030A0",
        "064028",
    ]
    
    chart_index = 0
    total_charts = end_data_row - start_data_row + 1
    
    for row_idx in range(len(anchor_grid)):
        for col_idx in range(len(anchor_grid[row_idx])):
            if chart_index >= total_charts:
                break  # No more data rows to chart
            
            data_row = start_data_row + chart_index
            chart_title = ws.cell(row=data_row, column=1).value  # Title in column A
            
            chart = LineChart()
            chart.title = chart_title
            chart.style = 2
            chart.x_axis.number_format = "@"
            chart.x_axis.delete = False
            chart.y_axis.delete = False
            chart.set_categories(x_ref)
            
            # Data references for each engine on the current row
            eng_refs = [
                Reference(ws, min_col=7,  max_col=12, min_row=data_row, max_row=data_row),  # Engine 1: G:L
                Reference(ws, min_col=13, max_col=18, min_row=data_row, max_row=data_row), # Engine 2: M:R
                Reference(ws, min_col=19, max_col=24, min_row=data_row, max_row=data_row), # Engine 3: S:X
                Reference(ws, min_col=25, max_col=30, min_row=data_row, max_row=data_row), # Engine 4: Y:AD
                Reference(ws, min_col=31, max_col=36, min_row=data_row, max_row=data_row), # Engine 5: AE:AJ
            ]
            
            # Apply custom line colors to each series
            for ref, eng_title, color in zip(
                eng_refs,
                ["Engine 1", "Engine 2", "Engine 3", "Engine 4", "Engine 5"],
                engine_colors
            ):
                series = Series(ref, title=eng_title)
                series.graphicalProperties.line.smooth = False
                series.graphicalProperties.line.solidFill = color  # Set the line color
                chart.series.append(series)
            
            # Fix chart dimensions
            chart.width = 15
            chart.height = 7
            
            # Place the chart in the fixed anchor cell
            anchor_cell = anchor_grid[row_idx][col_idx]
            ws.add_chart(chart, anchor_cell)
            
            chart_index += 1

if __name__ == "__main__":
    sheet_name = sys.argv[1]
    file_bytes = sys.stdin.buffer.read()

    wb = load_workbook(BytesIO(file_bytes))
    if sheet_name not in wb.sheetnames:
        raise ValueError(f"Sheet '{sheet_name}' not found in the workbook.")
    ws = wb[sheet_name]
    
    # 1) Set uniform column widths so spacing is consistent
    set_uniform_column_widths(ws, start_col="A", end_col="AE", width=15)
    
    # 2) Create the charts in a fixed grid with custom colors
    create_line_charts_fixed(ws, start_data_row=4, end_data_row=16)
    
    # 3) Save to stdout
    output_buffer = BytesIO()
    wb.save(output_buffer)
    sys.stdout.buffer.write(output_buffer.getvalue())
