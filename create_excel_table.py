import openpyxl
from openpyxl.worksheet.table import Table, TableStyleInfo

wb = openpyxl.Workbook()
sheet = wb.active
sheet.title = 'Table'

# add multi-line headers
headers = [
    {'label': 'Title', 'value': ''},
    {'label': 'Pattern_A', 'value': 'A'},
    {'label': 'Pattern_B', 'value': 'B'},
    {'label': 'Pattern_C', 'value': 'C'},
    {'label': 'Pattern_D', 'value': '1'},
    {'label': 'Pattern_E', 'value': '2'},
    {'label': 'Pattern_F', 'value': '3'},
]

# add data table
data_table = [
    {'name': 'Bob', 'age': 20, 'Address': 'xxx'},
    {'name': 'Alice', 'age': 21, 'Address': 'xxx'},
    {'name': 'John', 'age': 22, 'Address': 'xxx'},
]

# add headers to rows.
for header in headers:
    sheet.append([header['label'], header['value']])

# add a blank row
sheet.append([])

# add header for data table
sheet.append(list(data_table[0].keys()))

for record in data_table:
    sheet.append(list(record.values()))

# set start cell, note that the row number starts from len(headers) + 2
start_cell_column = 'A'
start_cell_row = len(headers) + 2
start_cell = f"{start_cell_column}{str(start_cell_row)}"

# calculate and set end cell
end_cell_column = chr(ord(start_cell_column) + len(list(data_table[0].keys()))-1)
end_cell_row = start_cell_row + len(data_table)
end_cell = f"{end_cell_column}{str(end_cell_row)}"

table = Table(displayName="Table1", ref=f"{start_cell}:{end_cell}")

#add a default style with striped rows and banded columns
table.tableStyleInfo = TableStyleInfo(
    name="TableStyleMedium9", showFirstColumn=False,
    showLastColumn=False, showRowStripes=True, showColumnStripes=True)

# add table to worksheet
sheet.add_table(table)

wb.save('my_second_excel.xlsx')