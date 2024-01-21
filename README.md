
# How Excel Manager works

### Method Chaining
What is useful about this example is that using method chaining to perform a series of operations on an instance of the `ExcelManager` class. 

Here's a step-by-step description of what each method in the chain does:

1. `add_header(data_headers)`: This method is used to add headers to the Excel sheet. The `data_headers` variable is a list of dictionaries, where each dictionary represents a header with a 'label' and 'value'. The 'label' is the actual header name and the 'value' is an additional piece of information associated with the header.

2. `add_blank()`: This method adds a blank row to the Excel sheet. This can be useful for separating different sections of the sheet for better readability.

3. `add_data_records_to_datatable(data_records)`: This method adds data records to the Excel sheet. The `data_records` variable is a list of dictionaries, where each dictionary represents a row of data in the sheet. The keys of the dictionary are the column names (which should match the header labels), and the values are the actual data for each column.

4. `create_range()`: This method calculates the range for the table based on the number of headers and records. The range is used when creating the table in the next step.

5. `make_table('Table20')`: This method creates a table in the Excel sheet with the name 'Table20'. The table includes all the headers and records added previously, and is defined by the range calculated in the previous step.

6. `set_table_style()`: This method sets the style of the table. In this case, it's using the 'TableStyleMedium9' style, which includes striped rows and columns for better readability.

7. `add_table_to_sheet()`: This method adds the previously created and styled table to the Excel sheet.

8. `save('my_excel_.xlsx')`: Finally, this method saves the Excel workbook to a file named 'my_excel_.xlsx'. After this method is called, all the changes made to the `ExcelManager` instance are written to the file.