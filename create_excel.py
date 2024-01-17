import openpyxl

# create a workbook
wb = openpyxl.Workbook()

# create a sheet
sheet = wb.active

# set the title of the sheet
sheet.title = 'My first sheet'

# add a list of dictionary to the sheet
records = [
    {'name': 'Bob', 'age': 20},
    {'name': 'Alice', 'age': 21},
    {'name': 'John', 'age': 22},
]

# add header
sheet.append(['Name', 'Age'])
for record in records:
    sheet.append([record['name'], record['age']])


# save the file
wb.save('my_first_excel.xlsx')