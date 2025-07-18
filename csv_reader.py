import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Alignment
from openpyxl.styles import Border, Side, Font
import matplotlib.pyplot as plt

# Read using multi-row headers
df = pd.read_csv('test_aircrack01.csv', header=[2, 3])
# Show the top of the DataFrame
print(df.head())

df.to_excel('output.xlsx', header=[2, 3])



# Load the workbook
wb = load_workbook('output.xlsx')

# Select the active sheet (or use wb['SheetName'])
ws = wb.active


# Merge cells B2 to D2
ws.merge_cells('B4:D4')
ws.merge_cells('E4:H4')
ws.merge_cells('I4:P4')
ws.merge_cells('Q4:W4')
ws.merge_cells('X4:AA4')


# Add CPU temp & freq readings
cf = pd.read_csv('temp_freq_raw.csv')
print(cf.head())

ws.merge_cells('AB4:AC4')
ws.cell(row=4, column=28).value = 'CPU monitor'
ws.cell(row=5, column=28).value = 'Temperature'
ws.cell(row=5, column=29).value = 'Frequency'

temperatures = cf['Temperature']
frequencies = cf['Frequency']

# Load temperatures
for i, row in zip(range(0, len(temperatures)), range(6, ws.max_row + 1)):
    cell = ws.cell(row=row, column=28)
    cell.value = temperatures[i] / 1000 # Convert from milidegree C to C

# Load frequencies
for i, row in zip(range(0, len(frequencies)), range(6, ws.max_row + 1)):
    cell = ws.cell(row=row, column=29)
    cell.value = round(frequencies[i] / 1000000, 3) # Convert from raw kHz to GHz


# Choose the column (e.g., C is column 3)
for column in range(5, 17):
    for row in range(6, ws.max_row + 1):  # start at 2 to skip header
        cell = ws.cell(row=row, column=column)  # Column C
        if isinstance(float(cell.value), (int, float)):
            print("OLD VALUE: ", cell.value)
            cell.value = round(float(cell.value)/1024)
            print("NEW VALUE: ", cell.value)


# Convert percentage numbers from strings to numbers (floats)
for column in range(17, 27):
    for row in range(6, ws.max_row + 1):  # start at 2 to skip header
        cell = ws.cell(row=row, column=column)  # Column C
        if isinstance(cell.value, str):
            print("OLD VALUE: ", cell.value)
            cell.value = float(cell.value)
            print("NEW VALUE: ", cell.value)

for column in range(2, 5):
    for row in range(6, ws.max_row + 1):  # start at 2 to skip header
        cell = ws.cell(row=row, column=column)  # Column C
        if isinstance(cell.value, str):
            print("OLD VALUE: ", cell.value)
            cell.value = float(cell.value)
            print("NEW VALUE: ", cell.value)


# convert timestamps into numbers
minute = -1
for row in range(6, ws.max_row + 1):  # start at 2 to skip header

    cell = ws.cell(row=row, column=1)  # Column A
    if isinstance(cell.value, str):
        print("OLD TIME: ", cell.value)
        cell.value = minute + 1
        print("NEW TIME: ", cell.value)
    minute += 1




# Define range (e.g., B2:D10)
for row in ws['A1:AC15']:
    for cell in row:
        cell.alignment = Alignment(horizontal='center', vertical='center')


# Define a thin border
thin = Side(border_style="thin", color="000000")
border = Border(top=thin, left=thin, right=thin, bottom=thin)

# Apply to each cell in the range
for row in ws['A4:AC15']:
    for cell in row:
        cell.font = Font(bold=True)
        cell.border = border


# Save changes to a new file or overwrite
wb.save('output.xlsx')

#TODO: Add temperature and frequency readings to the sheet