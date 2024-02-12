import tkinter as tk
from tkinter import ttk
import openpyxl

# Define path to your Excel file
excel_file_path = "employees.xlsx"

# Open the Excel file
workbook = openpyxl.load_workbook(excel_file_path)

# Choose the sheet you want to read (default is sheet 1)
sheet = workbook.active

# Read data from the sheet
data = []
for row in sheet.iter_rows(min_row=2):  # Skip header row
    data_row = [cell.value for cell in row]
    data.append(data_row)

# Update rows and columns based on Excel data
rows = len(data) + 1  # Include header row
columns = len(data[0])

# Define table dimensions
# rows = 4
# columns = 5

# Create the main window
window = tk.Tk()
window.title("Employee Attendance Management System")
window.geometry("800x600")

# Function to create a combobox with department options
def create_combobox(department_list):
    variable = tk.StringVar()
    combobox = ttk.Combobox(window, values=department_list, textvariable=variable, state='readonly')
    # combobox.set('Pas encore saisis!')
    return combobox, variable

# Function to create a table cell
def create_cell(value, row, column):
    label = tk.Label(window, text=value)
    label.grid(row=row, column=column, padx=5, pady=5)

# Function to create a combobox cell
def create_combobox_cell(department_list, row, column):
    combobox, variable = create_combobox(department_list)
    combobox.grid(row=row, column=column, padx=5, pady=5)
    return variable

# Define department options
codifications =  ['C', '1', '7', '6', '8', '9', 'A', 'R', 'T', 'I', '2', 'Cr', 'M']





# Read column names from Excel
column_names = [cell.value for cell in sheet[1]]
print(column_names)

# Create table headers using column names
for i, column_name in enumerate(column_names):
    label = tk.Label(window, text=column_name)
    label.grid(row=0, column=i, padx=5, pady=5)

# Create table cells
for row in range(1, rows+1):
    for col in range(columns):
        if col == 3:  # Check if current column is for combobox (adjust index as needed)
            variable = create_combobox_cell(codifications, row, col)
        else:
            value = data[row-1][col] if row-1 < len(data) else ""
            create_cell(value, row, col)

# Run the main loop
window.mainloop()