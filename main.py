import tkinter as tk
from tkinter import ttk
import ttkbootstrap as ttkb
from ttkbootstrap import *
import pandas as pd
import openpyxl
import datetime  
import datetime  

# Create tkinter window
window = tk.Tk()
window.title("Employee Attendance Management System")
window.geometry("800x600")

# Apply ttkbootstrap style
style = ttkb.Style(theme="cosmo")

# Global variables
treeview_search = None
combobox = None
confirmer = None
annuler = None


codifications =  ['C', '1', '7', '6', '8', '9', 'A', 'R', 'T', 'I', '2', 'Cr', 'M']
today = datetime.date.today()
today_day = today.day
month_day = today.month

def row_col_dic(today_day, month_day):
    row_sheet_map = {i: i + 13 for i in range(1, 13)}
    col_sheet_map = {i: chr(i + ord('B') - 1) for i in range(1, 26)}
    dic2 = {i: 'A' + chr(i + 39) for i in range(26, 32)}
    col_sheet_map.update(dic2)
    return row_sheet_map[month_day], col_sheet_map[today_day]


# Function to handle search button click
def search_employee():
    global treeview_search
    global combobox
    global confirmer
    global annuler
    # Get the Matricule (ID) entered by the user
    matricule = matricule_entry.get()
    if matricule:

        result_label.config(text=f"Searching for employee with Matricule {matricule}")
        # Filter data based on Matricule
        filtered_df = df[df['Matricule'] == int(matricule)]
        if not filtered_df.empty:

            treeview_search = ttk.Treeview(search_frame, columns=list(filtered_df.columns)+['Aujordhui'], height=2)
            treeview_search.grid(row=2, column=0, columnspan=len(filtered_df.columns), sticky="nsew")

            treeview_search.column("#0",width=0,minwidth=0)
            # Insert columns
            for column in filtered_df.columns:
                treeview_search.heading(column, text=column, anchor=tk.W)
                treeview_search.heading('Aujordhui', text='Aujordhui', anchor=tk.W)
                treeview_search.column(column, width=100, anchor=tk.W)
            # Insert data
            treeview_search.insert('', 'end', text=0, values=tuple(filtered_df.values[0]))

            result_label.config(text='Trouvé')

            combobox = ttk.Combobox(search_frame, values=codifications, textvariable='smt', state='readonly')
            combobox.grid(row=3, column=0)    
            confirmer = ttkb.Button(search_frame, text="Confirmer", command=affecter_jour, bootstyle='SUCCESS')
            confirmer.grid(row=3, column=2, )
            # Check if empty and update user status
            if not cell.value:
                # add_value_to_last_column(treeview_search,'Pas encore!')
                pass

            else:
                # add_value_to_last_column(treeview_search,'Fait!')
                pass
                        
            return  
        
        result_label.config(text='Ce matricute néxiste pas!')
        return

    result_label.config(text=f"Enter a valid Number!")



# Function to display employees in Treeview
def display_employees(df):

    # Clear existing items in Treeview
    # for item in employees_treeview.get_children():
    #     employees_treeview.delete(item)
    # Display data in Treeview
    for index, row in df.iterrows():
        excel_file_path = f'.\\PointageAnnuel\\{row.tolist()[1]}-{today.year}.xlsx'
        print(excel_file_path)
        # Open Excel file and access sheet
        workbook = openpyxl.load_workbook(excel_file_path)
        # Choose the sheet you want to read (default is sheet 1)
        sheet = workbook.active
        
        # Call row_col_dic function
        curr_day_row, curr_day_col = row_col_dic(today_day, month_day)

        # Access the cell for today's appointment
        cell = sheet[str(curr_day_col)+str(curr_day_row)]
        if not cell.value:
            employees_treeview.insert("", "end", iid=index, text=index+1 , values= row.tolist()+['Pas Encore!'])
        else:
            employees_treeview.insert("", "end", iid=index, text=index+1 , values= row.tolist()+['Fait!'])



# Function to handle reset button click
def reset_filter():
    treeview_search.grid_forget()
    combobox.grid_forget()
    confirmer.grid_forget()
    # Clear the result label
    result_label.config(text='')

def affecter_jour():
    selected_value = combobox.get()
    print("Selected value:", selected_value)


# Validation function to allow only integer values
def validate_matricule_input(value):
    if value.isdigit() or value == "":
        return True
    else:
        return False

# Create and configure frames
search_frame = ttk.Frame(window, padding="20")
search_frame.pack(expand=True, fill=tk.BOTH)


# Create and configure widgets for the reset button
reset_button = ttkb.Button(search_frame, text="RESET", command=reset_filter, bootstyle='SECONDARY-OUTLINE')
reset_button.grid(row=0, column=3, padx=5, pady=5)


# Create and configure widgets
matricule_label = ttk.Label(search_frame, text="Enter Matricule:")
matricule_label.grid(row=0, column=0, padx=5, pady=5)

matricule_entry = ttk.Entry(search_frame, validate="key")
matricule_entry.grid(row=0, column=1, padx=5, pady=5)

# Apply validation function to matricule_entry
validate_matricule = window.register(validate_matricule_input)
matricule_entry.config(validatecommand=(validate_matricule, "%P"))


search_button = ttk.Button(search_frame, text="Search", command=search_employee)
search_button.grid(row=0, column=2, padx=5, pady=5)

result_label = ttk.Label(search_frame, text="")
result_label.grid(row=1, columnspan=3, padx=5, pady=5)


# Create Treeview to display employees
employees_treeview = ttk.Treeview(window, columns=["Name", "Matricule", "Department", "Aujordhui"])
employees_treeview.heading("#0", text="Index", anchor=tk.W)
employees_treeview.heading("Name", text="Name", anchor=tk.W)
employees_treeview.heading("Matricule", text="Matricule", anchor=tk.CENTER)
employees_treeview.heading("Department", text="Department", anchor=tk.W)
employees_treeview.heading("Aujordhui", text="Aujordhui", anchor=tk.W)
# Adjust column widths and alignments
employees_treeview.column("#0", width=50,minwidth=50, anchor=tk.W)
employees_treeview.column("Name", width=150, anchor=tk.W)
employees_treeview.column("Matricule", width=100, anchor=tk.CENTER)
employees_treeview.column("Department", width=200, anchor=tk.W)
employees_treeview.column("Aujordhui", width=200, anchor=tk.W)
# Set up Treeview style
style = ttk.Style()
# style.configure("Treeview", anchor=tk.W, font=("Arial", 10, "bold"))

# Pack Treeview
employees_treeview.pack(pady=20)



df = pd.read_excel("employees.xlsx")
display_employees(df)















# PARTIE FEUILLE DE POINTAGE

# excel_file_path = "Template.xlsx"
# # Define month and day information
# # month_column_index = 0  # Adjust based on your Excel file
# # current_day_row_index = 2  # Adjust based on current date

# # Open Excel file and access sheet
# workbook = openpyxl.load_workbook(excel_file_path)
# # Choose the sheet you want to read (default is sheet 1)
# sheet = workbook.active

 
# def row_col_dic(today_day, month_day):
#     row_sheet_map = {i: i + 13 for i in range(1, 13)}
#     col_sheet_map = {i: chr(i + ord('B') - 1) for i in range(1, 26)}
#     dic2 = {i: 'A' + chr(i + 39) for i in range(26, 32)}
#     col_sheet_map.update(dic2)
#     return row_sheet_map[month_day], col_sheet_map[today_day]
  
# today = datetime.date.today()

# # Extract day and month
# today_day = today.day
# month_day = today.month

# # Call row_col_dic function
# curr_day_row, curr_day_col = row_col_dic(today_day, month_day)

# # Access the cell for today's appointment
# cell = sheet[str(curr_day_col)+str(curr_day_row)]


# def add_value_to_last_column(tree, value):
#     # Get all the item identifiers in the treeview
#     if tree:
#         item_identifiers = tree.get_children()  
#         # Iterate through each item identifier
#         for item_id in item_identifiers:
#             # Get the values of the item
#             item_values = tree.item(item_id)['values']
            
#             # Update the last value of the item
#             item_values.append(value) 
            
#             # Update the values of the item in the treeview
#             tree.item(item_id, values=item_values)


# # Check if empty and update user status
# if not cell.value:
#     print("User has not checked today's appointment")
#     # employees_treeview.set('Aujordhui', 'Pas encore!' )
#     add_value_to_last_column(employees_treeview,"Pas encore!")
#     # add_value_to_last_column(treeview_search,'Pas encore!')

# else:
#     print("User has checked today's appointment")
#     # employees_treeview.set('Aujordhui', 'Fait!' )
#     add_value_to_last_column(employees_treeview,"Fait!")
#     # add_value_to_last_column(treeview_search,'Fait!')
















# Start the tkinter event loop
window.mainloop()
