import tkinter as tk
from tkinter import ttk
import ttkbootstrap as ttkb
from ttkbootstrap import *
import pandas as pd

# Create tkinter window
window = tk.Tk()
window.title("Employee Attendance Management System")
window.geometry("800x600")

# Apply ttkbootstrap style
style = ttkb.Style(theme="cosmo")

# Function to handle search button click
def search_employee():

    # Get the Matricule (ID) entered by the user
    matricule = matricule_entry.get()
    if matricule:
        # Here you can implement the logic to search for the employee in your database or data source
        # For demonstration purposes, let's just display a message
        result_label.config(text=f"Searching for employee with Matricule {matricule}")

        # Read data from Excel file
        # df = pd.read_excel("employees.xlsx")

        # Filter data based on Matricule
        filtered_df = df[df['Matricule'] == int(matricule)]

        # Display filtered data in Treeview
        display_employees(filtered_df)
        return
    result_label.config(text=f"Enter a valid Number!")


# Function to display employees in Treeview
def display_employees(df):
    # Clear existing items in Treeview
    for item in employees_treeview.get_children():
        employees_treeview.delete(item)
    # Display data in Treeview
    for index, row in df.iterrows():
        employees_treeview.insert("", "end", iid=index, text=index+1 , values= row.tolist())

# Function to handle reset button click
def reset_filter():
    # Clear Matricule entry
    matricule_entry.delete(0, tk.END)

    # Read data from Excel file
    df = pd.read_excel("employees.xlsx")

    # Display all employees in Treeview
    display_employees(df)

    # Clear the result label
    result_label.config(text='')


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
employees_treeview = ttk.Treeview(window, columns=["Name", "Matricule", "Department"])
employees_treeview.heading("#0", text="Index", anchor=tk.W)
employees_treeview.heading("Name", text="Name", anchor=tk.W)
employees_treeview.heading("Matricule", text="Matricule", anchor=tk.CENTER)
employees_treeview.heading("Department", text="Department", anchor=tk.W)

# Adjust column widths and alignments
employees_treeview.column("#0", width=50,minwidth=50, anchor=tk.W)
employees_treeview.column("Name", width=150, anchor=tk.W)
employees_treeview.column("Matricule", width=100, anchor=tk.CENTER)
employees_treeview.column("Department", width=200, anchor=tk.W)
# Set up Treeview style
style = ttk.Style()
# style.configure("Treeview", anchor=tk.W, font=("Arial", 10, "bold"))

# Pack Treeview
employees_treeview.pack(pady=20)



df = pd.read_excel("employees.xlsx")
display_employees(df)

# Start the tkinter event loop
window.mainloop()
