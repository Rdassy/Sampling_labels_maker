# Import Dependencies
import os
from pathlib import Path
import pandas as pd
import pythoncom
from sys import exit
import tkinter as tk
from tkinter import*
from tkinter import filedialog
from tkinter import messagebox
from tkinter import simpledialog
import win32com.client as win32


# Choose File and fit to merge template
# Pick samples list file
root = tk.Tk()
root.withdraw()
tk.messagebox.showinfo("Welcome!", "Please select your samples list file.")
root.call('wm', 'attributes', '.', '-topmost', True)
file_path = filedialog.askopenfilename()

try: 
    df1 = pd.read_excel(file_path)
    df1 = df1.fillna("N")
except Exception:
    tk.messagebox.showinfo("Error", "No file selected, please try again")
    exit()
    

# Ask for operator
application_window = tk.Tk()
application_window.withdraw()
operator_name = simpledialog.askstring("Input", "What customer/operator are these samples for? Leave blank to keep unpopulated.",
                                parent=application_window)
if operator_name != "":
    print("The operator for these labels will be ", operator_name)
else:
    print("Operator will be left blank.")
    operator_name = "_____________"

# Ask for operator
application_window2 = tk.Tk()
application_window2.withdraw()
sample_date = simpledialog.askstring("Input", "If all on the same day, enter sample date in MM/DD/YYYY format. Leave blank to keep unpopulated.",
                                parent=application_window2)
if sample_date != "":
    print("The date for these labels will be ", sample_date)
else:
    print("Date will be left blank.")
    sample_date = "__/__/____"


# Create destination df
df2 = pd.DataFrame(columns=['Number','Date','Producer','Location','Sample Point', "Analysis Type", 'Product'])

# File naming code
new_file_path = Path(file_path).with_suffix('')
labels_file_path = f'{new_file_path}_labels'
merge_file_path = f'{new_file_path}_merge.xlsx'
try:
    writer = pd.ExcelWriter(merge_file_path, engine='xlsxwriter')
    
# Take care of PermissionError: [Errno 13] 
except PermissionError:
    root.withdraw()
    messagebox.showinfo('Error', 'A merge dialog box is already open in Word. Please close it and try again.')
    exit()

# Get a list of columns for iteration loop
columns = list(df1)

# Iterate through rows and columns to find which samoples are needed depending on input Y or N
for row in df1.iterrows():
    for i in columns:
        if df1[i][row[0]] == "N":
            pass
        elif df1[i][row[0]] == "Y":
            df2.loc[len(df2.index)] = [len(df2.index)+1,
                                       sample_date,
                                       operator_name,
                                       df1["Location"][row[0]],
                                       "__",
                                       i ,
                                       df1["Product"][row[0]]]
        else:
            pass

# Populating product field with a blank for all analyses that are not residuals.
for row in df2.iterrows():
    if df2["Analysis Type"][row[0]] != "PO4":
        df2.loc[row[0]]["Product"] = ""


# Write data to an excel file
df2.to_excel(writer, sheet_name="Samples", index=False)

# Get workbook
workbook = writer.book
writer.close()


# Find Working Directory
working_directory = os.getcwd()


# Create a Word application instance
wordApp = win32.Dispatch('Word.Application')
wordApp.Visible = True

# Open Word Template + Open Data Source
template_filename = os.path.join(working_directory,"Labels_Template.docx")
sourceDoc = wordApp.Documents.Open(template_filename)
mail_merge = sourceDoc.MailMerge

# Message Box to click sheet
root.withdraw()
messagebox.showinfo('Action Required', 'Please Navigate to the Word Document and choose which sheet to merge.')

# Execute SQL code to prep for merge
try:
    mail_merge.OpenDataSource(
    Name:=merge_file_path,
    sqlstatement:="SELECT * FROM [Samples$]"
    )

    # Perform the Mail Merge
    mail_merge.Destination = 0
    mail_merge.Execute(False)

# Take care of com_error
except pythoncom.com_error:
    root.withdraw()
    messagebox.showinfo('Error', 'You did not select a sheet to merge. Please restart the program and try again.')
    exit()

# If no error, save Files in both Word Doc and PDF

else: 
    targetDoc = wordApp.ActiveDocument

    # Save Files in Word Doc and PDF

    targetDoc.SaveAs2(f'{labels_file_path}.docx', 16)
    targetDoc.ExportAsFixedFormat(labels_file_path, exportformat:=17)

    wordApp.Visible = True

    # Close target file
    targetDoc.Close(False)
    targetDoc = None

    sourceDoc.MailMerge.MainDocumentType = -1
    
    #  Completion Message Box
    tk.messagebox.showinfo("All done!", "Both PDF and Word Label files were created in the same folder as the samples list. Printing from PDF is ideal, please select 1-sided print from bypass tray. Happy Sampling!")
    exit()




