#!/usr/bin/env python
# coding: utf-8

# ## Create GUI using Tkinter

# In[2]:


from pathlib import Path
import pandas as pd
import win32com.client as win32
import openpyxl
import os
import shutil
from tkinter import *
from tkinter import filedialog

window = Tk(className='Tk Split Excel by Vendor and Send Email')
window.geometry("800x300")

def get_email_file_path():
    global email_file_path
    # Open and then return email file path
    email_file_path = filedialog.askopenfilename(title = "Select A file", filetype = (("xlsx", "*.xlsx"), ("csv", "*.csv")))
    label1 = Label(window, text = "Email File path: " + email_file_path).pack()
    
# Create a button to search the email excel file
open_email_file_button = Button(window, text="Open Email File", command = get_email_file_path).pack()


def get_excel_file_path():
    global file_path
    # Open and then return file path
    file_path = filedialog.askopenfilename(title = "Select A file", filetype = (("xlsx", "*.xlsx"), ("csv", "*.csv")))
    label1 = Label(window, text = "Report File path: " + file_path).pack()
    
# Create a button to search the excel file
open_file_button = Button(window, text="Open Excel File", command = get_excel_file_path).pack()


def extract_data_create_excel_files():
    global path
    directory = "OOReport Output Files"
    parent_dir = r'C:\Users\admin\Desktop\Deleted\Email Automation'
    path = os.path.join(parent_dir, directory)
    
    # Delete "OOReport Output Files" folder if exist
    if os.path.exists(path):
        shutil.rmtree(path)
        
    # Create "OOReport Output Files" folder
    isExist = os.path.exists(path)
    if not isExist:
        os.mkdir(path)
        
    # Load excel data into dataframe
    # excel_file_name = 'OOReport 11.14.22.xlsx'
    # data =pd.read_excel(r'C:\Users\admin\Desktop\Deleted\Email Automation\OOReport 11.14.22.xlsx')
    data = pd.read_excel(file_path)
        
    # Get unique values from a specific column that we want to separate
    column_name = 'Vendor'
    unique_values = data[column_name].unique()
    
    # Filter the dataframe and export to Excel file

    for unique_value in unique_values:
    #if os.path.isfile(f"{unique_value}.xlsx"):
        wb = openpyxl.Workbook()
        dest_filename = f"{unique_value}.xlsx"
        wb.save(os.path.join(path, dest_filename))
        output_excel_path = os.path.join(path, dest_filename)
        # data_output = data.query(f"{column_name} == @unique_value") 
        data_output = data.loc[data[column_name] == unique_value]
        data_output.to_excel(excel_writer = output_excel_path, index=False)
        
        # Print label to show process complete
    label2 = Label(window, text = "Process Complete. Directory of Extracted Excel Files:").pack()
    label3 = Text(window, height=1, borderwidth=0)
    label3.insert(1.0, path)
    label3.pack()
    
    
extract_create_button = Button(window, text='Extract data by Vendor and then create corresponding Excel file',                  command=extract_data_create_excel_files).pack()       

# Iterate over email distribution list & send emails via Outlook
def SendEmail():
    # Load email info into dataframe
    email_info = pd.read_excel(email_file_path)
    outlook = win32.Dispatch('outlook.application')
    for index, row in email_info.iterrows():
        mail = outlook.CreateItem(0)
        mail.To = row["Email"]
        mail.CC = row["CC"]
        mail.Subject = f"OOReport for: {row['Vendor']}"
        mail.Body = f"""Hi {row['First Name']}

Please find the attached report for {row['Vendor']}.

Thanks,
xxxxxxx
xxxxxxxx
yuyyyyyyy
zzzzzz

"""
        excel_file = f"{row['Vendor']}.xlsx"
        attached_excel_path = os.path.join(path, excel_file)
        mail.Attachments.Add(Source=attached_excel_path)

        mail.Display()
        # mail.Send()
    label = Label(window, text="Send Email(s) Successfully.")
send_email_button = Button(window, text='Send Email with Attachments')
send_email_button.config(command=SendEmail) # Perform the function def click()
send_email_button.pack()


window.mainloop()
#print(file_path)


# In[ ]:




