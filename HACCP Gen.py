# -*- coding: utf-8 -*-
"""
Created on Fri Jun 14 10:22:02 2024

Convert Excel Workbooks into PDFs

@author: Waris Popal
"""
#Running this scrpt will convert all of the HACCP
#Files into pdfs and label the date it was generated

from win32com import client
import datetime
import os
import shutil

#Get current date
date_full = str(datetime.date.today())
year = date_full[0:4]
month = date_full[5:7]
date = date_full[8:10]
file_date = " " + month + "-" + date + "-" + year

#Path to store the edited logs (where the PDFs go)
excel_path = "C:/Perry Files/HACCP Logs/"

#Path to the edited logs for the day
#edited_logs_path = "C:\\Users\warisp897\\OneDrive - Virginia Tech\\HACCP LOGS\\"
edited_logs_path = "C:\\Users\\" + os.login + "\\OneDrive - Virginia Tech\\Filled HACCP Logs (FILL IN HERE)\\"

#Path of the clear logs that will be used for replacing
clear_logs_path = "C:\\Users\\" + os.login + "\\OneDrive - Virginia Tech\\Clear HACCP Logs (DO NOT EDIT)\\"

pdf_path = "C:\\Users\\" + os.login + "\\OneDrive - Virginia Tech\\HACCP PDFs\\"

#Open excel in python
excel = client.Dispatch("Excel.Application")

#Take each HACCP log name in the folder and export it to a pdf
try:
    #for file in os.listdir("C:\\Perry Files\\HACCP Logs"):
    for file in os.listdir(edited_logs_path):
        print(file)
        #Make sure it's an excel file
        if file.endswith(".xlsx"):
            
            #Remove the .xlsx extension
            file_mod = file.replace(".xlsx", "")
            
            #Open spreadsheet
            workbook = excel.Workbooks.Open(edited_logs_path + file_mod)
            
            #grab the first sheet
            worksheet = workbook.Worksheets[0]
            
            #Export sheet was pdf
            worksheet.ExportAsFixedFormat(0, pdf_path + file_mod + file_date)
            
            #Done working the excel spreadsheet, so 
            #close it to prevent permission issues
            workbook.Close(False)
            
            #Copy clear HACCP log and have it replace the filled one
            shutil.copyfile(clear_logs_path + file, edited_logs_path + file)
            
finally:
    #Close the excel application even if the initial code failed
    #(Leaving it open will prevent it from running)
    #workbook.Close(False)
    excel.Quit()


#If it doesn't run, go to task manager and click end task on Microsoft Excel
#(It will show up in "Background processes" after clicking "more details")