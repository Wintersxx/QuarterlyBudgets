import os
import json
import win32com.client
import datetime
import db_connection
from glob import glob
import pathlib

"""
This script imports all budget files located in the directory 'folder'
It opens the excel file and runs the budget import macro from the excel add-in 
This script will update the budget file before upload to make sure only 
applicable months (columns) are being uploaded from the DW Upload tab.
If the date is less than the earliest_date variable, it will not upload that column.
The add-in ensures that columns with zeroes do not get uploaded
"""

# Full path to the add-in file
# add_in_path = fr"C:\Users\kyle.anderson\AppData\Roaming\Microsoft\AddIns\Budget Import Plugin.xlam"
add_in_path = fr"C:\Users\kyle.anderson\AppData\Roaming\Microsoft\AddIns\PACS_Apex_AddIn.xlam"

# folder = fr"P:\PACS\Finance\Budgets\2024 Q1\Dericks2.5Request\*.xlsx"
folder = fr"P:\PACS\Finance\Budgets\2024 Q2\Deloitte\*.xlsx"


# APEX database parameters
apex_db_provider = "SQLOLEDB.1"
apex_db_server = "PACS-AZURE-APEX"
# apex_db_server = "20.3.147.250"
apex_db_database = "PACS"


# Declare the list of file names that did not run and set it blank to start the script
MissingFileList = []

"""NEW BLOCK OF CODE"""
earliest_date = datetime.date(2024, 1, 1)  # UPDATE IF YOU DON'T WANT MONTHS PRIOR UPDATED (NOT REQUIRED)
xl = win32com.client.Dispatch('Excel.Application')
xl.Visible = False
xl.ScreenUpdating = False
xl.Interactive = False

for file in glob(folder):
    if file.endswith(".xlsx"):
        wb = xl.Workbooks.Open(file, UpdateLinks=False)

        upload_sht = wb.Worksheets('DW Upload')

        """SECTION TO ADD DOUBLE CHECKING NAME AGAINST DW TABLE"""
        # upload_name = wb.Worksheets('FACILITY INFO').Range("B7").Value
        dw_name = wb.Worksheets('IS-DW').Range("M8").Value

        building = db_connection.db_get_facility_name(dw_name)
        if building is not False:
            wb.Worksheets('DW Upload').Range("A1").Value = building

            """UPLOAD BUDGET SECTION"""
            for i in range(5, 65):  # IDENTIFY COLUMNS IN BUDGET THAT ARE EARLIER THAN EARLIEST_DATE
                cell_date = upload_sht.Cells(5, i).Value
                cell_date = cell_date.date()
                if cell_date < earliest_date:
                    upload_sht.Cells(4, i).Value = "FALSE"
            """IMPORT BUDGET"""
            if xl.AddIns.Add(add_in_path).Installed:
                xl.AddIns.Add(add_in_path).Installed = True
                xl.Workbooks.Open(add_in_path)

            # 1 minute per file
            # for ia in range(1, xl.AddIns.Count):
            #     title = xl.AddIns(ia).Title
            #     if "Budget Import Plugin" in xl.AddIns(ia).Title:
            #         wb.Application.Run('Plugin.SendBudget')  # RUN THE SUBMIT BUDGET MACRO
            #         with open('budget_submit.txt', 'a') as f:
            #             f.write(f"{dw_name} budget submitted as {building} on {datetime.datetime.now()} \n")
            #         break

            # 3 minute per file
            for ib in range(1, xl.AddIns.Count):
                title = xl.AddIns(ib).Title
                if "Pacs_Apex_Addin" in xl.AddIns(ib).Title:
                    # IF THE FACILITY EXISTS IN THE DATABASE
                    DoesTheFacilityExist = wb.Application.Run('PACS_Upload.DoesTheFacilityExist', dw_name,
                                                              apex_db_provider, apex_db_server, apex_db_database)

                    if DoesTheFacilityExist == -1:
                        MissingFileList.append(file)
                    else:
                        # OTHERWISE, IF THE FACILITY EXISTS, START THE UPLOAD IN THE DATABASE
                        wb.Application.Run('PACS_Upload.UploadFacilityDetails', dw_name, apex_db_provider,
                                           apex_db_server, apex_db_database)
                        with open('budget_submit.txt', 'a') as f:
                            f.write(f"{dw_name} budget submitted as {building} on {datetime.datetime.now()} \n")
                        break
            wb.Close(False)
        else:
            with open('budget_submit.txt', 'a') as f:
                f.write(
                    f"{dw_name} budget could not be submitted.  A name could not be matched to the DW on {datetime.datetime.now()}")
                f.write("\n")
            wb.Close(False)

xl.Quit()  # Comment this out if your excel script closes
del xl

# Create the message with two possible reasons
message = "The list below includes all files that were NOT uploaded:\n1. Please check if the facility name on 'IS DW' M8 exists in the database""\n2. Please check if the facility name on 'IS DW' M8 is spelled correctly to match the facility in the database.\n\n"

# Add the list of missing facilities to the message
message += "Missing Files:\n"

for file in MissingFileList:
    message += f"- {file}\n"

# Print the message
print(message)
