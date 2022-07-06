from glob import glob
import openpyxl as oxl
import pandas as pd
import xlwings
import xlwings as xw
import sys
import time

app = xw.App(add_book=False)
xlwings.App.display_alerts = False

print("To access Folder")
folder = str(input("Enter \"YYYY Quarter #\": "))

"""Start date needs to equal the input"""

print("To verify budget integrity")
start_date = str(input("Start date for this budget period \"YYYY-MM-DD\": "))
budgets_sent = int(input("How many budgets were sent out?: "))
print("*" * 60)


path = fr"P:\PACS\Finance\Budgets\{folder}\Received\*.xlsx"
final_wb = xw.Book()

try:
    x = 2
    for file in glob(path):
        wb = xw.Book(file, update_links=False)
        full_budget = wb.sheets("RPT - ALL Lines")                  # open tab
        facility_info = wb.sheets("FACILITY INFO")
        main_page = wb.sheets("BUDGET WORKSHEET")
    print(path)
except TimeoutError:
    print(path, "File error")


dates = full_budget.range("C5:O5")  # pull all dates
noi = full_budget.range("C184:O184")                            # NOI for all dates
budget_start_date = facility_info.range("B15")
bed_ct = facility_info.range("B10")
occupancy_rate = main_page.range("J2")

Facility_name = wb.sheets("FACILITY INFO").range("B7").value    # pull Facility
NOI = wb.sheets("BUDGET WORKSHEET").range("I3").value           # pull NOI value

final_wb.sheets[0].range(f"P{x}").value = budget_start_date.value
final_wb.sheets[0].range(f"Q{x}").value = bed_ct.value
final_wb.sheets[0].range(f"R{x}").value = occupancy_rate.value

final_wb.sheets[0].range("C1:O1").value = dates.value
final_wb.sheets[0].range(f"C{x}:O{x}").value = noi.value        # drop in dates

final_wb.sheets[0].range(f"A{x}").value = Facility_name         # drop in Facility name in new workbook
final_wb.sheets[0].range(f"B{x}").value = NOI                   # drop in NOI value in workbook
x += 1
wb.close()

final_wb.sheets[0].range(f"a1").value = 'Facility'
final_wb.sheets[0].range(f"b1").value = 'NOI'
final_wb.sheets[0].range(f"p1").value = 'Budget_Start_Date'
final_wb.sheets[0].range(f"q1").value = 'Bed_Count'
final_wb.sheets[0].range(f"r1").value = 'Occupancy_Rate'

final_wb.save(fr"P:\PACS\Finance\Budgets\{folder}\budgets checked.xlsx")

# Checker to indicate early on if anything is amiss

if x-2 != budgets_sent:
    if x-2 < budgets_sent:
        print("Heads up, missing budgets from Admins")
    elif x-2 > budgets_sent:
        print("Something is off on the sent/receive count")
else:
    print("All budgets are accounted for.", x-2, "budgets received")
print("*" * 60)


"""PANDAS SECTION"""
facility_path = r"P:\PACS\Finance\Budgets\---Budget Tools and Automation Folder---\FacilityList w ID.xlsx"

df = pd.read_excel(facility_path)
df = df['Facility']

budgets_checked = fr"P:\PACS\Finance\Budgets\{folder}\budgets checked.xlsx"

df1 = pd.read_excel(budgets_checked)
df1 = df1['Facility']
difference = list(set(df) - set(df1))
difference.sort()
expt = pd.DataFrame(difference)
expt.to_csv(fr"P:\PACS\Finance\Budgets\{folder}\budgets_missing_output.csv")

"""PANDAS SECTION CONTINUED"""
dfb = pd.read_excel(budgets_checked)
dfb['Budget_Start_Date'] = dfb['Budget_Start_Date'].astype(str)
dff = dfb['Facility']
x = 0

""""Are there any issues with the upload dates?"""
for rows in dfb['Budget_Start_Date']:
    if rows != start_date:
        print("ERROR:", dff[x])
        with open(fr"P:\PACS\Finance\Budgets\{folder}\budget_receiver_errors.txt", 'a') as f:
            f.write(dff[x] + " has a budget date error\n")
    x += 1
print("*" * 60)

print()
print(time.process_time(), "minutes")
print(time.perf_counter(), "minutes")
