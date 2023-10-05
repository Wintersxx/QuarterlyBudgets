from glob import glob
import pandas as pd
import xlwings as xw
import time

folder = input("Enter \"YYYY Quarter #\": ")
start_date = input("Start date for this budget period \"YYYY-MM-DD\": ")
budgets_sent = int(input("How many budgets were sent out?: "))
path = fr"P:\PACS\Finance\Budgets\{folder}\Received\Uploaded\*.xlsx"
final_wb = xw.Book()

# Initialize the row counter
x = 2

# Open the final workbook outside the loop
# final_wb.sheets[0].range("A1:S1").value = ['Facility', 'NOI', 'Budget_Start_Date', 'Bed_Count', 'Occupancy_Rate',
#                                            'Fee%']

for file in glob(path):
    wb = xw.Book(file, update_links=False)
    full_budget = wb.sheets["RPT - ALL Lines"]
    facility_info = wb.sheets["FACILITY INFO"]

    try:
        main_page = wb.sheets["FORECAST WORKSHEET"]
    except:
        main_page = wb.sheets["BUDGET WORKSHEET"]

    # Extract data from Excel
    dates = full_budget.range("C5:O5").value
    noi = full_budget.range("C180:N180")
    budget_start_date = facility_info.range("B15").value
    bed_count = facility_info.range("B10").value
    occupancy_rate = main_page.range("J2").value
    fee_amount = main_page.range("E833").value
    facility_name = facility_info.range("B7").value
    noi_value = main_page.range("I3").value

    # Write data to final workbook
    final_wb.sheets[0].range(f"A{x}").value = facility_name
    final_wb.sheets[0].range(f"B{x}").value = noi_value
    final_wb.sheets[0].range(f"P{x}").value = budget_start_date
    final_wb.sheets[0].range(f"Q{x}").value = bed_count
    final_wb.sheets[0].range(f"R{x}").value = occupancy_rate
    final_wb.sheets[0].range(f"S{x}").value = fee_amount
    final_wb.sheets[0].range(f"C1:N1").value = dates
    final_wb.sheets[0].range(f"C{x}: N{x}").value = noi.value
    final_wb.sheets[0].range(f"a1").value = 'Facility'
    final_wb.sheets[0].range(f"b1").value = 'NOI'
    final_wb.sheets[0].range(f"p1").value = 'Budget_Start_Date'
    final_wb.sheets[0].range(f"q1").value = 'Bed_Count'
    final_wb.sheets[0].range(f"r1").value = 'Occupancy_Rate'
    final_wb.sheets[0].range(f"S1").value = 'Fee%'

    # Close the current workbook
    wb.close()

    x += 1

# Save and close the final workbook
final_wb.save(fr"P:\PACS\Finance\Budgets\{folder}\budgets checked.xlsx")
final_wb.close()

# Check budgets
if x - 2 != budgets_sent:
    if x - 2 < budgets_sent:
        print("Heads up, missing budgets from Admins")
    elif x - 2 > budgets_sent:
        print("Something is off on the sent/receive count")
else:
    print("All budgets are accounted for.", x - 2, "budgets received")

# Pandas section (you can optimize this further if needed)
facility_path = fr"Facility_list.csv"
df = pd.read_csv(facility_path)['Facility']
budgets_checked = fr"P:\PACS\Finance\Budgets\{folder}\budgets checked.xlsx"
df1 = pd.read_excel(budgets_checked)['Facility']
difference = list(set(df) - set(df1))
difference.sort()
expt = pd.DataFrame(difference)
expt.to_csv(fr"P:\PACS\Finance\Budgets\{folder}\budgets_missing_output.csv")

# Check for budget date errors
dfb = pd.read_excel(budgets_checked)
dfb['Budget_Start_Date'] = dfb['Budget_Start_Date'].astype(str)
dff = dfb['Facility']
x = 0

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
