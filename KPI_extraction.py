import re
import duckdb
import pandas as pd
from glob import glob
import os
from openpyxl import Workbook
import datetime
import numpy as np



##############
print("To access Folder")
folder = str(input("Enter \"YYYY Quarter #\": "))

"""Start date needs to equal the input"""

print("To verify forecast integrity")
start_date = str(input("Start date for this forecast period \"YYYY-MM-DD\": "))

print("*" * 40)

##############
folder_path = fr"P:\PACS\Finance\Budgets\{folder}"
master_xl_file = fr"{folder}_KPI.xlsx"
# research_path = fr"P:\PACS\Finance\Budgets\{folder}\Consolidated\*.xlsx"
research_path = fr"P:\PACS\Finance\Budgets\2024 Q1\Received - Adjusted\Acquisitions\PM7 240201\*.xlsx"

conn = duckdb.connect(database=':memory:', read_only=False)
conn.execute('INSTALL spatial;')
conn.execute('LOAD spatial;')

dfs=[]

# Loop through each file
for file in glob(research_path):
    try:
        query_str = fr"""SELECT * FROM st_read('{file}', layer='REPORTING');"""
        df = conn.query(query_str).df()

        # Extract facility name
        facility_name = df.iloc[0, 0]
        facility_name = re.sub(r'[^\w\s]', '', facility_name)

        # Continue processing the file
        # Add your processing logic here

    except duckdb.duckdb.IOException as e:
        if str(e).startswith("GDAL Error (4):"):
            print(f"Skipping file '{file}' due to unsupported format.")
            continue
        else:
            # Handle other IO exceptions
            print(f"Error processing file '{file}': {e}")
            continue

    # Extract data for each metric
    NOI = df.iloc[168, 2:14].values.astype(float)
    prop_int = df.iloc[147, 2:14].values.astype(float)
    dep_a = df.iloc[143, 2:14].values.astype(float)
    # Data
    EBITDA = np.add(NOI, np.add(prop_int, dep_a))

    occupancy_p = df.iloc[248, 2:14].values
    medicare = df.iloc[185, 2:14].values
    managed = df.iloc[188, 2:14].values
    skilled = (medicare.astype(float) + managed.astype(float))
    total_days = df.iloc[211, 2:14].values.astype(float)
    skilled_mix = (skilled / total_days).astype(float)
    labor_rev = df.iloc[68, 2:14].values.astype(float)
    nhppd = df.iloc[321, 2:14].values.astype(float)

    # Create DataFrames for each metric
    df_noi = pd.DataFrame({"Facility": [facility_name], "Metric": ["NOI"],
                           "January": NOI[0], "February": NOI[1], "March": NOI[2],
                           "April": NOI[3], "May": NOI[4], "June": NOI[5],
                           "July": NOI[6], "August": NOI[7], "September": NOI[8],
                           "October": NOI[9], "November": NOI[10], "December": NOI[11]})

    df_ebitda = pd.DataFrame({"Facility": [facility_name], "Metric": ["EBITDA"],
                              "January": EBITDA[0], "February": EBITDA[1], "March": EBITDA[2],
                              "April": EBITDA[3], "May": EBITDA[4], "June": EBITDA[5],
                              "July": EBITDA[6], "August": EBITDA[7], "September": EBITDA[8],
                              "October": EBITDA[9], "November": EBITDA[10], "December": EBITDA[11]})

    df_occupancy = pd.DataFrame({"Facility": [facility_name], "Metric": ["occupancy_p"],
                                 "January": occupancy_p[0], "February": occupancy_p[1], "March": occupancy_p[2],
                                 "April": occupancy_p[3], "May": occupancy_p[4], "June": occupancy_p[5],
                                 "July": occupancy_p[6], "August": occupancy_p[7], "September": occupancy_p[8],
                                 "October": occupancy_p[9], "November": occupancy_p[10], "December": occupancy_p[11]})

    df_skilled_mix = pd.DataFrame({"Facility": [facility_name], "Metric": ["skilled_mix"],
                                   "January": skilled_mix[0], "February": skilled_mix[1], "March": skilled_mix[2],
                                   "April": skilled_mix[3], "May": skilled_mix[4], "June": skilled_mix[5],
                                   "July": skilled_mix[6], "August": skilled_mix[7], "September": skilled_mix[8],
                                   "October": skilled_mix[9], "November": skilled_mix[10], "December": skilled_mix[11]})

    df_labor_rev = pd.DataFrame({"Facility": [facility_name], "Metric": ["labor_rev"],
                                 "January": labor_rev[0], "February": labor_rev[1], "March": labor_rev[2],
                                 "April": labor_rev[3], "May": labor_rev[4], "June": labor_rev[5],
                                 "July": labor_rev[6], "August": labor_rev[7], "September": labor_rev[8],
                                 "October": labor_rev[9], "November": labor_rev[10], "December": labor_rev[11]})

    df_nhppd = pd.DataFrame({"Facility": [facility_name], "Metric": ["nhppd"],
                             "January": nhppd[0], "February": nhppd[1], "March": nhppd[2],
                             "April": nhppd[3], "May": nhppd[4], "June": nhppd[5],
                             "July": nhppd[6], "August": nhppd[7], "September": nhppd[8],
                             "October": nhppd[9], "November": nhppd[10], "December": nhppd[11]})

    # Append the DataFrames to the list
    dfs.extend([df_noi, df_ebitda, df_occupancy, df_skilled_mix, df_labor_rev, df_nhppd])
    print(facility_name)

# Concatenate all DataFrames in the list
df_result = pd.concat(dfs, ignore_index=True)

# Save the result to a CSV file
df_result.to_csv("Q1_acq.csv", index=False)

#
# missing = pd.read_csv(fr"C:\Users\kyle.anderson\pythonProject\Budget_tools\Facility_list_receivedbudgets.csv",index_col=False)
#
# outstanding = []
# for i in df_result['Facility']:
#     try:
#         # Check if 'i' is present in the 'Facility' column of the missing DataFrame
#         missing.loc[missing['Facility'] == i].iloc[0]
#     except IndexError:
#         # If 'i' is not found, append it to the outstanding list
#         outstanding.append(i)
#
# outstanding_df = pd.DataFrame({'Facility': outstanding})
# date = datetime.date.today()
# outstanding_df.to_csv(fr'{folder_path}\Missing_facilities_{date}.csv', index=False)
#
#
# print('done')
#


