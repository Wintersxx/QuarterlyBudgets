import re
import duckdb
import pandas as pd
from glob import glob
import os
from openpyxl import Workbook
import datetime
import numpy as np
import threadpoolctl

# research_path = fr"P:\PACS\Finance\Budgets\{folder}\Consolidated\*.xlsx"
q3 = fr"P:\Finance\Budgets\2024 Q3\Sent to Admins\*.xlsx"

conn = duckdb.connect(database=':memory:', read_only=False)
conn.execute('INSTALL spatial;')
conn.execute('LOAD spatial;')

dfs = []

# Loop through each file
for file in glob(q3):
    query_str = fr"""SELECT * FROM st_read('{file}', layer='DW Upload');"""
    df = conn.execute(query_str).df()
    try:
        # Extract facility name
        facility_name = df.iloc[0, 0]
        facility_name = re.sub(r'[^\w\s]', '', facility_name)
        date_vals = df.iloc[4, 4:28].values
        print(date_vals)

        # Financial
        revenue = df.iloc[24, 4:28].values.astype(float)
        partB = df.iloc[18, 4:28].values.astype(float)
        otherRev = df.iloc[21, 4:28].values.astype(float)
        pro_fees = df.iloc[126, 4:28].values.astype(float)
        insurance_liability = df.iloc[130, 4:28].values.astype(float)
        NOI = df.iloc[164, 4:28].values.astype(float)
        prop_int = df.iloc[143, 4:28].values.astype(float)
        dep_a = df.iloc[139, 4:28].values.astype(float)
        bad_debt = df.iloc[146, 4:28].values.astype(float)
        EBITDA = np.add(NOI, np.add(prop_int, dep_a))
        # Census
        occupancy_p = df.iloc[244, 4:28].values
        medicare_census = df.iloc[181, 4:28].values
        medicaid_census = df.iloc[182, 4:28].values
        managed_census = df.iloc[184, 4:28].values
        skilled = (medicare_census.astype(float) + managed_census.astype(float))
        total_days = df.iloc[207, 4:28].values.astype(float)
        skilled_mix = (skilled / total_days).astype(float)
        # Labor
        agency = df.iloc[60, 4:28].values.astype(float)
        labor_expense = df.iloc[62, 4:28].values.astype(float)
        labor_rev = df.iloc[64, 4:28].values.astype(float)
        nhppd = df.iloc[317, 4:28].values.astype(float)
        # Therapy
        physical_therapy = df.iloc[150, 4:28].values.astype(float)
        occupation_therapy = df.iloc[151, 4:28].values.astype(float)
        speech_therapy = df.iloc[152, 4:28].values.astype(float)
        therapy_total = (physical_therapy + occupation_therapy + speech_therapy).astype(float)
        # ADC
        snfADC = df.iloc[219, 4:28].values.astype(float)
        medicareADC = df.iloc[211, 4:28].values.astype(float)
        medicaidADC = df.iloc[212, 4:28].values.astype(float)
        managedADC = df.iloc[214, 4:28].values.astype(float)

    except:
        pass
        print(facility_name, "broke")

    # Create DataFrames for each metric
    # Create DataFrames for each metric
    df_rev = pd.DataFrame({
        "Facility": [facility_name],
        "Metric": ["Total_Revenue"],
        "January": revenue[0], "February": revenue[1], "March": revenue[2],
        "April": revenue[3], "May": revenue[4], "June": revenue[5],
        "July": revenue[6], "August": revenue[7], "September": revenue[8],
        "October": revenue[9], "November": revenue[10], "December": revenue[11],
        "January_13": revenue[12], "February_13": revenue[13], "March_13": revenue[14],
        "April_13": revenue[15], "May_13": revenue[16], "June_13": revenue[17],
        "July_13": revenue[18], "August_13": revenue[19], "September_13": revenue[20],
        "October_13": revenue[21], "November_13": revenue[22], "December_13": revenue[23]
    })

    df_partB = pd.DataFrame({
        "Facility": [facility_name],
        "Metric": ["PartB"],
        "January": partB[0], "February": partB[1], "March": partB[2],
        "April": partB[3], "May": partB[4], "June": partB[5],
        "July": partB[6], "August": partB[7], "September": partB[8],
        "October": partB[9], "November": partB[10], "December": partB[11],
        "January_13": partB[12], "February_13": partB[13], "March_13": partB[14],
        "April_13": partB[15], "May_13": partB[16], "June_13": partB[17],
        "July_13": partB[18], "August_13": partB[19], "September_13": partB[20],
        "October_13": partB[21], "November_13": partB[22], "December_13": partB[23]
    })

    df_otherRev = pd.DataFrame({
        "Facility": [facility_name],
        "Metric": ["Other_Revenue"],
        "January": otherRev[0], "February": otherRev[1], "March": otherRev[2],
        "April": otherRev[3], "May": otherRev[4], "June": otherRev[5],
        "July": otherRev[6], "August": otherRev[7], "September": otherRev[8],
        "October": otherRev[9], "November": otherRev[10], "December": otherRev[11],
        "January_13": otherRev[12], "February_13": otherRev[13], "March_13": otherRev[14],
        "April_13": otherRev[15], "May_13": otherRev[16], "June_13": otherRev[17],
        "July_13": otherRev[18], "August_13": otherRev[19], "September_13": otherRev[20],
        "October_13": otherRev[21], "November_13": otherRev[22], "December_13": otherRev[23]
    })

    df_proFees = pd.DataFrame({
        "Facility": [facility_name],
        "Metric": ["ProFees"],
        "January": pro_fees[0], "February": pro_fees[1], "March": pro_fees[2],
        "April": pro_fees[3], "May": pro_fees[4], "June": pro_fees[5],
        "July": pro_fees[6], "August": pro_fees[7], "September": pro_fees[8],
        "October": pro_fees[9], "November": pro_fees[10], "December": pro_fees[11],
        "January_13": pro_fees[12], "February_13": pro_fees[13], "March_13": pro_fees[14],
        "April_13": pro_fees[15], "May_13": pro_fees[16], "June_13": pro_fees[17],
        "July_13": pro_fees[18], "August_13": pro_fees[19], "September_13": pro_fees[20],
        "October_13": pro_fees[21], "November_13": pro_fees[22], "December_13": pro_fees[23]
    })

    df_liability = pd.DataFrame({
        "Facility": [facility_name],
        "Metric": ["Liability"],
        "January": insurance_liability[0], "February": insurance_liability[1], "March": insurance_liability[2],
        "April": insurance_liability[3], "May": insurance_liability[4], "June": insurance_liability[5],
        "July": insurance_liability[6], "August": insurance_liability[7], "September": insurance_liability[8],
        "October": insurance_liability[9], "November": insurance_liability[10], "December": insurance_liability[11],
        "January_13": insurance_liability[12], "February_13": insurance_liability[13],
        "March_13": insurance_liability[14],
        "April_13": insurance_liability[15], "May_13": insurance_liability[16], "June_13": insurance_liability[17],
        "July_13": insurance_liability[18], "August_13": insurance_liability[19],
        "September_13": insurance_liability[20],
        "October_13": insurance_liability[21], "November_13": insurance_liability[22],
        "December_13": insurance_liability[23]
    })

    df_noi = pd.DataFrame({
        "Facility": [facility_name],
        "Metric": ["NOI"],
        "January": NOI[0], "February": NOI[1], "March": NOI[2],
        "April": NOI[3], "May": NOI[4], "June": NOI[5],
        "July": NOI[6], "August": NOI[7], "September": NOI[8],
        "October": NOI[9], "November": NOI[10], "December": NOI[11],
        "January_13": NOI[12], "February_13": NOI[13], "March_13": NOI[14],
        "April_13": NOI[15], "May_13": NOI[16], "June_13": NOI[17],
        "July_13": NOI[18], "August_13": NOI[19], "September_13": NOI[20],
        "October_13": NOI[21], "November_13": NOI[22], "December_13": NOI[23]
    })

    df_bad = pd.DataFrame({
        "Facility": [facility_name],
        "Metric": ["Bad_Debt"],
        "January": bad_debt[0], "February": bad_debt[1], "March": bad_debt[2],
        "April": bad_debt[3], "May": bad_debt[4], "June": bad_debt[5],
        "July": bad_debt[6], "August": bad_debt[7], "September": bad_debt[8],
        "October": bad_debt[9], "November": bad_debt[10], "December": bad_debt[11],
        "January_13": bad_debt[12], "February_13": bad_debt[13], "March_13": bad_debt[14],
        "April_13": bad_debt[15], "May_13": bad_debt[16], "June_13": bad_debt[17],
        "July_13": bad_debt[18], "August_13": bad_debt[19], "September_13": bad_debt[20],
        "October_13": bad_debt[21], "November_13": bad_debt[22], "December_13": bad_debt[23]
    })

    df_ebitda = pd.DataFrame({
        "Facility": [facility_name],
        "Metric": ["EBITDA"],
        "January": EBITDA[0], "February": EBITDA[1], "March": EBITDA[2],
        "April": EBITDA[3], "May": EBITDA[4], "June": EBITDA[5],
        "July": EBITDA[6], "August": EBITDA[7], "September": EBITDA[8],
        "October": EBITDA[9], "November": EBITDA[10], "December": EBITDA[11],
        "January_13": EBITDA[12], "February_13": EBITDA[13], "March_13": EBITDA[14],
        "April_13": EBITDA[15], "May_13": EBITDA[16], "June_13": EBITDA[17],
        "July_13": EBITDA[18], "August_13": EBITDA[19], "September_13": EBITDA[20],
        "October_13": EBITDA[21], "November_13": EBITDA[22], "December_13": EBITDA[23]
    })

    df_medicare_census = pd.DataFrame({
        "Facility": [facility_name],
        "Metric": ["Medicare_census"],
        "January": medicare_census[0], "February": medicare_census[1], "March": medicare_census[2],
        "April": medicare_census[3], "May": medicare_census[4], "June": medicare_census[5],
        "July": medicare_census[6], "August": medicare_census[7], "September": medicare_census[8],
        "October": medicare_census[9], "November": medicare_census[10], "December": medicare_census[11],
        "January_13": medicare_census[12], "February_13": medicare_census[13], "March_13": medicare_census[14],
        "April_13": medicare_census[15], "May_13": medicare_census[16], "June_13": medicare_census[17],
        "July_13": medicare_census[18], "August_13": medicare_census[19], "September_13": medicare_census[20],
        "October_13": medicare_census[21], "November_13": medicare_census[22], "December_13": medicare_census[23]
    })

    df_medicaid_census = pd.DataFrame({
        "Facility": [facility_name],
        "Metric": ["Medicaid_census"],
        "January": medicaid_census[0], "February": medicaid_census[1], "March": medicaid_census[2],
        "April": medicaid_census[3], "May": medicaid_census[4], "June": medicaid_census[5],
        "July": medicaid_census[6], "August": medicaid_census[7], "September": medicaid_census[8],
        "October": medicaid_census[9], "November": medicaid_census[10], "December": medicaid_census[11],
        "January_13": medicaid_census[12], "February_13": medicaid_census[13], "March_13": medicaid_census[14],
        "April_13": medicaid_census[15], "May_13": medicaid_census[16], "June_13": medicaid_census[17],
        "July_13": medicaid_census[18], "August_13": medicaid_census[19], "September_13": medicaid_census[20],
        "October_13": medicaid_census[21], "November_13": medicaid_census[22], "December_13": medicaid_census[23]
    })

    df_managed_census = pd.DataFrame({
        "Facility": [facility_name],
        "Metric": ["Managed_census"],
        "January": managed_census[0], "February": managed_census[1], "March": managed_census[2],
        "April": managed_census[3], "May": managed_census[4], "June": managed_census[5],
        "July": managed_census[6], "August": managed_census[7], "September": managed_census[8],
        "October": managed_census[9], "November": managed_census[10], "December": managed_census[11],
        "January_13": managed_census[12], "February_13": managed_census[13], "March_13": managed_census[14],
        "April_13": managed_census[15], "May_13": managed_census[16], "June_13": managed_census[17],
        "July_13": managed_census[18], "August_13": managed_census[19], "September_13": managed_census[20],
        "October_13": managed_census[21], "November_13": managed_census[22], "December_13": managed_census[23]
    })

    df_days = pd.DataFrame({
        "Facility": [facility_name],
        "Metric": ["Total_days"],
        "January": total_days[0], "February": total_days[1], "March": total_days[2],
        "April": total_days[3], "May": total_days[4], "June": total_days[5],
        "July": total_days[6], "August": total_days[7], "September": total_days[8],
        "October": total_days[9], "November": total_days[10], "December": total_days[11],
        "January_13": total_days[12], "February_13": total_days[13], "March_13": total_days[14],
        "April_13": total_days[15], "May_13": total_days[16], "June_13": total_days[17],
        "July_13": total_days[18], "August_13": total_days[19], "September_13": total_days[20],
        "October_13": total_days[21], "November_13": total_days[22], "December_13": total_days[23]
    })

    df_occupancy = pd.DataFrame({
        "Facility": [facility_name],
        "Metric": ["Occupancy_p"],
        "January": occupancy_p[0], "February": occupancy_p[1], "March": occupancy_p[2],
        "April": occupancy_p[3], "May": occupancy_p[4], "June": occupancy_p[5],
        "July": occupancy_p[6], "August": occupancy_p[7], "September": occupancy_p[8],
        "October": occupancy_p[9], "November": occupancy_p[10], "December": occupancy_p[11],
        "January_13": occupancy_p[12], "February_13": occupancy_p[13], "March_13": occupancy_p[14],
        "April_13": occupancy_p[15], "May_13": occupancy_p[16], "June_13": occupancy_p[17],
        "July_13": occupancy_p[18], "August_13": occupancy_p[19], "September_13": occupancy_p[20],
        "October_13": occupancy_p[21], "November_13": occupancy_p[22], "December_13": occupancy_p[23]
    })

    df_skilled_mix = pd.DataFrame({
        "Facility": [facility_name],
        "Metric": ["Skilled_mix"],
        "January": skilled_mix[0], "February": skilled_mix[1], "March": skilled_mix[2],
        "April": skilled_mix[3], "May": skilled_mix[4], "June": skilled_mix[5],
        "July": skilled_mix[6], "August": skilled_mix[7], "September": skilled_mix[8],
        "October": skilled_mix[9], "November": skilled_mix[10], "December": skilled_mix[11],
        "January_13": skilled_mix[12], "February_13": skilled_mix[13], "March_13": skilled_mix[14],
        "April_13": skilled_mix[15], "May_13": skilled_mix[16], "June_13": skilled_mix[17],
        "July_13": skilled_mix[18], "August_13": skilled_mix[19], "September_13": skilled_mix[20],
        "October_13": skilled_mix[21], "November_13": skilled_mix[22], "December_13": skilled_mix[23]
    })

    df_therapyTotal = pd.DataFrame({
        "Facility": [facility_name],
        "Metric": ["Therapy"],
        "January": therapy_total[0], "February": therapy_total[1], "March": therapy_total[2],
        "April": therapy_total[3], "May": therapy_total[4], "June": therapy_total[5],
        "July": therapy_total[6], "August": therapy_total[7], "September": therapy_total[8],
        "October": therapy_total[9], "November": therapy_total[10], "December": therapy_total[11],
        "January_13": therapy_total[12], "February_13": therapy_total[13], "March_13": therapy_total[14],
        "April_13": therapy_total[15], "May_13": therapy_total[16], "June_13": therapy_total[17],
        "July_13": therapy_total[18], "August_13": therapy_total[19], "September_13": therapy_total[20],
        "October_13": therapy_total[21], "November_13": therapy_total[22], "December_13": therapy_total[23]
    })

    df_agency = pd.DataFrame({
        "Facility": [facility_name],
        "Metric": ["Agency"],
        "January": agency[0], "February": agency[1], "March": agency[2],
        "April": agency[3], "May": agency[4], "June": agency[5],
        "July": agency[6], "August": agency[7], "September": agency[8],
        "October": agency[9], "November": agency[10], "December": agency[11],
        "January_13": agency[12], "February_13": agency[13], "March_13": agency[14],
        "April_13": agency[15], "May_13": agency[16], "June_13": agency[17],
        "July_13": agency[18], "August_13": agency[19], "September_13": agency[20],
        "October_13": agency[21], "November_13": agency[22], "December_13": agency[23]
    })

    df_laborExpense = pd.DataFrame({
        "Facility": [facility_name],
        "Metric": ["Labor_Expense"],
        "January": labor_expense[0], "February": labor_expense[1], "March": labor_expense[2],
        "April": labor_expense[3], "May": labor_expense[4], "June": labor_expense[5],
        "July": labor_expense[6], "August": labor_expense[7], "September": labor_expense[8],
        "October": labor_expense[9], "November": labor_expense[10], "December": labor_expense[11],
        "January_13": labor_expense[12], "February_13": labor_expense[13], "March_13": labor_expense[14],
        "April_13": labor_expense[15], "May_13": labor_expense[16], "June_13": labor_expense[17],
        "July_13": labor_expense[18], "August_13": labor_expense[19], "September_13": labor_expense[20],
        "October_13": labor_expense[21], "November_13": labor_expense[22], "December_13": labor_expense[23]
    })

    df_laborRev = pd.DataFrame({
        "Facility": [facility_name],
        "Metric": ["Labor_%_Rev"],
        "January": labor_rev[0], "February": labor_rev[1], "March": labor_rev[2],
        "April": labor_rev[3], "May": labor_rev[4], "June": labor_rev[5],
        "July": labor_rev[6], "August": labor_rev[7], "September": labor_rev[8],
        "October": labor_rev[9], "November": labor_rev[10], "December": labor_rev[11],
        "January_13": labor_rev[12], "February_13": labor_rev[13], "March_13": labor_rev[14],
        "April_13": labor_rev[15], "May_13": labor_rev[16], "June_13": labor_rev[17],
        "July_13": labor_rev[18], "August_13": labor_rev[19], "September_13": labor_rev[20],
        "October_13": labor_rev[21], "November_13": labor_rev[22], "December_13": labor_rev[23]
    })

    df_nhppd = pd.DataFrame({
        "Facility": [facility_name],
        "Metric": ["NHPPD"],
        "January": nhppd[0], "February": nhppd[1], "March": nhppd[2],
        "April": nhppd[3], "May": nhppd[4], "June": nhppd[5],
        "July": nhppd[6], "August": nhppd[7], "September": nhppd[8],
        "October": nhppd[9], "November": nhppd[10], "December": nhppd[11],
        "January_13": nhppd[12], "February_13": nhppd[13], "March_13": nhppd[14],
        "April_13": nhppd[15], "May_13": nhppd[16], "June_13": nhppd[17],
        "July_13": nhppd[18], "August_13": nhppd[19], "September_13": nhppd[20],
        "October_13": nhppd[21], "November_13": nhppd[22], "December_13": nhppd[23]
    })

    df_snfADC = pd.DataFrame({
        "Facility": [facility_name],
        "Metric": ["SnfADC"],
        "January": snfADC[0], "February": snfADC[1], "March": snfADC[2],
        "April": snfADC[3], "May": snfADC[4], "June": snfADC[5],
        "July": snfADC[6], "August": snfADC[7], "September": snfADC[8],
        "October": snfADC[9], "November": snfADC[10], "December": snfADC[11],
        "January_13": snfADC[12], "February_13": snfADC[13], "March_13": snfADC[14],
        "April_13": snfADC[15], "May_13": snfADC[16], "June_13": snfADC[17],
        "July_13": snfADC[18], "August_13": snfADC[19], "September_13": snfADC[20],
        "October_13": snfADC[21], "November_13": snfADC[22], "December_13": snfADC[23]
    })

    df_medicareADC = pd.DataFrame({
        "Facility": [facility_name],
        "Metric": ["MedicareADC"],
        "January": medicareADC[0], "February": medicareADC[1], "March": medicareADC[2],
        "April": medicareADC[3], "May": medicareADC[4], "June": medicareADC[5],
        "July": medicareADC[6], "August": medicareADC[7], "September": medicareADC[8],
        "October": medicareADC[9], "November": medicareADC[10], "December": medicareADC[11],
        "January_13": medicareADC[12], "February_13": medicareADC[13], "March_13": medicareADC[14],
        "April_13": medicareADC[15], "May_13": medicareADC[16], "June_13": medicareADC[17],
        "July_13": medicareADC[18], "August_13": medicareADC[19], "September_13": medicareADC[20],
        "October_13": medicareADC[21], "November_13": medicareADC[22], "December_13": medicareADC[23]
    })

    df_medicaidADC = pd.DataFrame({
        "Facility": [facility_name],
        "Metric": ["MedicaidADC"],
        "January": medicaidADC[0], "February": medicaidADC[1], "March": medicaidADC[2],
        "April": medicaidADC[3], "May": medicaidADC[4], "June": medicaidADC[5],
        "July": medicaidADC[6], "August": medicaidADC[7], "September": medicaidADC[8],
        "October": medicaidADC[9], "November": medicaidADC[10], "December": medicaidADC[11],
        "January_13": medicaidADC[12], "February_13": medicaidADC[13], "March_13": medicaidADC[14],
        "April_13": medicaidADC[15], "May_13": medicaidADC[16], "June_13": medicaidADC[17],
        "July_13": medicaidADC[18], "August_13": medicaidADC[19], "September_13": medicaidADC[20],
        "October_13": medicaidADC[21], "November_13": medicaidADC[22], "December_13": medicaidADC[23]
    })

    df_managedADC = pd.DataFrame({
        "Facility": [facility_name],
        "Metric": ["ManagedADC"],
        "January": managedADC[0], "February": managedADC[1], "March": managedADC[2],
        "April": managedADC[3], "May": managedADC[4], "June": managedADC[5],
        "July": managedADC[6], "August": managedADC[7], "September": managedADC[8],
        "October": managedADC[9], "November": managedADC[10], "December": managedADC[11],
        "January_13": managedADC[12], "February_13": managedADC[13], "March_13": managedADC[14],
        "April_13": managedADC[15], "May_13": managedADC[16], "June_13": managedADC[17],
        "July_13": managedADC[18], "August_13": managedADC[19], "September_13": managedADC[20],
        "October_13": managedADC[21], "November_13": managedADC[22], "December_13": managedADC[23]
    })

    # Append the DataFrames to the list
    dfs.extend([
    df_medicare_census, df_medicaid_census, df_managed_census, df_days, df_occupancy,
    df_medicareADC, df_medicaidADC, df_managedADC, df_snfADC,
    df_rev, df_partB, df_otherRev, df_agency, df_laborExpense, df_laborRev,
    df_therapyTotal, df_skilled_mix, df_liability, df_proFees, df_bad, df_noi, df_ebitda, df_nhppd])

    print(facility_name)

# Concatenate all DataFrames in the list
df_result = pd.concat(dfs, ignore_index=True)
# Rename columns in the DataFrame
rename_mapping = {col: date_vals[i] for i, col in enumerate(df_result.columns[2:27])}
df_result.rename(columns=rename_mapping, inplace=True)
df_result['Unique_Metric'] = df_result['Facility'] + df_result['Metric']
df_result = df_result[['Facility', 'Unique_Metric', 'Metric'] + df_result.columns.difference(['Facility', 'Unique_Metric', 'Metric']).tolist()]


# Save the result to a CSV file
df_result.to_csv("Q3_KPI_prior-send.csv", index=False)





