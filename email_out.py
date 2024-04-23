import xlwings as xw
import time
from win32com.client import Dispatch
import pandas as pd
import os
from glob import glob

email_file_path = fr"P:\PACS\Finance\Budgets\240222 - MASTER FORECAST TEMPLATE - 2024 Q2.xlsx"
forecast_template_path = fr"P:\PACS\Finance\Budgets\2024 Q2\Sent to Admins\Revise\*.xlsx"
epl = fr"EPL_adjust.csv"
df_a = pd.read_csv(epl, index_col=False)

"""EMAILS DATAFRAME"""
mail_df = pd.read_excel(email_file_path, sheet_name="Email List")
mailer = Dispatch("Outlook.Application")

time.sleep(2)

for facility in glob(forecast_template_path):
    try:
        file_name = os.path.basename(facility)
        # Remove file extension to get the Facility name
        file_name = os.path.splitext(file_name)[0]

        # Split Facility name based on hyphen ('-')
        file_name = file_name.split('-')

        # Use the first part as the complete Facility name
        file_name = file_name[0]

        msg = mailer.CreateItem(0)

        try:
            admin_email = mail_df.loc[mail_df['Facilities'] == file_name, "Admin_email"].iloc[0]
            admin_fname = mail_df.loc[mail_df['Facilities'] == file_name, "Admin_fname"].iloc[0]
            rdo_email = mail_df.loc[mail_df['Facilities'] == file_name, "RVPO"].iloc[0]
            msg.To = admin_email
            msg.CC = rdo_email
        except:
            admin_name = "", ""
            admin_email = mail_df.loc[mail_df['Facilities'] == file_name, "Admin_email"].iloc[0]
            pass
        if file_name not in df_a['Facility'].values:
            epl = 0
            pro_fees = 0
        else:
            df_a.loc[df_a['Facility'] == file_name, 'EPL'].values[0]
            epl = df_a.loc[df_a['Facility'] == file_name, 'EPL'].values[0]
            df_a.loc[df_a['Facility'] == file_name, 'Pro_Fees_Labor_Claim'].values[0]
            pro_fees = df_a.loc[df_a['Facility'] == file_name, 'Pro_Fees_Labor_Claim'].values[0]

        """PUT EMAIL TOGETHER"""
        msg.Subject = f"2024 Q2 Forecast Update-{file_name}"
        msg.Body = fr"""Hi {admin_fname},

    Attached is the forecast template file for your facility.  You will be updating your 2024 forecast with a focus on Q2.
    The historical financial statements used as the basis for these forecasts are October 1 2023 through February 29 2024. We have provided the forecasts preloaded with this data in an effort to give a reasonable starting point.
    Ultimately the administrator owns the forecast and is empowered to make necessary changes.


    Please submit: Friday the 22nd of March. If you are at any risk of a survey, please submit ASAP and we can help where haste is needed. This will allow us to have time to upload new data for your reporting and dashboard purposes before Q2 starts.


    *Please note, due to changes in outstanding litigation related to EPL (Employment Practices Liability), an adjustment needs to be included in the Professional Fees line of the forecast. Validate the amounts entered into this line, if the amount looks wrong, include {pro_fees} + {epl} + Historical_amount in cell E809 to be consistent with the reserve made by accounting for these claims. If you have any questions about this amount please reach out to Ryan or Kyle.

    The FP&A team is prepared to help you with your forecasting process, for questions and submission please contact: forecasts@pacs.com
    *****************
    forecasts@pacs.com
    *****************

    * We are also providing a demo for new admins or those with questions, the group session will be Thursday the 21st of March @10 am MST, and we can schedule a 1 on 1 call for more in-depth questions.
    https://teams.microsoft.com/l/meetup-join/19%3ameeting_YzAyNjA0N2UtNWMzNy00OTU3LWJlNTUtMjhjZDUzZTA3ODRi%40thread.v2/0?context=%7b%22Tid%22%3a%2293c1680b-e1a5-4495-a6ba-9d1d4ea523b0%22%2c%22Oid%22%3a%2291b0c145-022a-4358-9e62-c28c076c0d38%22%7d

    Thank you,

                """
        msg.Attachments.Add(Source=facility)
        # msg.Display(True)
        msg.Save()
        msg.Recipients.ResolveAll
        # msg.Send()
        print(f"Forecast ready: {facility}")
    except:
        print(f"Issue saving and emailing: {facility}")
