import win32com.client
import pandas as pd
from datetime import datetime

# Connect to Outlook
try:
    outlook = win32com.client.Dispatch("Outlook.Application")
    namespace = outlook.GetNamespace("MAPI")
except Exception as e:
    print("Error connecting to Outlook:", e)
    exit(1)

# Access the calendar folder
calendar = namespace.GetDefaultFolder(9)  # 9 corresponds to the calendar

# Define the start and end dates for the search
# Change dates to preferred range here:
# Format is: yyyy, mm, dd
start_date = datetime(2023, 7, 7)
end_date = datetime(2024, 7, 30)

# Restrict the calendar items to the defined date range
items = calendar.Items
items.IncludeRecurrences = True
items.Sort("[Start]")

# Restrict to items within the date range
restriction = f"[Start] >= '{start_date.strftime('%m/%d/%Y')}' AND [End] <= '{end_date.strftime('%m/%d/%Y')}'"
restricted_items = items.Restrict(restriction)

# Extract appointments with the title "WFH" and filter by date range
wfh_appointments = []
for item in restricted_items:
    appointment_date = item.Start.date()
    if item.Subject == "WFH" and start_date.date() <= appointment_date <= end_date.date():
        # Change number to however many hours you want to list against each Working From Home day
        #In this case, it is 7.5 hours
        wfh_appointments.append([appointment_date, 7.5])

# Create a DataFrame
df = pd.DataFrame(wfh_appointments, columns=["Date", "Hours"])

# Write to an Excel file
output_file = "WFH_Appointments.xlsx"
df.to_excel(output_file, index=False)
print(f"WFH appointments have been written to {output_file}")
