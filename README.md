# wfh_log
Automatically log your working-from-home days into an Excel spreadsheet from your Outlook calendar

A python script that I created so that I can easily extract all the days that I worked from home for tax and other purposes.

# How to use:
1. Whenever you work from home, save an appointment on your Outlook calendar titled: WFH (I save it as a 30 minute appointment just before I start work)
2. Change start_date and end_date variables within the script to whatever date range you want to extract data for (see comments in code)
3. Change float number in wfh_appointments.append to however many hours you usually work from home (default set at 7.5)
4. Save and run code ---> creates Excel spreadsheet called "WFH_Appointments.xlsx" in local folder

