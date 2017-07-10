This script was created by  myself to help my financing group create billing invoices to our customers.
The script performs couple of key steps.

1. Connects to SQL Data Base serverand retrieves data for all customers DBs with SQL Query.
2. Formats that data so that its easy for finance group to use.
3. Saves data into a Excel Worksheet.
4. Sends the email to finance group with the Excel Worksheet attached.

The script is run every month, so that the invoices can be send out at the end of the month.

I have writen this script in a way that there is minimal to no need to maintain. Ex. If new customer are added,
script will be able to include them in the report without any need to make changes to script.

The script uses python 3.6, openpyxl module to manipulate Excel Worksheet, nd pyodbc module to connect and retrieve data from SQL server.
The script has been build to run on Windows server 2012 64bit and SQL 2014 environment.

The script has one main function called "ea_monthly_report()" where all the main steps above are performed.
That said, i have also created multiple functions that deal with specifics ex. "send_email(file_attachment)" that handles
sending of the email with attachment. This is to allow for easier maintenance and reuse of the code. I can take the "send_email(file_attachment)"
function and reuse it in other programs or scripts.
