"""This is testing the email and smtplib module for Python
"""

import pyodbc
import datetime
import smtplib
from openpyxl import Workbook
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email import encoders


# Send email with Report
def send_email(file_attachment):
    """This function will send email with the attachment.
    It takes attachment file name as argument.
    """

    # Define email body
    body = "This is EA Monthly report. See attached file for Total Exam Volume for each customer."
    content = MIMEText(body, 'plain')
    
    # Open file attachment
    filename = file_attachment
    infile = open(filename, "rb")

    # Set up attachment to be send in email
    part = MIMEBase("application", "octet-stream")
    part.set_payload(infile.read())
    encoders.encode_base64(part)
    part.add_header("Content-Disposition", "attachment", filename=filename)

    msg = MIMEMultipart("alternative")

    # Define email recipients
    to_email = "nerminkekic@ge.com"
    from_email = "nerminkekic@ge.com"

    # Create email content
    msg["Subject"] = "EA Monthly Report"
    msg["From"] = from_email
    msg["To"] = to_email
    msg.attach(part)
    msg.attach(content)

    # Send email to SMTP server
    s = smtplib.SMTP("10.4.1.1", 25)
    s.sendmail(from_email, to_email, msg.as_string())
    s.close()

send_email(file)


