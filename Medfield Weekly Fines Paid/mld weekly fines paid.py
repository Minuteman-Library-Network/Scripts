#!/usr/bin/env python3

#Run in py313
"""
Jeremy Goldstein
Minuteman Library Network

Generates weekly report of fines paid totals
Report is then emailed to staff an Excel attachment
"""

import psycopg2
import xlsxwriter
import smtplib
import configparser
import os
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email.utils import formatdate
from email import encoders
from datetime import date
import traceback

# function takes a sql query as a parameter, connects to a database and returns the results
def run_query(query):
    # read config file with database login details
    config = configparser.ConfigParser()
    config.read("C:\\Scripts\\Creds\\config.ini")

    # Connecting to PostgreSQL database
    try:
        conn = psycopg2.connect(config["sql"]["connection_string"])
    except psycopg2.Error as e:
        print("Unable to connect to database: " + str(e))

    # Opening a session and querying the database
    cursor = conn.cursor()
    cursor.execute(query)
    # Storing the results in a variable. We'll use it later.
    rows = cursor.fetchall()
    # close database connection
    conn.close()
    # return variables containing query results and column headers
    return rows

# convert sql query results into formatted excel file
def excel_writer(query_results, excel_file):

    # Creating the Excel file for staff
    workbook = xlsxwriter.Workbook(excel_file, {'remove_timezone': True})
    worksheet = workbook.add_worksheet()

    # Formatting our Excel worksheet
    worksheet.set_landscape()
    worksheet.hide_gridlines(0)

    # Formatting Cells
    eformat= workbook.add_format({'text_wrap': True, 'valign': 'top'})
    eformatlabel= workbook.add_format({'text_wrap': True, 'valign': 'top', 'bold': True})
    eformat2= workbook.add_format({'num_format': 'mm/dd/yy'})

    # Setting the column widths
    worksheet.set_column(0,0,12.86)
    worksheet.set_column(1,1,8.43)
    worksheet.set_column(2,2,9.3)
    worksheet.set_column(3,3,12.71)
    worksheet.set_column(4,4,14.86)
    worksheet.set_column(5,5,10.86)
    worksheet.set_column(6,6,8.43)

    # Inserting a header
    worksheet.set_header('Medfield Weekly Fines Paid')

    # Adding column labels
    worksheet.write(0,0,'Date', eformatlabel)
    worksheet.write(0,1,'Overdue', eformatlabel)
    worksheet.write(0,2,'Lost Book', eformatlabel)
    worksheet.write(0,3,'Replacement', eformatlabel)
    worksheet.write(0,4,'Manual Charge', eformatlabel)
    worksheet.write(0,5,'Adjustment', eformatlabel)
    worksheet.write(0,6,'Total', eformatlabel)

    # Writing the report for staff to the Excel worksheet
    for rownum, row in enumerate(query_results):
        worksheet.write(rownum+1,0,row[0], eformat2)
        worksheet.write(rownum+1,1,row[1], eformat)
        worksheet.write(rownum+1,2,row[2], eformat)
        worksheet.write(rownum+1,3,row[3], eformat)
        worksheet.write(rownum+1,4,row[4], eformat)
        worksheet.write(rownum+1,5,row[5], eformat)
        worksheet.write(rownum+1,6,row[6], eformat)
    
    workbook.close()
    return excel_file

# function takes a file as a parameter and attaches that file to an outgoing email
def send_email(subject, message, attachment, recipient):
    # read config file with credentials for email account
    config = configparser.ConfigParser()
    config.read("C:\\Scripts\\Creds\\config.ini")
    # read config file with recipient list for email
    config_recipient = configparser.ConfigParser()
    config_recipient.read("C:\\Scripts\\Creds\\emails.ini")

    # These are variables for the email that will be sent, taken from .ini files referenced above
    emailhost = config["email"]["host"]
    emailuser = config["email"]["user"]
    emailpass = config["email"]["pw"]
    emailport = config["email"]["port"]
    emailfrom = config["email"]["sender"]
    # plain text of email message
    emailmessage = message

    # Creating the email message
    msg = MIMEMultipart()
    msg["From"] = emailfrom
    if type(recipient) is list:
        msg["To"] = ", ".join(recipient)
    else:
        msg["To"] = recipient
    msg["Date"] = formatdate(localtime=True)
    msg["Subject"] = subject
    msg.attach(MIMEText(emailmessage))
    part = MIMEBase("application", "octet-stream")
    part.set_payload(open(attachment, "rb").read())
    encoders.encode_base64(part)
    part.add_header(
        "Content-Disposition", "attachment; filename=%s" % attachment.rsplit("/", 1)[-1]
    )
    msg.attach(part)

    # Sending the email message
    smtp = smtplib.SMTP(emailhost, emailport)
    smtp.ehlo()
    smtp.starttls()
    smtp.login(emailuser, emailpass)
    smtp.sendmail(emailfrom, recipient, msg.as_string())
    smtp.quit()


# function constructs and sends outgoing email given a subject, a recipient and body text in both txt and html forms
def send_email_error(subject, message, recipient):
    # read config file with Sierra login credentials
    config = configparser.ConfigParser()
    config.read("C:\\Scripts\\Creds\\config.ini")

    # These are variables for the email that will be sent.
    # Make sure to use your own library's email server (emailhost)
    emailhost = config["email"]["host"]
    emailuser = config["email"]["user"]
    emailpass = config["email"]["pw"]
    emailport = config["email"]["port"]
    emailfrom = config["email"]["sender"]

    # Creating the email message
    msg = MIMEMultipart()
    emailmessage = message
    msg["From"] = emailfrom
    if type(recipient) is list:
        msg["To"] = ", ".join(recipient)
    else:
        msg["To"] = recipient
    msg["Date"] = formatdate(localtime=True)
    msg["Subject"] = subject
    msg.attach(MIMEText(emailmessage))

    # Sending the email message
    smtp = smtplib.SMTP(emailhost, emailport)
    # for Gmail connection used within Minuteman
    smtp.ehlo()
    smtp.starttls()
    smtp.login(emailuser, emailpass)
    smtp.sendmail(emailfrom, recipient, msg.as_string())
    smtp.quit()

def main():

    query = r"""
    SELECT 
      f.paid_date_gmt::DATE AS DATE,
      COALESCE(SUM(f.paid_now_amt) FILTER(WHERE f.charge_type_code IN ('2','6'))::MONEY,0.00::MONEY) AS overdue,
      COALESCE(SUM(f.paid_now_amt) FILTER(WHERE f.charge_type_code = '5')::MONEY,0.00::MONEY) AS lost_book,
      COALESCE(SUM(f.paid_now_amt) FILTER(WHERE f.charge_type_code = '3')::MONEY,0.00::MONEY) AS replacement,
      COALESCE(SUM(f.paid_now_amt) FILTER(WHERE f.charge_type_code = '1')::MONEY,0.00::MONEY) AS manual_charge,
      COALESCE(SUM(f.billing_fee_amt) FILTER(WHERE f.charge_type_code = '4')::MONEY,0.00::MONEY) AS adjustment,
      COALESCE(SUM(f.paid_now_amt) FILTER(WHERE f.charge_type_code BETWEEN '1' AND '6')::MONEY,0.00::MONEY) AS total

    FROM     sierra_view.fines_paid f

    WHERE (f.tty_num BETWEEN 501 AND 502 OR f.tty_num BETWEEN 504 AND 509)
      AND f.payment_status_code NOT IN ('0','3')
      AND f.paid_now_amt > 0
      AND f.paid_date_gmt::DATE >= CURRENT_DATE - INTERVAL '1 week'

    GROUP BY 1
    ORDER BY 1
    """

    query_results = run_query(query)

    # generate excel file from those query results
    excel_file = "/Scripts/Medfield Weekly Fines Paid/Temp Files/mld weekly fines paid{}.xlsx".format(date.today())
    excel_file = excel_writer(query_results, excel_file)

    # send email
    email_subject = 'MLD weekly fines paid'
    email_message = '''***This is an automated email***


The MLD weekly fines paid report has been attached.'''
	# read config file with recipient list for email
    config_recipient = configparser.ConfigParser()
    config_recipient.read("C:\\Scripts\\Creds\\emails.ini")
    email_recipient = config_recipient["medfield_weekly_fines_paid"]["recipients"].split()  
    send_email(email_subject, email_message, excel_file, email_recipient)

    os.remove(excel_file)


# run main function and send error email to admin of script encounters an error
if __name__ == "__main__":
    try:
        main()
    except Exception:
        # read config file with recipient list for email
        config_recipient = configparser.ConfigParser()
        config_recipient.read("C:\\Scripts\\Creds\\emails.ini")
        emailto = config_recipient["script_error"]["recipients"].split()

        # craft email subject and message containing error message details from traceback
        email_subject = "Medfield Weekly Fines Paid Script error"
        email_message = (
            "Your script failed with the following error:\n\n" + traceback.format_exc()
        )

        send_email_error(email_subject, email_message, emailto)
        raise
