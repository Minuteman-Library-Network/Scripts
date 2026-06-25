#!/usr/bin/env python3

#Run in py313
"""
Jeremy Goldstein
Minuteman Library Network

Generates monthly report on renewals to patrons associated with Framingham
Report is produced as an Excel file, which is then emailed to staff.
"""

import psycopg2
import xlsxwriter
import os
import configparser
import smtplib
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

    workbook = xlsxwriter.Workbook(excel_file)
    worksheet = workbook.add_worksheet()


    #Formatting our Excel worksheet
    worksheet.set_landscape()
    worksheet.hide_gridlines(0)

    #Formatting Cells
    eformat= workbook.add_format({'text_wrap': True, 'valign': 'top'})
    eformatlabel= workbook.add_format({'text_wrap': True, 'valign': 'top', 'bold': True})


    # Setting the column widths
    worksheet.set_column(0,0,20.57)
    worksheet.set_column(1,1,10.14)
    worksheet.set_column(2,2,15.71)
    worksheet.set_column(3,3,19.57)

    #Inserting a header
    worksheet.set_header('Framingham Monthly Renewal Stats')

    # Adding column labels
    worksheet.write(0,0,'Application', eformatlabel)
    worksheet.write(0,1,'Count', eformatlabel)
    worksheet.write(0,2,'FPL_checkouts', eformatlabel)
    worksheet.write(0,3,'FPL_ptype_checkouts', eformatlabel)


    # Writing the report for staff to the Excel worksheet
    for rownum, row in enumerate(query_results):
        worksheet.write(rownum+1,0,row[0], eformat)
        worksheet.write(rownum+1,1,row[1], eformat)
        worksheet.write(rownum+1,2,row[2], eformat)
        worksheet.write(rownum+1,3,row[3], eformat)
    
    workbook.close()

# function takes a file as a parameter and attaches that file to an outgoing email
def send_email(subject, message, attachment):
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
    emailto = config_recipient["framingham_renewals"]["recipients"].split()
    # plain text of email message
    emailmessage = message

    # Creating the email message
    msg = MIMEMultipart()
    msg["From"] = emailfrom
    if type(emailto) is list:
        msg["To"] = ", ".join(emailto)
    else:
        msg["To"] = emailto
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
    smtp.sendmail(emailfrom, emailto, msg.as_string())
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
    # query to identify LoT transactions in the past month
    query = """
    SELECT
      c.application_name,
      COUNT(C.id),
      COUNT(C.id) FILTER (WHERE (C.loanrule_code_num BETWEEN '123' AND '133' OR C.loanrule_code_num BETWEEN '600' AND '608')) AS FPL_checkouts,
      COUNT(C.id) FILTER (WHERE p.ptype_code = '12') AS FPL_ptype_checkouts

    FROM sierra_view.circ_trans C
    JOIN sierra_view.patron_record p
      ON C.patron_record_id = p.id

    WHERE C.transaction_gmt >= NOW() - INTERVAL '1 month' AND C.op_code = 'r'
    GROUP BY 1
    ORDER BY 2 desc
    """

    query_results = run_query(query)

    # generate excel file from those query results
    excel_file = "/Scripts/Framingham Renewal stats/Temp Files/fpl monthly renewal stats{}.xlsx".format(date.today())
    excel_writer(query_results, excel_file)

    # send email
    email_subject = "FPL monthly renewal stats"
    email_message = """***This is an automated email***


The FPL monthly renewal stats report has been attached."""
    send_email(email_subject, email_message, excel_file)

    # delete local file
    os.remove(excel_file)


# run main function and send error email to admin of script encounters an error
if __name__ == "__main__":
    try:
        main()
    except Exception:
        # read config file with recipient list for email
        config_recipient = configparser.ConfigParser()
        config_recipient.read("C:\\Scripts\\Creds\\emails.ini")
        emailto = config_recipient["script_error_extended"]["recipients"].split()

        # craft email subject and message containing error message details from traceback
        email_subject = "Framingham Renewal Stats error"
        email_message = (
            "Your script failed with the following error:\n\n" + traceback.format_exc()
        )

        send_email_error(email_subject, email_message, emailto)
        raise