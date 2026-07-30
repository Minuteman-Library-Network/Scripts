#!/usr/bin/env python3

# Run in py313

"""
Create and email a list of open orders for Needham

Author: Jeremy Goldstein
Contact Info: jgoldstein@minlib.net
"""

import psycopg2
import xlsxwriter
import smtplib
import os
import configparser
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
    #Creating the Excel file for staff
    workbook = xlsxwriter.Workbook(excel_file)
    worksheet = workbook.add_worksheet()


    # Formatting our Excel worksheet
    worksheet.set_landscape()
    worksheet.hide_gridlines(0)

    # Formatting Cells
    eformat= workbook.add_format({'text_wrap': True, 'valign': 'top', 'align': 'left'})
    eformatlabel= workbook.add_format({'text_wrap': True, 'valign': 'top', 'bold': True})
    dateformat = workbook.add_format({'num_format': 'mm/dd/yyyy', 'align': 'left'})


    # Setting the column widths
    worksheet.set_column(0,0,13.86)
    worksheet.set_column(1,1,13.71)
    worksheet.set_column(2,2,13.86)
    worksheet.set_column(3,3,14.14)
    worksheet.set_column(4,4,17.14)
    worksheet.set_column(5,5,14.57)

    # Inserting a header
    worksheet.set_header('Needham Unreceived Orders')

    # Adding column labels
    worksheet.write(0,0,'Item Created', eformatlabel)
    worksheet.write(0,1,'Bib Number', eformatlabel)
    worksheet.write(0,2,'Item Number', eformatlabel)
    worksheet.write(0,3,'Order Number', eformatlabel)
    worksheet.write(0,4,'Order Status Code', eformatlabel)
    worksheet.write(0,5,'Received Date', eformatlabel)


    # Writing the report for staff to the Excel worksheet
    for rownum, row in enumerate(query_results):
        worksheet.write(rownum+1,0,row[0], dateformat)
        worksheet.write(rownum+1,1,row[1], eformat)
        worksheet.write(rownum+1,2,row[2], eformat)
        worksheet.write(rownum+1,3,row[3], eformat)
        worksheet.write(rownum+1,4,row[4], eformat)
        worksheet.write(rownum+1,5,row[5], dateformat)
    
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
              rmi.creation_date_gmt::DATE AS item_created,
              rmb.record_type_code||rmb.record_num||'a' AS bib_number,
              rmi.record_type_code||rmi.record_num||'a' AS item_number,
              rmo.record_type_code||rmo.record_num||'a' AS order_number,
              o.order_status_code,
              o.received_date_gmt::DATE AS received_date

            FROM sierra_view.bib_record_item_record_link l
            JOIN sierra_view.item_record i
              ON l.item_record_id = i.id
              AND i.location_code ~ '^nee'
            JOIN sierra_view.record_metadata rmi
              ON i.id = rmi.id
              AND rmi.creation_date_gmt::DATE >= '2023-11-05'
            JOIN sierra_view.bib_record_order_record_link lo
              ON l.bib_record_id = lo.bib_record_id
            JOIN sierra_view.order_record o
              ON lo.order_record_id = o.id
              AND o.accounting_unit_code_num = '28'
              AND o.order_status_code = 'o'
            JOIN sierra_view.record_metadata rmo
              ON o.id = rmo.id
            JOIN sierra_view.record_metadata rmb
              ON lo.bib_record_id = rmb.id

            ORDER BY 1,2
        """

    query_results = run_query(query)
    excel_file = "/Scripts/Needham Unreceived Orders/Temp Files/Needham Unreceived Orders {}.xlsx".format(date.today())
    excel_writer(query_results, excel_file)

    # generate email message
    email_subject = "Needham unreceived Orders"
    email_message = """***This is an automated email**

The Needham Unreceived Orders report has been attached."""

    # read config file with recipient list for email
    config_recipient = configparser.ConfigParser()
    config_recipient.read("C:\\Scripts\\Creds\\emails.ini")
    email_recipient = config_recipient["needham_unreceived_orders"]["recipients"].split()
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
        emailto = config_recipient["script_error_extended"]["recipients"].split()

        # craft email subject and message containing error message details from traceback
        email_subject = "Needham Unreceived Orders script error"
        email_message = (
            "Your script failed with the following error:\n\n" + traceback.format_exc()
        )

        send_email_error(email_subject, email_message, emailto)
        raise