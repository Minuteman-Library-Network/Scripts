#!/usr/bin/env python3

#Run in py313
"""
Jeremy Goldstein
Minuteman Library Network

Generates weekly new patrons report for Watertown
Report is then emailed to staff an Excel attachment
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

def excel_writer(query_results, excel_file):
    # Creating the Excel file for staff
    workbook = xlsxwriter.Workbook(excel_file)
    worksheet = workbook.add_worksheet()

    # Formatting our Excel worksheet
    worksheet.set_landscape()
    worksheet.hide_gridlines(0)

    # Formatting Cells
    eformat= workbook.add_format({'text_wrap': True, 'valign': 'top'})
    eformatlabel= workbook.add_format({'text_wrap': True, 'valign': 'top', 'bold': True})
    eformat2= workbook.add_format({'num_format': 'mm/dd/yy'})

    # Setting the column widths
    worksheet.set_column(0,0,30.57)
    worksheet.set_column(1,1,13.3)
    worksheet.set_column(2,2,12)
    worksheet.set_column(3,3,13.8)
    worksheet.set_column(4,4,24.14)
    worksheet.set_column(5,5,14)
    worksheet.set_column(6,6,14)

    # Inserting a header
    worksheet.set_header('Watertown New Patrons')

    # Adding column labels
    worksheet.write(0,0,'Email', eformatlabel)
    worksheet.write(0,1,'Last Name', eformatlabel)
    worksheet.write(0,2,'First Name', eformatlabel)
    worksheet.write(0,3,'Middle Name', eformatlabel)
    worksheet.write(0,4,'Patron Agency', eformatlabel)
    worksheet.write(0,5,'City', eformatlabel)
    worksheet.write(0,6,'Creation Date', eformatlabel)

    # Writing the report for staff to the Excel worksheet
    for rownum, row in enumerate(query_results):
        worksheet.write(rownum+1,0,row[0], eformat)
        worksheet.write(rownum+1,1,row[1], eformat)
        worksheet.write(rownum+1,2,row[2], eformat)
        worksheet.write(rownum+1,3,row[3], eformat)
        worksheet.write(rownum+1,4,row[4], eformat)
        worksheet.write(rownum+1,5,row[5], eformat)
        worksheet.write(rownum+1,6,row[6], eformat2)
    
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
              COALESCE(e.content,'') AS email,
              pn.last_name,
              pn.first_name,
              pn.middle_name,
              ap.name AS agency,
              INITCAP(COALESCE(REGEXP_REPLACE(REGEXP_REPLACE(a.city,'\d','','g'),'\sma$','','i'),'')) AS city,
              rm.creation_date_gmt::DATE AS creation_date

            FROM sierra_view.patron_record p
            JOIN sierra_view.agency_property_myuser ap
              ON p.patron_agency_code_num = ap.code
            JOIN sierra_view.record_metadata rm
              ON p.id = rm.id
            JOIN sierra_view.patron_record_fullname pn
              ON p.id = pn.patron_record_id
              AND pn.display_order = '0'
            JOIN sierra_view.subfield e
              ON p.id = e.record_id
              AND e.field_type_code = 'z'
              AND e.occ_num = 0
            JOIN sierra_view.patron_record_address a
              ON p.id = a.patron_record_id

            WHERE p.ptype_code = '35'
              AND ap.code IN ('10','47')
            AND rm.creation_date_gmt::DATE BETWEEN (NOW()::DATE - INTERVAL '1 week') AND (NOW()::DATE - INTERVAL '1 day')
            """

    query_results = run_query(query)

    # generate excel file from those query results
    excel_file = "/Scripts/Watertown New Patrons/Temp Files/WATNewPatrons{}.xlsx".format(date.today())
    excel_file = excel_writer(query_results, excel_file)

    email_subject = 'Watertown New Patrons'
    email_message = '''***This is an automated email***


    The Watertown New Patrons report has been attached.'''
	# read config file with recipient list for email
    config_recipient = configparser.ConfigParser()
    config_recipient.read("C:\\Scripts\\Creds\\emails.ini")
    email_recipient = config_recipient["watertown_new_patrons"]["recipients"].split()  
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
        email_subject = "Watertown New Patrons Script Error"
        email_message = (
            "Your script failed with the following error:\n\n" + traceback.format_exc()
        )

        send_email_error(email_subject, email_message, emailto)
        raise

