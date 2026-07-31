#!/usr/bin/env python3

#Run in py313
"""
Jeremy Goldstein
Minuteman Library Network

Generates monthly Report on the cost per circ
in the past month broken out by item type

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
    workbook = xlsxwriter.Workbook(excel_file)
    worksheet = workbook.add_worksheet()


    # Formatting our Excel worksheet
    worksheet.set_landscape()
    worksheet.hide_gridlines(0)

    # Formatting Cells
    eformat= workbook.add_format({'text_wrap': True, 'valign': 'top'})
    eformatlabel= workbook.add_format({'text_wrap': True, 'valign': 'top', 'bold': True})
    eformat2= workbook.add_format({'num_format': 'mm/dd/yy hh:mm:ss'})


    # Setting the column widths
    worksheet.set_column(0,0,16.86)
    worksheet.set_column(1,1,11.86)
    worksheet.set_column(2,2,11.14)
    worksheet.set_column(3,3,14.14)

    # Inserting a header
    worksheet.set_header('Wayland Monthly Circ Value')

    # Adding column labels
    worksheet.write(0,0,'IType', eformatlabel)
    worksheet.write(0,1,'Value', eformatlabel)
    worksheet.write(0,2,'Circ_Count', eformatlabel)
    worksheet.write(0,3,'Value_Per_Circ', eformatlabel)


    # Writing the report for staff to the Excel worksheet
    for rownum, row in enumerate(query_results):
        worksheet.write(rownum+1,0,row[0], eformat)
        worksheet.write(rownum+1,1,row[1], eformat)
        worksheet.write(rownum+1,2,row[2], eformat)
        worksheet.write(rownum+1,3,row[3], eformat)
    
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
              *
            FROM (
              SELECT
                it.name AS itype,
                CAST(SUM(i.price) AS MONEY) AS value,
                COUNT(DISTINCT c.id) AS circ_count,
                (CAST(SUM(i.price) AS MONEY) / COUNT(DISTINCT c.id)) as value_per_circ

              FROM sierra_view.circ_trans c
              JOIN sierra_view.item_record i
                ON c.item_record_id = i.id
              JOIN sierra_view.itype_property_myuser it
                ON i.itype_code_num = it.code

              WHERE c.op_code IN ('o','r')
                AND c.transaction_gmt::DATE >= (current_date - INTERVAL '1 month')
                AND c.stat_group_code_num BETWEEN '740' AND '749'

              GROUP BY 1

              UNION
              
              SELECT
                'Total' AS itype,
                CAST(SUM(i.price) AS MONEY) AS value,
                COUNT(DISTINCT c.id) AS circ_count,
                (CAST(SUM(i.price) AS MONEY) / COUNT(DISTINCT c.id)) as value_per_circ

              FROM sierra_view.circ_trans c
              JOIN sierra_view.item_record i
                ON c.item_record_id = i.id
              JOIN sierra_view.itype_property_myuser it
                ON i.itype_code_num = it.code

              WHERE c.op_code IN ('o','r')
                AND c.transaction_gmt::DATE >= (current_date - INTERVAL '1 month')
                AND c.stat_group_code_num BETWEEN '740' AND '749'
            )inner_query

            ORDER BY CASE
	          WHEN itype = 'Total' THEN 2
	          ELSE 1
            END,itype
            """

    query_results = run_query(query)

    # generate excel file from those query results
    excel_file = "/Scripts/Wayland Monthly Reports/Temp Files/WYLMonthlyCircValue{}.xlsx".format(date.today())
    excel_file = excel_writer(query_results, excel_file)

    # send email
    email_subject = 'Wayland Monthly Circ Value'
    email_message = '''***This is an automated email***


The Wayland Monthly Circ Value report has been attached.'''
	# read config file with recipient list for email
    config_recipient = configparser.ConfigParser()
    config_recipient.read("C:\\Scripts\\Creds\\emails.ini")
    email_recipient = config_recipient["wayland_monthly_circ_value"]["recipients"].split()  
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
        email_subject = "Wayland Monthly Circ Value Script error"
        email_message = (
            "Your script failed with the following error:\n\n" + traceback.format_exc()
        )

        send_email_error(email_subject, email_message, emailto)
        raise
