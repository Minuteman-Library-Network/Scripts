#!/usr/bin/env python3

#Run in py313
"""
Jeremy Goldstein
Minuteman Library Network

Generates monthly report of checkins at Medfield
with a breakdown by itype and days early/overdue
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

def excel_writer(query_results, excel_file):
    #Creating the Excel file for staff
    workbook = xlsxwriter.Workbook(excel_file, {'remove_timezone': True})
    worksheet = workbook.add_worksheet('monthly checkins')

    #Formatting our Excel worksheet
    worksheet.set_landscape()
    worksheet.hide_gridlines(0)

    #Formatting Cells
    eformat= workbook.add_format({'text_wrap': True, 'valign': 'top'})
    eformatlabel= workbook.add_format({'text_wrap': True, 'valign': 'top', 'bold': True})

    # Setting the column widths
    worksheet.set_column(0,0,15.86)
    worksheet.set_column(1,1,13.14)
    worksheet.set_column(2,2,32.86)
    worksheet.set_column(3,3,20)
    worksheet.set_column(4,4,16.86)
    worksheet.set_column(5,5,19)
    worksheet.set_column(6,6,25)
    worksheet.set_column(7,7,26)
    worksheet.set_column(8,8,27)
    worksheet.set_column(9,9,33.86)

    #Inserting a header
    worksheet.set_header('monthly checkins report')

    # Adding column labels
    worksheet.write(0,0,'itype', eformatlabel)
    worksheet.write(0,1,'total_checkins', eformatlabel)
    worksheet.write(0,2,'returned_greater_than_1_day_early', eformatlabel)
    worksheet.write(0,3,'returned_1_day_early', eformatlabel)
    worksheet.write(0,4,'returned_on_time', eformatlabel)
    worksheet.write(0,5,'returned_1_day_late', eformatlabel)
    worksheet.write(0,6,'returned_2_to_7_days_late', eformatlabel)
    worksheet.write(0,7,'returned_8_to_14_days_late', eformatlabel)
    worksheet.write(0,8,'returned_15_to_21_days_late', eformatlabel)
    worksheet.write(0,9,'returned_greater_than_21_days_late', eformatlabel)

    # Writing the report for staff to the Excel worksheet

    for rownum, row in enumerate(query_results):
        worksheet.write(rownum+1,0,row[0], eformat)
        worksheet.write(rownum+1,1,row[1], eformat)
        worksheet.write(rownum+1,2,row[2], eformat)
        worksheet.write(rownum+1,3,row[3], eformat)
        worksheet.write(rownum+1,4,row[4], eformat)
        worksheet.write(rownum+1,5,row[5], eformat)
        worksheet.write(rownum+1,6,row[6], eformat)
        worksheet.write(rownum+1,7,row[7], eformat)
        worksheet.write(rownum+1,8,row[8], eformat)
        worksheet.write(rownum+1,9,row[9], eformat)
    
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
              COUNT(t.id) AS total_checkins,
              COUNT(t.id) FILTER (WHERE t.due_date_gmt::DATE - t.transaction_gmt::DATE > 1) AS returned_greater_than_1_day_early,
              COUNT(t.id) FILTER (WHERE t.due_date_gmt::DATE - t.transaction_gmt::DATE = 1) AS returned_1_day_early,
              COUNT(t.id) FILTER (WHERE t.due_date_gmt::DATE = t.transaction_gmt::DATE) AS returned_on_time,
              COUNT(t.id) FILTER (WHERE t.transaction_gmt::DATE - t.due_date_gmt::DATE = 1) AS returned_1_day_late,
              COUNT(t.id) FILTER (WHERE t.transaction_gmt::DATE - t.due_date_gmt::DATE BETWEEN 2 AND 7) AS returned_2_to_7_days_late,
              COUNT(t.id) FILTER (WHERE t.transaction_gmt::DATE - t.due_date_gmt::DATE BETWEEN 8 AND 14) AS returned_8_to_14_days_late,
              COUNT(t.id) FILTER (WHERE t.transaction_gmt::DATE - t.due_date_gmt::DATE BETWEEN 15 AND 21) AS returned_15_to_21_days_late,
              COUNT(t.id) FILTER (WHERE t.transaction_gmt::DATE - t.due_date_gmt::DATE > 21) AS returned_greater_than_21_days_late
            
            FROM sierra_view.circ_trans t
            JOIN sierra_view.itype_property_myuser it
              ON t.itype_code_num = it.code

            WHERE t.op_code = 'i'
              AND t.transaction_gmt::DATE BETWEEN (CURRENT_DATE - INTERVAL '1 month') AND (CURRENT_DATE - INTERVAL '1 day')
              AND t.stat_group_code_num::varchar ~ '^50[0-9]'
            GROUP BY 1
            
            UNION
            
            SELECT
              'TOTAL' AS itype,
              COUNT(t.id) AS total_checkins,
              COUNT(t.id) FILTER (WHERE t.due_date_gmt::DATE - t.transaction_gmt::DATE > 1) AS returned_greater_than_1_day_early,
              COUNT(t.id) FILTER (WHERE t.due_date_gmt::DATE - t.transaction_gmt::DATE = 1) AS returned_1_day_early,
              COUNT(t.id) FILTER (WHERE t.due_date_gmt::DATE = t.transaction_gmt::DATE) AS returned_on_time,
              COUNT(t.id) FILTER (WHERE t.transaction_gmt::DATE - t.due_date_gmt::DATE = 1) AS returned_1_day_late,
              COUNT(t.id) FILTER (WHERE t.transaction_gmt::DATE - t.due_date_gmt::DATE BETWEEN 2 AND 7) AS returned_2_to_7_days_late,
              COUNT(t.id) FILTER (WHERE t.transaction_gmt::DATE - t.due_date_gmt::DATE BETWEEN 8 AND 14) AS returned_8_to_14_days_late,
              COUNT(t.id) FILTER (WHERE t.transaction_gmt::DATE - t.due_date_gmt::DATE BETWEEN 15 AND 21) AS returned_15_to_21_days_late,
              COUNT(t.id) FILTER (WHERE t.transaction_gmt::DATE - t.due_date_gmt::DATE > 21) AS returned_greater_than_21_days_late
            
            FROM sierra_view.circ_trans t
            JOIN sierra_view.itype_property_myuser it
              ON t.itype_code_num = it.code

            WHERE t.op_code = 'i'
              AND t.transaction_gmt::DATE BETWEEN (CURRENT_DATE - INTERVAL '1 month') AND (CURRENT_DATE - INTERVAL '1 day')
              AND t.stat_group_code_num::varchar ~ '^50[0-9]'
            GROUP BY 1
          )a
            
          ORDER BY CASE
            WHEN itype = 'TOTAL' THEN 2
            ELSE 1
          END, itype
          """

    query_results = run_query(query)

    # generate excel file from those query results
    excel_file = "/Scripts/Medfield Monthly Check Ins/Temp Files/mld_monthly_checkins{}.xlsx".format(date.today())
    excel_file = excel_writer(query_results, excel_file)

    email_subject = 'Monthly Checkin Report'
    email_message = '''***This is an automated email***


    The monthly checkin report has been attached.'''
	# read config file with recipient list for email
    config_recipient = configparser.ConfigParser()
    config_recipient.read("C:\\Scripts\\Creds\\emails.ini")
    email_recipient = config_recipient["medfield_monthly_checkins"]["recipients"].split()  
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
        email_subject = "Medfield Monthly Check Ins Script Error"
        email_message = (
            "Your script failed with the following error:\n\n" + traceback.format_exc()
        )

        send_email_error(email_subject, email_message, emailto)
        raise
