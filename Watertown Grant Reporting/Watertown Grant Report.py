#!/usr/bin/env python3

"""
Jeremy Goldstein
Minuteman Library Network

Generates monthly collection usage report required
for the reporting of a grant in Watertown
"""
# run in py313

import configparser
import xlsxwriter
import psycopg2
import pysftp
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email.utils import formatdate
from email import encoders
import traceback
from datetime import date
import os

# from oauth2client.service_account import ServiceAccountCredentials
# from googleapiclient.discovery import build


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
    # return variable containing query results
    return rows


# log items that were corrected to an existing Google Sheet
def excel_writer(query_results, excel_file):
    workbook = xlsxwriter.Workbook(excel_file, {"remove_timezone": True})
    worksheet = workbook.add_worksheet()

    # Formatting our Excel worksheet
    worksheet.set_landscape()
    worksheet.hide_gridlines(0)

    # Formatting Cells
    eformat = workbook.add_format({"text_wrap": True, "valign": "top"})
    eformatlabel = workbook.add_format(
        {"text_wrap": True, "valign": "top", "bold": True}
    )
    dateformat = workbook.add_format(
        {"text_wrap": True, "num_format": "yyyy/mm/dd", "valign": "top"}
    )

    # Setting the column widths
    worksheet.set_column(0, 0, 65.22)
    worksheet.set_column(1, 1, 14.89)

    # Inserting a header
    worksheet.set_header("Monthly Grant Usage Report")

    # Adding column labels
    worksheet.write(0, 0, "User", eformatlabel)
    worksheet.write(0, 1, "Date", eformatlabel)

    # Writing the report for staff to the Excel worksheet
    for rownum, row in enumerate(query_results):
        worksheet.write(rownum + 1, 0, row[0], eformat)
        worksheet.write(rownum + 1, 1, row[1], dateformat)

    workbook.close()


# upload report to SIC directory and optionally remove older files
def sftp_file(local_file):

    config = configparser.ConfigParser()
    config.read("C:\\Scripts\\Creds\\config.ini")

    cnopts = pysftp.CnOpts()

    srv = pysftp.Connection(
        host=config["sic"]["sic_host"],
        username=config["sic"]["sic_user"],
        password=config["sic"]["sic_pw"],
        cnopts=cnopts,
    )

    local_file = local_file

    # upload file
    srv.cwd("/reports/Library-Specific Reports/Watertown/Custom/")
    srv.put(local_file)
    # close sftp connection
    srv.close()
    # remove local copy of file
    os.remove(local_file)


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
    """
    query to gather encrypted list of patrons who checked out items in a particular collection
    used to generate a unique user count, mandated as a reporting requirement for a grant
    """
    query = """
    SELECT 
      DISTINCT ENCODE(SHA256(t.patron_record_id::VARCHAR::BYTEA||rmp.record_num::VARCHAR::BYTEA||TO_CHAR(rmp.creation_date_gmt,'dayMonDD YYYY SS')::BYTEA),'hex') AS user,
      t.transaction_gmt::DATE AS checkout_date
 
    FROM sierra_view.circ_trans t
    JOIN sierra_view.item_record_property i
      ON t.item_record_id = i.item_record_id
    JOIN sierra_view.record_metadata rmp
      ON t.patron_record_id = rmp.id
  
    WHERE t.op_code = 'o'
      AND t.transaction_gmt::DATE < CURRENT_DATE
      AND t.transaction_gmt::DATE >= (CURRENT_DATE - INTERVAL '1 month')
      AND i.barcode IN 
        (
        '34868007565914',
        '34868007565922',
        '34868007565930',
        '34868007565948',
        '34868007565955',
        '34868007565963',
        '34868007565971',
        '34868007565989',
        '34868007565997',
        '34868007565815',
        '34868007565823',
        '34868007565831',
        '34868007565849',
        '34868007565856',
        '34868007565864',
        '34868007565872',
        '34868007565880',
        '34868007565898',
        '34868007565906',
        '34868007565708',
        '34868007565716',
        '34868007565724',
        '34868007565732',
        '34868007565740',
        '34868007565757',
        '34868007565765',
        '34868007565773',
        '34868007565781',
        '34868007565799',
        '34868007566003'
        )
    """

    # run query and write results to an excel file
    results = run_query(query)
    # Name of Excel File
    excel_file = (
        "/Scripts/Watertown Grant Reporting/Temp Files/WATGrantReport{}.xlsx".format(
            date.today()
        )
    )
    excel_writer(results, excel_file)

    # sftp file to intranet site for distribution and delete local copy upon transfer
    sftp_file(
        "C:\\Scripts\\Watertown Grant Reporting\\Temp Files\\WATGrantReport{}.xlsx".format(
            date.today()
        )
    )


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
        email_subject = "Watertown grant report script error"
        email_message = (
            "Your script failed with the following error:\n\n" + traceback.format_exc()
        )

        send_email_error(email_subject, email_message, emailto)
        raise
