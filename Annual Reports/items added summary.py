#!/usr/bin/env python3

"""
Jeremy Goldstein
Minuteman Library Network

Generates report of new items per library for use with annual reports
"""
# run in py38

import psycopg2
import configparser
import xlsxwriter
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

    # Creating the Excel file for staff
    workbook = xlsxwriter.Workbook(excel_file)
    worksheet = workbook.add_worksheet()

    # Formatting our Excel worksheet
    worksheet.set_landscape()
    worksheet.hide_gridlines(0)

    # Formatting Cells
    eformat = workbook.add_format({"text_wrap": True, "valign": "top"})
    eformatlabel = workbook.add_format(
        {"text_wrap": True, "valign": "top", "bold": True}
    )

    # Setting the column widths
    worksheet.set_column(0, 0, 26.43)
    worksheet.set_column(1, 1, 22.57)
    worksheet.set_column(2, 2, 16.00)

    # Inserting a header
    worksheet.set_header("Items Added Summary")

    # Adding column labels
    worksheet.write(0, 0, "LOCATION", eformatlabel)
    worksheet.write(0, 1, "ITEM RECORDS ADDED", eformatlabel)
    worksheet.write(0, 2, "ADJUSTED TOTAL", eformatlabel)

    # Writing the report for staff to the Excel worksheet
    for rownum, row in enumerate(query_results):
        worksheet.write(rownum + 1, 0, row[0], eformat)
        worksheet.write(rownum + 1, 1, row[1], eformat)
        worksheet.write(rownum + 1, 2, row[2], eformat)

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
    emailto = config_recipient["annual_reports"]["recipients"].split()
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
    # query to count new items added at each location
    query = """\
            SELECT
              --account for instance where the library name is not the municipality name
              CASE
                WHEN SUBSTRING(i.location_code, 1, 2) = 'ac' THEN 'ACTON'
                WHEN SUBSTRING(i.location_code, 1, 2) = 'ar' THEN 'ARLINGTON'
                WHEN SUBSTRING(i.location_code, 1, 2) = 'br' THEN 'BROOKLINE'
                WHEN SUBSTRING(i.location_code, 1, 2) = 'ca' THEN 'CAMBRIDGE'
                WHEN SUBSTRING(i.location_code, 1, 2) = 'co' THEN 'CONCORD'
                WHEN SUBSTRING(i.location_code, 1, 2) = 'dd' THEN 'DEDHAM'
                WHEN SUBSTRING(i.location_code, 1, 2) = 'fp' THEN 'FRAMINGHAM'
                WHEN SUBSTRING(i.location_code, 1, 2) = 'na' THEN 'NATICK'
                WHEN SUBSTRING(i.location_code, 1, 2) = 'so' THEN 'SOMERVILLE'
                WHEN SUBSTRING(i.location_code, 1, 2) = 'wl' THEN 'WALTHAM'
                WHEN SUBSTRING(i.location_code, 1, 2) = 'wa' THEN 'WATERTOWN'
                WHEN SUBSTRING(i.location_code, 1, 2) = 'we' THEN 'WELLESLEY'
                WHEN SUBSTRING(i.location_code, 1, 2) = 'ww' THEN 'WESTWOOD'
                ELSE l.NAME 
              END AS Library,
              COUNT(i.id) AS "ITEM RECORDS ADDED",
              COUNT(i.id) FILTER(WHERE i.itype_code_num NOT IN ('10','107', '158', '239', '240', '241', '242', '244', '248', '249')) AS "ADJUSTED TOTAL"

            FROM sierra_view.item_record i
            JOIN sierra_view.location_myuser l
              ON SUBSTRING(i.location_code, 1, 3) = l.code 
            JOIN sierra_view.record_metadata m
              ON i.id = m.id
              
            WHERE i.itype_code_num != '80'
              AND m.creation_date_gmt >= NOW() - INTERVAL '1 year'
            GROUP BY 1
            ORDER BY 1
            """
    query_results = run_query(query)

    # generate excel file from those query results
    excel_file = "/Scripts/Annual Reports/Archive/items added summary{}.xlsx".format(
        date.today()
    )
    excel_writer(query_results, excel_file)

    # send email with attached file
    email_subject = "Items Added Summary"
    email_message = """***Items Added Summary***


The Items Added Summary report has been attached."""
    send_email(email_subject, email_message, excel_file)


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
        email_subject = "annual reports: items added summary script error"
        email_message = (
            "Your script failed with the following error:\n\n" + traceback.format_exc()
        )

        send_email_error(email_subject, email_message, emailto)
        raise
