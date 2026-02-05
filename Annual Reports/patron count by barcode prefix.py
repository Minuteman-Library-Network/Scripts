#!/usr/bin/env python3

"""
Jeremy Goldstein
Minuteman Library Network

Generates report of current patrons by barcode prefix
Used to identify the library that issued a card
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
    worksheet.set_column(0, 0, 16.43)
    worksheet.set_column(1, 1, 16.43)
    worksheet.set_column(2, 2, 16.43)
    worksheet.set_column(3, 3, 16.43)

    # Inserting a header
    worksheet.set_header("Patron Count By Barcode Prefix")

    # Adding column labels
    worksheet.write(0, 0, "Library", eformatlabel)
    worksheet.write(0, 1, "Total_Count", eformatlabel)
    worksheet.write(0, 2, "Active_Count", eformatlabel)
    worksheet.write(0, 3, "New_Count", eformatlabel)

    # Writing the report for staff to the Excel worksheet
    for rownum, row in enumerate(query_results):
        worksheet.write(rownum + 1, 0, row[0], eformat)
        worksheet.write(rownum + 1, 1, row[1], eformat)
        worksheet.write(rownum + 1, 2, row[2], eformat)
        worksheet.write(rownum + 1, 3, row[3], eformat)

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


def main():
    # query to identify patron records with incorrect owed_amt fields
    query = """\
           SELECT
             CASE
               WHEN b.index_entry LIKE '22211%' THEN 'Acton'
               WHEN b.index_entry LIKE '24860%' THEN 'Arlington'
               WHEN b.index_entry LIKE '20308%' THEN 'Ashland'
               WHEN b.index_entry LIKE '24861%' THEN 'Bedford'
               WHEN b.index_entry LIKE '24862%' THEN 'Belmont'
               WHEN b.index_entry LIKE '21712%' THEN 'Brookline'
               WHEN b.index_entry LIKE '21189%' THEN 'Cambridge'
               WHEN b.index_entry LIKE '24863%' THEN 'Concord'
               WHEN b.index_entry LIKE '20423%' THEN 'Dean'
               WHEN b.index_entry LIKE '26504%' THEN 'Dedham'
               WHEN b.index_entry LIKE '26304%' THEN 'Dover'
               WHEN b.index_entry LIKE '21213%' THEN 'Framingham Public'
               WHEN b.index_entry LIKE '23014%' THEN 'Framingham State'
               WHEN b.index_entry LIKE '26998%' THEN 'Franklin'
               WHEN b.index_entry LIKE '26287%' THEN 'Holliston'
               WHEN b.index_entry LIKE '23015%' THEN 'Lasell'
               WHEN b.index_entry LIKE '21619%' THEN 'Lexington'
               WHEN b.index_entry LIKE '24864%' THEN 'Lincoln'
               WHEN b.index_entry LIKE '26294%' THEN 'Mass Bay'
               WHEN b.index_entry LIKE '25957%' THEN 'Maynard'
               WHEN b.index_entry LIKE '21848%' THEN 'Medfield'
               WHEN b.index_entry LIKE '24865%' THEN 'Medford'
               WHEN b.index_entry LIKE '21852%' THEN 'Medway'
               WHEN b.index_entry LIKE '26216%' THEN 'Millis'
               WHEN b.index_entry LIKE '20022%' THEN 'Mount Ida'
               WHEN b.index_entry LIKE '23016%' THEN 'Natick'
               WHEN b.index_entry LIKE '23017%' THEN 'Needham'
               WHEN b.index_entry LIKE '21323%' THEN 'Newton'
               WHEN b.index_entry LIKE '22405%' THEN 'Norwood'
               WHEN b.index_entry LIKE '22101%' THEN 'Olin'
               WHEN b.index_entry LIKE '21911%' THEN 'Pine Manor'
               WHEN b.index_entry LIKE '21927%' THEN 'Regis'
               WHEN b.index_entry LIKE '28106%' THEN 'Sherborn'
               WHEN b.index_entry LIKE '21155%' THEN 'Somerville'
               WHEN b.index_entry LIKE '22051%' THEN 'Stow'
               WHEN b.index_entry LIKE '24866%' THEN 'Sudbury'
               WHEN b.index_entry LIKE '24867%' THEN 'Waltham'
               WHEN b.index_entry LIKE '24868%' THEN 'Watertown'
               WHEN b.index_entry LIKE '24869%' THEN 'Wayland'
               WHEN b.index_entry LIKE '24870%' THEN 'Wellesley'
               WHEN b.index_entry LIKE '24871%' THEN 'Weston'
               WHEN b.index_entry LIKE '23018%' THEN 'Westwood'
               WHEN b.index_entry LIKE '24872%' THEN 'Winchester'
               WHEN b.index_entry LIKE '21906%' THEN 'Woburn'
               ELSE 'Unknown'
             END AS Library,
             COUNT(p.id) AS Total_Patrons,
             COUNT(p.id) FILTER(WHERE p.activity_gmt > (localtimestamp - INTERVAL '1 year')) AS Active_Patrons,
             COUNT(p.id) FILTER(WHERE m.creation_date_gmt > (localtimestamp - INTERVAL '1 year')) AS New_Patrons
           FROM sierra_view.patron_record p
           JOIN sierra_view.record_metadata m
             ON p.id = m.id
           JOIN sierra_view.phrase_entry b
             ON p.id = b.record_id AND b.varfield_type_code = 'b'
           GROUP BY 1
           ORDER BY 1
           """
    query_results = run_query(query)

    # generate excel file from those query results
    excel_file = "/Scripts/Annual Reports/Archive/PatronCountByBarcode{}.xlsx".format(
        date.today()
    )
    excel_writer(query_results, excel_file)

    # send email with attached file
    email_subject = "Patron Count By Barcode"
    email_message = """***This is an automated email***


The Patron Count by Barcode report has been attached."""
    send_email(email_subject, email_message, excel_file)


main()
