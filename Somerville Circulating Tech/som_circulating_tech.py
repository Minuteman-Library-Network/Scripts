#!/usr/bin/env python3

#Run in py313
"""
Jeremy Goldstein
Minuteman Library Network

Generates weekly report on the present status of all of Somerville's circulating tech items
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

    #Creating the Excel file for staff
    workbook = xlsxwriter.Workbook(excel_file)
    worksheet = workbook.add_worksheet()


    #Formatting our Excel worksheet
    worksheet.set_landscape()
    worksheet.hide_gridlines(0)

    #Formatting Cells
    eformat= workbook.add_format({'text_wrap': True, 'valign': 'top', 'align': 'left'})
    eformatlabel= workbook.add_format({'text_wrap': True, 'valign': 'top', 'bold': True})
    eformat2= workbook.add_format({'num_format': 'mm/dd/yy'})
    dateformat = workbook.add_format({'num_format': 'mm/dd/yyyy', 'align': 'left'})


    # Setting the column widths
    worksheet.set_column(0,0,12.1)
    worksheet.set_column(1,1,28.3)
    worksheet.set_column(2,2,13.5)
    worksheet.set_column(3,3,13.9)
    worksheet.set_column(4,4,40.75)
    worksheet.set_column(5,5,17.5)
    worksheet.set_column(6,6,16.75)
    worksheet.set_column(7,7,17.8)
    worksheet.set_column(8,8,12.75)
    worksheet.set_column(9,9,16.3)
    worksheet.set_column(10,10,13.9)

    #Inserting a header
    worksheet.set_header('Somerville Circulating Tech')

    # Adding column labels
    worksheet.write(0,0,'Bib Number', eformatlabel)
    worksheet.write(0,1,'Bib Title', eformatlabel)
    worksheet.write(0,2,'Item Number', eformatlabel)
    worksheet.write(0,3,'Item Location', eformatlabel)
    worksheet.write(0,4,'Call Number', eformatlabel)
    worksheet.write(0,5,'Barcode', eformatlabel)
    worksheet.write(0,6,'Due Date', eformatlabel)
    worksheet.write(0,7,'Checked Out Date', eformatlabel)
    worksheet.write(0,8,'# Notices Sent', eformatlabel)
    worksheet.write(0,9,'Last Notice Sent', eformatlabel)
    worksheet.write(0,10,'Item Status', eformatlabel)

    # Writing the report for staff to the Excel worksheet
    for rownum, row in enumerate(query_results):
        worksheet.write(rownum+1,0,row[0], eformat)
        worksheet.write(rownum+1,1,row[1], eformat)
        worksheet.write(rownum+1,2,row[2], eformat)
        worksheet.write(rownum+1,3,row[3], eformat)
        worksheet.write(rownum+1,4,row[4], eformat)
        worksheet.write(rownum+1,5,row[5], eformat)
        worksheet.write(rownum+1,6,row[6], dateformat)
        worksheet.write(rownum+1,7,row[7], dateformat)
        worksheet.write(rownum+1,8,row[8], eformat)
        worksheet.write(rownum+1,9,row[9], dateformat)
        worksheet.write(rownum+1,10,row[10], eformat)
    
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
    emailto = config_recipient["somerville_circulating_tech"]["recipients"].split()
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
    # query to identify bills from the past week for childrens items belonging to Cambridge
    query = """
    SELECT 
      rmb.record_type_code||rmb.record_num||'a' AS "Bib Number",
      bp.best_title AS "Bib Title",
      rmi.record_type_code||rmi.record_num|| 'a'  AS "Item Number",
      i.location_code AS "Item Location",
      TRIM(REGEXP_REPLACE(ip.call_number,'\|.',' ','g')) AS "Call Number",
      ip.barcode AS "Barcode",
      o.due_gmt::DATE AS "Due Date",
      o.checkout_gmt::DATE AS "Checked Out Date",
      o.overdue_count AS "Notices Sent",
      o.overdue_gmt::DATE AS "Last Notice Sent",
      CASE
        WHEN o.id IS NOT NULL THEN 'CHECKED OUT'
        ELSE s.name
      END AS "Item Status"
      
    FROM sierra_view.record_metadata rmb
    JOIN sierra_view.bib_record_property bp
      ON rmb.id = bp.bib_record_id
    JOIN sierra_view.bib_record_item_record_link bri
      ON rmb.id = bri.bib_record_id
    JOIN sierra_view.record_metadata rmi
      ON bri.item_record_id = rmi.id
    JOIN sierra_view.item_record i
      ON rmi.id = i.id
    JOIN sierra_view.item_record_property ip
      ON i.id = ip.item_record_id
    JOIN sierra_view.item_status_property_myuser s
      ON i.item_status_code = s.code
    LEFT JOIN sierra_view.checkout o
      ON i.id = o.item_record_id
    
    WHERE rmb.record_type_code = 'b' 
      AND rmb.record_type_code||rmb.record_num IN (
      'b4025951',
      'b4025971',
      'b3885108',
      'b3628392',
      'b4041179',
      'b4025951',
      'b4025951',
      'b4034839',
      'b4164776',
      'b4175105',
      'b4223928',
      'b4226395',
      'b4226396',
      'b4041179',
      'b4224769',
      'b4224771',
      'b3628392',
      'b4224776',
      'b4224777',
      'b3885108',
      'b4224780',
      'b4224781',
      'b4164776',
      'b4224782',
      'b4224783',
      'b4025971',
      'b4224784',
      'b4224787',
      'b4025951',
      'b4224789',
      'b4224791',
      'b4175105',
      'b4396824',
      'b4396829',
      'b4396830'
      )
    ORDER BY 1
    """
    
    query_results = run_query(query)

    # generate excel file from those query results
    excel_file = "/Scripts/Somerville Circulating Tech/Temp Files/Somerville Circulating Tech {}.xlsx".format(date.today())
    excel_writer(query_results, excel_file)

    # send email
    email_subject = "Somerville Circulating Tech"
    email_message = """***This is an automated email***


The Somerville Circulating Tech report has been attached."""
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
        email_subject = "Somerville Circulating Tech script error"
        email_message = (
            "Your script failed with the following error:\n\n" + traceback.format_exc()
        )

        send_email_error(email_subject, email_message, emailto)
        raise
    
