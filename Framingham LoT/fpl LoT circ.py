#!/usr/bin/env python3

#Run in py313
"""
Jeremy Goldstein
Minuteman Library Network

Generates monthly report on the circulation of library of things materials owned by Framingham
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

    #Creating the Excel file for staff
    workbook = xlsxwriter.Workbook(excel_file)
    worksheet = workbook.add_worksheet()


    #Formatting our Excel worksheet
    worksheet.set_landscape()
    worksheet.hide_gridlines(0)

    #Formatting Cells
    eformat= workbook.add_format({'text_wrap': True, 'valign': 'top'})
    eformatlabel= workbook.add_format({'text_wrap': True, 'valign': 'top', 'bold': True})
    eformat2= workbook.add_format({'num_format': 'mm/dd/yy hh:mm:ss'})
    eformat3= workbook.add_format({'num_format': 'mm/dd/yy'})


    # Setting the column widths
    worksheet.set_column(0,0,16.29)
    worksheet.set_column(1,1,10.29)
    worksheet.set_column(2,2,15.29)
    worksheet.set_column(3,3,68.43)
    worksheet.set_column(4,4,8.86)
    worksheet.set_column(5,5,10.14)
    worksheet.set_column(6,6,10.14)
    worksheet.set_column(7,7,15.43)
    worksheet.set_column(8,8,48.14)
    worksheet.set_column(9,9,15.29)
    worksheet.set_column(10,10,11.14)
    worksheet.set_column(11,11,20.86)
    worksheet.set_column(12,12,10.57)
    worksheet.set_column(13,13,9.43)
    worksheet.set_column(14,14,13.00)
    worksheet.set_column(15,15,13.57)
    worksheet.set_column(16,16,34.14)

    # Inserting a header
    worksheet.set_header('Framingham Monthly Library Of Things Circulation')

    # Adding column labels
    worksheet.write(0,0,'Transaction_time', eformatlabel)
    worksheet.write(0,1,'Application', eformatlabel)
    worksheet.write(0,2,'Transaction_type', eformatlabel)
    worksheet.write(0,3,'Best_title', eformatlabel)
    worksheet.write(0,4,'Mat_type', eformatlabel)
    worksheet.write(0,5,'Bib_num', eformatlabel)
    worksheet.write(0,6,'Item_num', eformatlabel)
    worksheet.write(0,7,'Barcode', eformatlabel)
    worksheet.write(0,8,'Call_number_norm', eformatlabel)
    worksheet.write(0,9,'Stat_group_num', eformatlabel)
    worksheet.write(0,10,'Due_date', eformatlabel)
    worksheet.write(0,11,'Count_type_code_num', eformatlabel)
    worksheet.write(0,12,'Itype_code', eformatlabel)
    worksheet.write(0,13,'Scat_code', eformatlabel)
    worksheet.write(0,14,'Location_code', eformatlabel)
    worksheet.write(0,15,'Loanrule_num', eformatlabel)
    worksheet.write(0,16,'Patron_encrypted', eformatlabel)


    # Writing the report for staff to the Excel worksheet
    for rownum, row in enumerate(query_results):
        worksheet.write(rownum+1,0,row[0], eformat2)
        worksheet.write(rownum+1,1,row[1], eformat)
        worksheet.write(rownum+1,2,row[2], eformat)
        worksheet.write(rownum+1,3,row[3], eformat)
        worksheet.write(rownum+1,4,row[4], eformat)
        worksheet.write(rownum+1,5,row[5], eformat)
        worksheet.write(rownum+1,6,row[6], eformat)
        worksheet.write(rownum+1,7,row[7], eformat)
        worksheet.write(rownum+1,8,row[8], eformat)
        worksheet.write(rownum+1,9,row[9], eformat)
        worksheet.write(rownum+1,10,row[10], eformat3)
        worksheet.write(rownum+1,11,row[11], eformat)
        worksheet.write(rownum+1,12,row[12], eformat)
        worksheet.write(rownum+1,13,row[14], eformat)
        worksheet.write(rownum+1,14,row[14], eformat)
        worksheet.write(rownum+1,15,row[15], eformat)
        worksheet.write(rownum+1,16,row[16], eformat)
    
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
    emailto = config_recipient["framingham_lot"]["recipients"].split()
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
      to_char(c.transaction_gmt, 'mm/dd/yyyy HH24:MI:SS'),
      c.application_name,
      CASE
        WHEN c.op_code = 'o' THEN 'checkout'
        WHEN c.op_code = 'i' THEN 'checkin'
        WHEN c.op_code = 'n' THEN 'hold'
        WHEN c.op_code = 'h' THEN 'hold with recall'
        WHEN c.op_code = 'nb' THEN 'bib hold'
        WHEN c.op_code = 'hb' THEN 'hold recall bib'
        WHEN c.op_code = 'ni' THEN 'item hold'
        WHEN c.op_code = 'hi' THEN 'hold recall item'
        WHEN c.op_code = 'nv' THEN 'volume hold'
        WHEN c.op_code = 'hv' THEN 'hold recall volume'
        WHEN c.op_code = 'f' THEN 'filled hold'
        WHEN c.op_code = 'r' THEN 'renewal'
        WHEN c.op_code = 'b' THEN 'booking'
        WHEN c.op_code = 'u' THEN 'use count'
        ELSE 'unknown'
      END AS transaction_type,
      b.best_title,
      b.material_code,
      rmb.record_type_code||rmb.record_num||'a' AS bib_num,
      rmi.record_type_code||rmi.record_num||'a' AS item_num,
      i.barcode,
      i.call_number_norm,
      c.stat_group_code_num,
      c.due_date_gmt::DATE,
      c.count_type_code_num,
      c.itype_code_num,
      c.icode1,
      c.item_location_code,
      c.loanrule_code_num,
      md5(CAST(c.patron_record_id AS varchar))

    FROM sierra_view.circ_trans c
    JOIN sierra_view.bib_record_property b
      ON c.bib_record_id = b.bib_record_id
    JOIN sierra_view.item_record_property i
      ON c.item_record_id = i.item_record_id
    JOIN sierra_view.record_metadata rmi
      ON c.item_record_id = rmi.id
    JOIN sierra_view.record_metadata rmb
      ON c.bib_record_id = rmb.id

    WHERE c.itype_code_num IN ('245', '246', '250', '251', '252', '253')
      AND c.icode1 = '138'
      AND c.item_agency_code_num = '18'
      AND c.transaction_gmt >= CURRENT_DATE - INTERVAL '1 month'

    ORDER BY i.barcode, c.transaction_gmt
    """

    query_results = run_query(query)

    # generate excel file from those query results
    excel_file = "/Scripts/Framingham LoT/Temp Files/fpl LoT circ{}.xlsx".format(date.today())
    excel_writer(query_results, excel_file)

    # send email
    email_subject = "FPL monthly LoT Circulation"
    email_message = """***This is an automated email***


The fpl LoT circ report has been attached."""
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
        email_subject = "Framingham LoT script error"
        email_message = (
            "Your script failed with the following error:\n\n" + traceback.format_exc()
        )

        send_email_error(email_subject, email_message, emailto)
        raise