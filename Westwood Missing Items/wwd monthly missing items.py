#!/usr/bin/env python3

#Run in py313
"""
Jeremy Goldstein
Minuteman Library Network

Generates monthly missing and lost in transit reports for Westwood
Creates two excel files and emails them to designated staff as attachments
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

def excel_writer(query_results, excel_file1, excel_file2):
    #Creating the Excel file for staff
    workbook = xlsxwriter.Workbook(excel_file1, {'remove_timezone': True})
    workbook2 = xlsxwriter.Workbook(excel_file2,{'remove_timezone': True})

    worksheet1 = workbook.add_worksheet('WWD')
    worksheet2 = workbook.add_worksheet('WW2')

    worksheet3 = workbook2.add_worksheet('WWD')
    worksheet4 = workbook2.add_worksheet('WW2')

    # Formatting our Excel worksheet
    worksheet1.set_landscape()
    worksheet1.hide_gridlines(0)
    worksheet2.set_landscape()
    worksheet2.hide_gridlines(0)
    worksheet3.set_landscape()
    worksheet3.hide_gridlines(0)
    worksheet4.set_landscape()
    worksheet4.hide_gridlines(0)

    # Formatting Cells
    eformat= workbook.add_format({'text_wrap': True, 'valign': 'top'})
    eformatlabel= workbook.add_format({'text_wrap': True, 'valign': 'top', 'bold': True})
    eformatdate= workbook.add_format({'num_format': 'mm/dd/yy'})

    eformat2= workbook2.add_format({'text_wrap': True, 'valign': 'top'})
    eformatlabel2= workbook2.add_format({'text_wrap': True, 'valign': 'top', 'bold': True})
    eformatdate2= workbook2.add_format({'num_format': 'mm/dd/yy'})

    # Setting the column widths
    worksheet1.set_column(0,0,73.43)
    worksheet1.set_column(1,1,39.14)
    worksheet1.set_column(2,2,14.43)
    worksheet1.set_column(3,3,17.29)
    worksheet2.set_column(0,0,73.43)
    worksheet2.set_column(1,1,39.14)
    worksheet2.set_column(2,2,14.43)
    worksheet2.set_column(3,3,17.29)
    worksheet3.set_column(0,0,73.43)
    worksheet3.set_column(1,1,39.14)
    worksheet3.set_column(2,2,14.43)
    worksheet3.set_column(3,3,17.29)
    worksheet4.set_column(0,0,73.43)
    worksheet4.set_column(1,1,39.14)
    worksheet4.set_column(2,2,14.43)
    worksheet4.set_column(3,3,17.29)

    # Inserting a header
    worksheet1.set_header('Westwood Missing Items')
    worksheet2.set_header('Westwood Missing Items')
    worksheet3.set_header('Westwood In Transit Items')
    worksheet3.set_header('Westwood In Transit Items')

    # Adding column labels
    worksheet1.write(0,0,'Title', eformatlabel)
    worksheet1.write(0,1,'Call_Number', eformatlabel)
    worksheet1.write(0,2,'Barcode', eformatlabel)
    worksheet1.write(0,3,'Last_Updated_Date', eformatlabel)
    worksheet2.write(0,0,'Title', eformatlabel)
    worksheet2.write(0,1,'Call_Number', eformatlabel)
    worksheet2.write(0,2,'Barcode', eformatlabel)
    worksheet2.write(0,3,'Last_Updated_Date', eformatlabel)
    worksheet3.write(0,0,'Title', eformatlabel)
    worksheet3.write(0,1,'Call_Number', eformatlabel)
    worksheet3.write(0,2,'Barcode', eformatlabel)
    worksheet3.write(0,3,'Last_Updated_Date', eformatlabel)
    worksheet4.write(0,0,'Title', eformatlabel)
    worksheet4.write(0,1,'Call_Number', eformatlabel)
    worksheet4.write(0,2,'Barcode', eformatlabel)
    worksheet4.write(0,3,'Last_Updated_Date', eformatlabel)

    # Writing the report for staff to the Excel worksheet
    row1 = 1
    row2 = 1
    row3 = 1
    row4 = 1

    for rownum, row in enumerate(query_results):
        if row[0] == 'm' and row[1] == 'wwd':
            worksheet1.write(row1,0,row[2], eformat)
            worksheet1.write(row1,1,row[3], eformat)
            worksheet1.write(row1,2,row[4], eformat)
            worksheet1.write(row1,3,row[5], eformatdate)
            row1 += 1
        elif row[0] == 'm' and row[1] == 'ww2':
            worksheet2.write(row2,0,row[2], eformat)
            worksheet2.write(row2,1,row[3], eformat)
            worksheet2.write(row2,2,row[4], eformat)
            worksheet2.write(row2,3,row[5], eformatdate)
            row2 += 1
        elif row[0] == 't' and row[1] == 'wwd':
            worksheet3.write(row3,0,row[2], eformat2)
            worksheet3.write(row3,1,row[3], eformat2)
            worksheet3.write(row3,2,row[4], eformat2)
            worksheet3.write(row3,3,row[5], eformatdate2)
            row3 += 1
        elif row[0] == 't' and row[1] == 'ww2':
            worksheet4.write(row4,0,row[2], eformat2)
            worksheet4.write(row4,1,row[3], eformat2)
            worksheet4.write(row4,2,row[4], eformat2)
            worksheet4.write(row4,3,row[5], eformatdate2)
            row4 += 1
    
    workbook.close()
    workbook2.close()
    return excel_file1, excel_file2

# function takes a file as a parameter and attaches that file to an outgoing email
def send_email(subject, message, attachment1, attachment2, recipient):
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
    part.set_payload(open(attachment1, "rb").read())
    encoders.encode_base64(part)
    part.add_header(
        "Content-Disposition", "attachment; filename=%s" % attachment1.rsplit("/", 1)[-1]
    )
    msg.attach(part)
    part2 = MIMEBase('application', "octet-stream")
    part2.set_payload(open(attachment2,"rb").read())
    encoders.encode_base64(part2)
    part2.add_header('Content-Disposition','attachment; filename=%s' % attachment2.rsplit("/", 1)[-1])
    msg.attach(part2)

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
              i.item_status_code,
              SUBSTRING(i.location_code FROM 1 FOR 3),
              b.best_title,
              REPLACE(ip.call_number,'|a', ''),
              ip.barcode,
              m.record_last_updated_gmt
            
            FROM sierra_view.item_record i
            JOIN sierra_view.item_record_property ip
              ON i.id = ip.item_record_id
            JOIN sierra_view.bib_record_item_record_link l
              ON i.id = l.item_record_id
            JOIN sierra_view.bib_record_property b
              ON l.bib_record_id = b.bib_record_id
            JOIN sierra_view.record_metadata m
              ON i.id = m.id
            
            WHERE i.item_status_code IN ('m','t')
              AND i.location_code ~ '^ww'
              AND m.record_last_updated_gmt <= NOW() - INTERVAL '1 month'
            ORDER BY CASE
              WHEN SUBSTRING(i.location_code FROM 4 FOR 1) = 'j' THEN 2
              WHEN SUBSTRING(i.location_code FROM 4 FOR 1) = 'y' THEN 3
              ELSE 1
            END, 4
            """

    query_results = run_query(query)

    # generate excel file from those query results
    excel_file1 = "/Scripts/Westwood Missing Items/Temp Files/wwd missing items{}.xlsx".format(date.today())
    excel_file2 = "/Scripts/Westwood Missing Items/Temp Files/wwd in transit items{}.xlsx".format(date.today())
    excel_file1, excel_file2 = excel_writer(query_results, excel_file1, excel_file2)

    email_subject = 'Westwood Missing Items'
    email_message = '''***This is an automated email***


    The Westwood missing and long in transit reports have been attached.'''
	# read config file with recipient list for email
    config_recipient = configparser.ConfigParser()
    config_recipient.read("C:\\Scripts\\Creds\\emails.ini")
    email_recipient = config_recipient["westwood_missing_items"]["recipients"].split()  
    send_email(email_subject, email_message, excel_file1, excel_file2, email_recipient)

    os.remove(excel_file1)
    os.remove(excel_file2)

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
        email_subject = "Westwood Missing Items Script Error"
        email_message = (
            "Your script failed with the following error:\n\n" + traceback.format_exc()
        )

        send_email_error(email_subject, email_message, emailto)
        raise