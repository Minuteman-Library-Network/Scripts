#!/usr/bin/env python3

# run in py313

"""
Create and email weekly custom purchase alerts
tailored for needs of Concord's and YA collections

"""

import psycopg2
import xlsxwriter
import os
import configparser
import sys
import time
from datetime import date
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email.utils import formatdate
from email import encoders
import traceback

# run sql query against Sierra database and return results
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

#convert sql query results into formatted excel file
def excel_writer(query_results,excel_file):

    #Creating the Excel file for staff
    workbook = xlsxwriter.Workbook(excel_file,{'remove_timezone': True})
    worksheet = workbook.add_worksheet()


    #Formatting our Excel worksheet
    worksheet.set_landscape()
    worksheet.hide_gridlines(0)

    #Formatting Cells
    eformat= workbook.add_format({'text_wrap': True, 'valign': 'top'})
    eformatlabel= workbook.add_format({'text_wrap': True, 'valign': 'top', 'bold': True})
    link_format = workbook.add_format({'color': 'blue', 'underline': 1})

    # Setting the column widths
    worksheet.set_column(0,0,14.14)
    worksheet.set_column(1,1,42.43)
    worksheet.set_column(2,2,26.57)
    worksheet.set_column(3,3,10.14)
    worksheet.set_column(4,4,13)
    worksheet.set_column(5,5,9.14)
    worksheet.set_column(6,6,12.71)
    worksheet.set_column(7,7,8.71)
    worksheet.set_column(8,8,12.43)
    worksheet.set_column(9,9,12.71)
    worksheet.set_column(10,10,9.57)
    worksheet.set_column(11,11,8.57)
    worksheet.set_column(12,12,12.43)

    #Inserting a header
    worksheet.set_header('Purchase Alert')

    # Adding column labels
    worksheet.write(0,0,'Record_number', eformatlabel)
    worksheet.write(0,1,'Title', eformatlabel)
    worksheet.write(0,2,'Author', eformatlabel)
    worksheet.write(0,3,'PublicationYear', eformatlabel)
    worksheet.write(0,4,'MatType', eformatlabel)
    worksheet.write(0,5,'TotalItemCount', eformatlabel)
    worksheet.write(0,6,'TotalAvailableItemCount', eformatlabel)
    worksheet.write(0,7,'TotalHoldCount', eformatlabel)
    worksheet.write(0,8,'TotalDemandRatio', eformatlabel)
    worksheet.write(0,9,'LocalAvailableItemCount', eformatlabel)
    worksheet.write(0,10,'LocalOrderCopies', eformatlabel)
    worksheet.write(0,11,'LocalHoldCount', eformatlabel)
    worksheet.write(0,12,'LocalDemandRatio', eformatlabel)


    # Writing the report for staff to the Excel worksheet
    for rownum, row in enumerate(query_results):
        worksheet.write(rownum+1,0,row[0], eformat)
        worksheet.write_url(rownum+1,1,row[13], link_format, row[1])
        worksheet.write(rownum+1,2,row[2], eformat)
        worksheet.write(rownum+1,3,row[3], eformat)
        worksheet.write(rownum+1,4,row[4], eformat)
        worksheet.write(rownum+1,5,row[5], eformat)
        worksheet.write(rownum+1,6,row[6], eformat)
        worksheet.write(rownum+1,7,row[7], eformat)
        worksheet.write(rownum+1,8,row[8], eformat)
        worksheet.write(rownum+1,9,row[9], eformat)
        worksheet.write(rownum+1,10,row[10], eformat)
        worksheet.write(rownum+1,11,row[11], eformat)
        worksheet.write(rownum+1,12,row[12], eformat)
    
    workbook.close()
    
    return excel_file



def send_email(subject, message, excelfile1, excelfile2, excelfile3, excelfile4):
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
    emailto = config_recipient["concord_purchase_alert"]["recipients"].split()
    # plain text of email message
    emailmessage = message

    #Creating the email message
    msg = MIMEMultipart()
    msg['From'] = emailfrom
    if type(emailto) is list:
        msg['To'] = ', '.join(emailto)
    else:
        msg['To'] = emailto
    msg['Date'] = formatdate(localtime = True)
    msg['Subject'] = subject
    msg.attach (MIMEText(emailmessage))
    part = MIMEBase('application', "octet-stream")
    part.set_payload(open(excelfile1,"rb").read())
    encoders.encode_base64(part)
    part.add_header('Content-Disposition','attachment; filename=%s' % excelfile1.rsplit("/", 1)[-1])
    msg.attach(part)
    part2 = MIMEBase('application', "octet-stream")
    part2.set_payload(open(excelfile2,"rb").read())
    encoders.encode_base64(part2)
    part2.add_header('Content-Disposition','attachment; filename=%s' % excelfile2.rsplit("/", 1)[-1])
    msg.attach(part2)
    part3 = MIMEBase('application', "octet-stream")
    part3.set_payload(open(excelfile3,"rb").read())
    encoders.encode_base64(part3)
    part3.add_header('Content-Disposition','attachment; filename=%s' % excelfile3.rsplit("/", 1)[-1])
    msg.attach(part3)
    part4 = MIMEBase('application', "octet-stream")
    part4.set_payload(open(excelfile4,"rb").read())
    encoders.encode_base64(part4)
    part4.add_header('Content-Disposition','attachment; filename=%s' % excelfile4.rsplit("/", 1)[-1])
    msg.attach(part4)

    #Sending the email message
    smtp = smtplib.SMTP(emailhost, emailport)
    #for Google connection
    smtp.ehlo()
    smtp.starttls()
    smtp.login(emailuser, emailpass)
    smtp.sendmail(emailfrom, emailto, msg.as_string())
    smtp.quit()# function constructs and sends outgoing email given a subject, a recipient and body text in both txt and html forms

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
	
    #Produce 4 excel files to Concord's specs
    query_results1 = run_query(open("PurchaseAlertconj.sql", "r").read())
    excel_file1 =  "/Scripts/Concord Purchase Alert/Temp Files/conjPurchaseAlert{}.xlsx".format(date.today())
    excel_writer(query_results1, excel_file1)
    query_results2 = run_query(open("PurchaseAlertco2j.sql", "r").read())
    excel_file2 =  "/Scripts/Concord Purchase Alert/Temp Files/co2jPurchaseAlert{}.xlsx".format(date.today())
    excel_writer(query_results2, excel_file2)
    query_results3 = run_query(open("PurchaseAlertcony.sql", "r").read())
    excel_file3 =  "/Scripts/Concord Purchase Alert/Temp Files/conyPurchaseAlert{}.xlsx".format(date.today())
    excel_writer(query_results3, excel_file3)
    query_results4 = run_query(open("PurchaseAlertco2y.sql", "r").read())
    excel_file4 =  "/Scripts/Concord Purchase Alert/Temp Files/co2yPurchaseAlert{}.xlsx".format(date.today())
    excel_writer(query_results4, excel_file4)
    
    # send email
    email_subject = "Concord Purchase Alerts"
    email_message = """***This is an automated email***


The Concord Purchase alerts have been attached."""
    send_email(email_subject, email_message, excel_file1, excel_file2, excel_file3, excel_file4)

    # delete local files
    os.remove(excel_file1)
    os.remove(excel_file2)
    os.remove(excel_file3)
    os.remove(excel_file4)


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
        email_subject = "Concord Purchase Alert script error"
        email_message = (
            "Your script failed with the following error:\n\n" + traceback.format_exc()
        )

        send_email_error(email_subject, email_message, emailto)
        raise




