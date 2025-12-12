#!/usr/bin/env python3

#Run in py38

"""
Create and email a list of patrons with amt owed discrepancies in their patron records
Due to a mismatch between the patron_record.amt_owed and the actual amount of active fines

Author: Jeremy Goldstein
Contact Info: jgoldstein@minlib.net
"""

import psycopg2
import configparser
import csv
import smtplib
import os
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email.utils import formatdate
from email import encoders
from datetime import date

def run_query(query):
    # read config file with Sierra login credentials
    config = configparser.ConfigParser()
    config.read('C:\\Creds\\config.ini')

    # Connecting to Sierra PostgreSQL database
    try:
        conn = psycopg2.connect(config["sql"]["connection_string"])
    except psycopg2.Error as e:
        print("Unable to connect to database: " + str(e))

    # Opening a session and querying the database
    cursor = conn.cursor()
    cursor.execute(query)
    # For now, just storing the data in a variable. We'll use it later.
    rows = cursor.fetchall()
    # Gather column headers, which are not included in the cursor.fetchall() 
    columns = [i[0] for i in cursor.description]
    conn.close()
    return rows, columns

def write_csv(query_results, headers):
    
    csvfile = 'amt_owed_errors{}.csv'.format(date.today())
    
    with open(csvfile,'w', encoding='utf-8', newline='') as tempFile:
        myFile = csv.writer(tempFile, delimiter=',')
        myFile.writerow(headers)
        myFile.writerows(query_results)
    tempFile.close()
    
    return csvfile

def send_email(attachment):
    # read config file with Sierra login credentials
    config = configparser.ConfigParser()
    config.read('C:\\Creds\\config.ini')
    config_recipient = configparser.ConfigParser()
    config_recipient.read('C:\\Creds\\emails.ini')

    # These are variables for the email that will be sent.
    # Make sure to use your own library's email server (emailhost)
    emailhost = config["email"]["host"]
    # user and pw was not in the original script, necessary for Minuteman's Gmail accounts
    emailuser = config["email"]["user"]
    emailpass = config["email"]["pw"]
    emailport = config["email"]["port"]
    emailsubject = "monthly amount owed errors report"
    emailmessage = """***This is an automated email***
    
    
    The monthly amt owed errors has been attached."""

    # Enter your own email information
    emailfrom = config["email"]["sender"]
    emailto = config_recipient["amt_owed_errors"]["recipients"].split()

    # Creating the email message
    msg = MIMEMultipart()
    msg["From"] = emailfrom
    if type(emailto) is list:
        msg["To"] = ", ".join(emailto)
    else:
        msg["To"] = emailto
    msg["Date"] = formatdate(localtime=True)
    msg["Subject"] = emailsubject
    msg.attach(MIMEText(emailmessage))
    part = MIMEBase("application", "octet-stream")
    part.set_payload(open(attachment, "rb").read())
    encoders.encode_base64(part)
    part.add_header("Content-Disposition", "attachment; filename=%s" % attachment)
    msg.attach(part)

    # Sending the email message
    smtp = smtplib.SMTP(emailhost, emailport)
    # for Gmail connection used within Minuteman
    smtp.ehlo()
    smtp.starttls()
    smtp.login(emailuser, emailpass)
    smtp.sendmail(emailfrom, emailto, msg.as_string())
    smtp.quit()

def main():
	query = """\
           SELECT
             rm.record_type_code||rm.record_num || 'a' AS Patron_ID,
             p.owed_amt::MONEY AS owed_amt,
             SUM(COALESCE(f.item_charge_amt, 0.00) + COALESCE(f.processing_fee_amt, 0.00) + COALESCE(f.billing_fee_amt, 0.00) - COALESCE(f.paid_amt, 0.00))::MONEY AS TotalFines
           FROM sierra_view.record_metadata rm
           JOIN sierra_view.patron_record p
             ON p.id = rm.id
           LEFT JOIN sierra_view.fine f
             ON f.patron_record_id = p.id
           GROUP BY 1,2,p.owed_amt
           HAVING p.owed_amt != SUM(COALESCE(f.item_charge_amt, 0.00) + COALESCE(f.processing_fee_amt, 0.00) + COALESCE(f.billing_fee_amt, 0.00) - COALESCE(f.paid_amt, 0.00))
           """
	  
	query_results, headers = run_query(query)
	local_file = write_csv(query_results, headers)
	send_email(local_file)       
	os.remove(local_file)
    
main()
