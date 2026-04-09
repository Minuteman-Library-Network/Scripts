#!/usr/bin/env python3
# run in py313

"""
Jeremy Goldstein
Minuteman Library Network

Gather data required for Shoutbomb service and ftp files to vendor
"""

import psycopg2
import csv
import configparser
import os
from datetime import datetime
from datetime import time
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email.utils import formatdate
from email import encoders
import traceback


# populate csv file with results of a sql query
def csv_writer(query_results, csv_file):

    with open(csv_file, "w", encoding="utf-8", newline="") as tempFile:
        myFile = csv.writer(tempFile, delimiter="|")
        myFile.writerows(query_results)
    tempFile.close()

    return csv_file


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
    holds_query = """
    SELECT 
      TRIM(TRAILING '/' FROM b.best_title) AS title,
      TO_CHAR(rmi.record_last_updated_gmt,'MM-DD-YYYY') AS last_update, 
      rmi.record_type_code||rmi.record_num||'a' AS item_no, 
      rmp.record_type_code||rmp.record_num||'a' AS patron_no, 
      h.pickup_location_code AS pickup_location,
      irp.barcode AS item_barcode,
      h.id AS hold_id

    FROM sierra_view.hold h
    JOIN sierra_view.item_record_property irp
      ON h.record_id = irp.item_record_id
    JOIN sierra_view.record_metadata rmi
      ON irp.item_record_id = rmi.id
    JOIN sierra_view.patron_record p
      ON h.patron_record_id = p.id
    JOIN sierra_view.record_metadata rmp
      ON h.patron_record_id = rmp.id
    JOIN sierra_view.bib_record_item_record_link l 
      ON irp.item_record_id = l.item_record_id
    JOIN sierra_view.bib_record_property b
      ON l.bib_record_id = b.bib_record_id

    WHERE h.status IN ('b','i')
	  AND h.pickup_location_code IS NOT NULL
      AND h.pickup_location_code !~ '^(ca)|(do)|(le)'
      AND NOW() - rmi.record_last_updated_gmt < INTERVAL '4 hours'
    """

    holds = run_query(holds_query)

    holds_file_name = "/Scripts/Shoutbomb/Temp_Files/holds{}.txt".format(
        datetime.now().strftime("%Y%m%d-%H%M")
    )

    csv_writer(holds, holds_file_name)

    os.system("c:\\Scripts\\Shoutbomb\\ftp_holds.bat")


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
        email_subject = "shoutbomb holds only script error"
        email_message = (
            "Your script failed with the following error:\n\n" + traceback.format_exc()
        )

        send_email_error(email_subject, email_message, emailto)
        raise
