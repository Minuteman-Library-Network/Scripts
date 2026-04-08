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
      AND NOW() - rmi.record_last_updated_gmt < INTERVAL '18 hours'
    """

    renewals_query = """
    SELECT 
      rmp.record_type_code||rmp.record_num||'a' AS patron_no,
      REPLACE(i.barcode,' ','') AS item_barcode,
      TRIM(TRAILING '/' FROM b.best_title) AS title,
      TO_CHAR(c.due_gmt,'MM-DD-YYYY') AS due_date,
      rmi.record_type_code||rmi.record_num ||'a' AS item_no,
      ROUND(p.owed_amt,2) AS money_owed,
      c.loanrule_code_num AS loan_rule,
      NULLIF(COUNT(ih.id),0) AS item_holds,
      NULLIF(COUNT(bh.id),0) AS bib_holds,
      c.renewal_count AS renewals,
      rmb.record_type_code||rmb.record_num||'a' AS bib_no,
      c.id AS checkout_id
  
    FROM sierra_view.checkout c
    JOIN sierra_view.patron_record p
      ON c.patron_record_id = p.id
    JOIN sierra_view.record_metadata rmp
      ON p.id = rmp.id
    JOIN sierra_view.item_record_property i
      ON c.item_record_id = i.item_record_id
    JOIN sierra_view.record_metadata rmi
      ON i.item_record_id = rmi.id
    JOIN sierra_view.bib_record_item_record_link l
      ON i.item_record_id = l.item_record_id
    JOIN sierra_view.bib_record_property b
      ON l.bib_record_id = b.bib_record_id
    JOIN sierra_view.record_metadata rmb
      ON l.bib_record_id = rmb.id
    LEFT JOIN sierra_view.hold bh
      ON l.bib_record_id = bh.record_id 
    LEFT JOIN sierra_view.hold ih
      ON i.id = ih.record_id AND ih.status = '0'         

    WHERE c.due_gmt::DATE - CURRENT_DATE BETWEEN 0 AND 3
    GROUP BY 1,2,3,4,5,6,7,10,11,12
    ORDER BY patron_no
    """

    overdues_query = """
    SELECT 
      rmp.record_type_code||rmp.record_num||'a' AS patron_no,
      REPLACE(i.barcode,' ','') AS item_barcode,
      TRIM(TRAILING '/' FROM b.best_title) AS title,
      TO_CHAR(c.due_gmt,'MM-DD-YYYY') AS due_date,
      rmi.record_type_code||rmi.record_num ||'a' AS item_no,
      ROUND(p.owed_amt,2) AS money_owed,
      c.loanrule_code_num AS loan_rule,
      NULLIF(COUNT(ih.id),0) AS item_holds,
      NULLIF(COUNT(bh.id),0) AS bib_holds,
      c.renewal_count AS renewals,
      rmb.record_type_code||rmb.record_num||'a' AS bib_no,
      c.id AS checkout_id
  
    FROM sierra_view.checkout c
    JOIN sierra_view.patron_record p
      ON c.patron_record_id = p.id
    JOIN sierra_view.record_metadata rmp
      ON p.id = rmp.id
    JOIN sierra_view.item_record_property i
      ON c.item_record_id = i.item_record_id
    JOIN sierra_view.record_metadata rmi
      ON i.item_record_id = rmi.id
    JOIN sierra_view.bib_record_item_record_link l
      ON i.item_record_id = l.item_record_id
    JOIN sierra_view.bib_record_property b
      ON l.bib_record_id = b.bib_record_id
    JOIN sierra_view.record_metadata rmb
      ON l.bib_record_id = rmb.id
    LEFT JOIN sierra_view.hold bh
      ON l.bib_record_id = bh.record_id 
    LEFT JOIN sierra_view.hold ih
      ON i.id = ih.record_id AND ih.status = '0'        
  
    WHERE CURRENT_DATE - c.due_gmt::DATE BETWEEN 1 AND 30
    GROUP BY 1,2,3,4,5,6,7,10,11,12
    ORDER BY patron_no
    """

    holds = run_query(holds_query)
    renews = run_query(renewals_query)
    overdues = run_query(overdues_query)

    holds_file_name = "/Scripts/Shoutbomb/Temp Files/holds{}.txt".format(
        datetime.now().strftime("%Y%m%d-%H%M")
    )
    renews_file_name = "/Scripts/Shoutbomb/Temp Files/renew{}.txt".format(
        datetime.now().strftime("%Y%m%d-%H%M")
    )
    overdues_file_name = "/Scripts/Shoutbomb/Temp Files/overdue{}.txt".format(
        datetime.now().strftime("%Y%m%d-%H%M")
    )

    csv_writer(holds, holds_file_name)
    csv_writer(renews, renews_file_name)
    csv_writer(overdues, overdues_file_name)

    os.system("c:\\Scripts\\Shoutbomb\\ftp_all.bat")


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
        email_subject = "shoutbomb all script error"
        email_message = (
            "Your script failed with the following error:\n\n" + traceback.format_exc()
        )

        send_email_error(email_subject, email_message, emailto)
        raise
