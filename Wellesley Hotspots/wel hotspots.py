#!/usr/bin/env python3

# Run in py313

"""
Create and email a list of WEL's hotspots and their current status

Author: Jeremy Goldstein
Contact Info: jgoldstein@minlib.net
"""

import psycopg2
import csv
import configparser
import smtplib
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
    # Gather column headers, which are not included in cursor.fetchall() and store in another variable
    columns = [i[0] for i in cursor.description]
    # close database connection
    conn.close()
    # return variables containing query results and column headers
    return rows, columns

# function takes the results of a query and converts them to a csv file
def write_csv(query_results, headers, csv_file):
    # provide a name for the csv file and save the file to a variable
    

    # open csvfile in write mode and add a row to it for the headers and each line of query_results
    with open(csv_file, "w", encoding="utf-8", newline="") as tempFile:
        myFile = csv.writer(tempFile, delimiter=",")
        myFile.writerow(headers)
        myFile.writerows(query_results)
    tempFile.close()
    # return variable containing the newly created csv file
    return csv_file



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
    emailto = config_recipient["wellesley_hotspots"]["recipients"].split()
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
    # query to identify patron records with incorrect owed_amt fields
    query = r"""
SELECT
  rm.record_type_code||rm.record_num||
    COALESCE(
      CAST(
        NULLIF(
          (
            ( rm.record_num % 10 ) * 2 +
            ( rm.record_num / 10 % 10 ) * 3 +
            ( rm.record_num / 100 % 10 ) * 4 +
            ( rm.record_num / 1000 % 10 ) * 5 +
            ( rm.record_num / 10000 % 10 ) * 6 +
            ( rm.record_num / 100000 % 10 ) * 7 +
            ( rm.record_num / 1000000 % 10  ) * 8 +
            ( rm.record_num / 10000000 ) * 9
          ) % 11,
        10)
      AS CHAR(1))
  ,'x') AS record_number,
  REGEXP_REPLACE(ip.call_number,'^\|a','') AS call_number,
  CASE
	WHEN o.due_gmt IS NOT NULL THEN 'CHECKED OUT'
	ELSE stat.name
  END AS status,
  COALESCE(TO_CHAR(o.due_gmt,'YYYY-MM-DD'),'N/A') AS due_date,
  ''''||ip.barcode AS barcode

FROM sierra_view.item_record i
JOIN sierra_view.item_record_property ip
  ON i.id = ip.item_record_id
JOIN sierra_view.record_metadata rm
  ON i.id = rm.id
JOIN sierra_view.item_status_property_myuser stat
  ON i.item_status_code = stat.code
LEFT JOIN sierra_view.checkout o
  ON i.id = o.item_record_id

WHERE (i.itype_code_num = '257' OR (i.itype_code_num = '5' AND i.icode1 = '200'))
  AND i.item_status_code != 'a'
  AND i.icode2 = '-'
  AND i.location_code ~ '^we'
  AND ip.call_number_norm LIKE '%hotspot%'

ORDER BY 4,3,2
    """

    query_results, headers = run_query(query)

    # generate csv file from those query results
    csv_file = "/Scripts/Wellesley Hotspots/Temp Files/wel hotspots{}.csv".format(date.today())
    local_file = write_csv(query_results, headers, csv_file)

    # send email with attached file  
    email_subject = "wel hotspots"
    email_message = """***This is an automated email***
    
    
    The Wellesley Hotspots report has been attached."""
    send_email(email_subject, email_message, local_file)

    # delete csv file once email has been sent
    os.remove(local_file)


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
        email_subject = "Wellesley Hotspots script error"
        email_message = (
            "Your script failed with the following error:\n\n" + traceback.format_exc()
        )

        send_email_error(email_subject, email_message, emailto)
        raise