#!/usr/bin/env python3

# Run in py313

"""
Script to automate Create and ftp weekly files 
to Orange Boy for Savannah platform

Author: Jeremy Goldstein
Contact Info: jgoldstein@minlib.net
"""

import psycopg2
import csv
import os
import pysftp
import configparser
from datetime import date
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email.utils import formatdate
from email import encoders
import traceback

# function takes a sql query as a parameter, connects to a database and returns the results
def run_query(query, csv_file):
    # read config file with Sierra login credentials
    config = configparser.ConfigParser()
    config.read("C:\\Scripts\\Creds\\config.ini")

    # Connecting to Sierra PostgreSQL database
    try:
        conn = psycopg2.connect(config["sql"]["connection_string"])
    except psycopg2.Error as e:
        print("Unable to connect to database: " + str(e))

    # Opening a session and querying the database
    cursor = conn.cursor()
    cursor.execute(query)
    # splitting out header information from query results
    headers = [i[0] for i in cursor.description]
    rows = cursor.fetchall()
    conn.close()

    # run csvWriter function to populate csv_file based on query results
    end_file = csv_writer(rows, headers, csv_file)

    return end_file

# populate csv file with results of a sql query
def csv_writer(query_results, headers, csv_file):

    with open(csv_file, "w", encoding="utf-8", newline="") as tempFile:
        myFile = csv.writer(tempFile, delimiter=",")
        myFile.writerow(headers)
        myFile.writerows(query_results)
    tempFile.close()

    return csv_file

# function to sftp a specified file
def sftp_file(file, library):
    """
    config.ini contains data like the following
    [orangeboy]
    host = ftp.xxx.xxx
    user_abc = username
    key_abc = C:/users//MyUser//.ssh/keyfile
    """
    config = configparser.ConfigParser()
    config.read("C:\\Scripts\\Creds\\config.ini")

    # set connection option to disable check for host key
    cnopts = pysftp.CnOpts()
    cnopts.hostkeys = None
    
    # open sftp connection
    srv = pysftp.Connection(
        host=config["orangeboy"]["host"],
        username=config["orangeboy"]["user_" + library],
        private_key=config["orangeboy"]["key_" + library],
        cnopts=cnopts,
    )
    # upload specified file to root directory
    srv.put(file)

    # close connection
    srv.close()
    # remove local copy when done
    os.remove(file)


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


#redo from here and figure out private key
def main(library):
    try:
        circ_query = open(library + "_circulation.sql","r").read()
        patron_query = open(library + "_patron.sql","r").read()

        circ_file_name = "/Scripts/Orange Boy/Temp Files/" + library + "_circulation_{}.csv".format(date.today().strftime('%m-%d-%Y'))
        patron_file_name = "/Scripts/Orange Boy/Temp Files/" + library + "_patron_{}.csv".format(date.today().strftime('%m-%d-%Y'))

        circ_file_csv = run_query(circ_query,circ_file_name)
        sftp_file(circ_file_csv,library)
        patron_file_csv = run_query(patron_query,patron_file_name)
        sftp_file(patron_file_csv,library)

    except Exception:
        # read config file with recipient list for email
        config_recipient = configparser.ConfigParser()
        config_recipient.read("C:\\Scripts\\Creds\\emails.ini")
        emailto = config_recipient["script_error"]["recipients"].split()

        # craft email subject and message containing error message details from traceback
        email_subject = "Orange Boy " + library + " script error"
        email_message = (
            "Your script failed with the following error:\n\n" + traceback.format_exc()
        )

        send_email_error(email_subject, email_message, emailto)
        raise

main('NEWTON')
main('NATICK')
main('BROOKLINE')

