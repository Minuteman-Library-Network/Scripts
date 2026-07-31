#!/usr/bin/env python3

#Run in py313
"""
Jeremy Goldstein
Minuteman Library Network

Generates monthly Report on the number of unique patrons
and average checkout per patrons by service location

Report is then emailed to staff as a csv attachment
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
    # open csvfile in write mode and add a row to it for the headers and each line of query_results
    with open(csv_file, "w", encoding="utf-8", newline="") as tempFile:
        myFile = csv.writer(tempFile, delimiter=",")
        myFile.writerow(headers)
        myFile.writerows(query_results)
    tempFile.close()
    # return variable containing the newly created csv file
    return csv_file

# function takes a file as a parameter and attaches that file to an outgoing email
def send_email(subject, message, attachment, recipient):
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
        TO_CHAR(t.transaction_gmt, 'MM-DD-YY') AS transaction_date,
        l.name AS checkout_location,
        COUNT(t.id) AS total_checkouts,
        COUNT(DISTINCT(t.patron_record_id)) AS total_unique_patrons,
        ROUND(COUNT(t.id)::NUMERIC/COUNT(DISTINCT(t.patron_record_id)),2) AS avg_checkouts_per_patron

      FROM sierra_view.circ_trans t
      JOIN sierra_view.statistic_group_myuser s
        ON t.stat_group_code_num = s.code
      JOIN sierra_view.patron_record p
        ON t.patron_record_id = p.id
      JOIN sierra_view.location_myuser l
        ON s.location_code = l."code"
  
      WHERE t.op_code = 'o'
        AND t.transaction_gmt > NOW()::DATE - INTERVAL '1 month'
        AND p.pcode3 = '115'
      GROUP BY 1,2
    """

    query_results, headers = run_query(query)
    
    # generate csv  file from those query results
    csv_file = "/Scripts/Wayland Monthly Reports/Temp Files/wyl_monthly_patrons_by_checkoutloc{}.csv".format(date.today())
    local_file = write_csv(query_results, headers, csv_file)

    # send email
    email_subject = "Monthly WYL Patrons by Checkout Location"
    email_message = """***This is an automated email***
        
    The monthly report of Wayland Patrons by Checkout Location has been attached. 
    """
	# read config file with recipient list for email
    config_recipient = configparser.ConfigParser()
    config_recipient.read("C:\\Scripts\\Creds\\emails.ini")
    email_recipient = config_recipient["wayland_monthly_patrons_by_checkoutloc"]["recipients"].split()  
    send_email(email_subject, email_message, local_file, email_recipient)
    
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
        email_subject = "Wayland Monthly Patrons By Checkout Location script error"
        email_message = (
            "Your script failed with the following error:\n\n" + traceback.format_exc()
        )

        send_email_error(email_subject, email_message, emailto)
        raise
