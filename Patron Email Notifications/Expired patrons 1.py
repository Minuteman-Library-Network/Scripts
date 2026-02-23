#!/usr/bin/env python3

# Run in py38

"""
Jeremy Goldstein
Minuteman Library Network
Generate and send email notification to patrons with whose library cards expired yesterday
"""

import psycopg2
import smtplib
import configparser
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email.utils import formatdate
from email import encoders
from datetime import date
import traceback


def run_query(query):
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
    # For now, just storing the data in a variable. We'll use it later.
    rows = cursor.fetchall()
    conn.close()
    return rows


# function constructs and sends outgoing email given a subject, a recipient and body text in both txt and html forms
def send_email(subject, message_text, message_html, recipient):
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

    # Creating the email message with html and plaintxt options
    msg = MIMEMultipart("alternative")
    part1 = MIMEText(message_text, "plain")
    part2 = MIMEText(message_html, "html")
    msg["From"] = emailfrom
    if type(recipient) is list:
        msg["To"] = ", ".join(recipient)
    else:
        msg["To"] = recipient
    msg["Date"] = formatdate(localtime=True)
    msg["Subject"] = subject
    msg.attach(part1)
    msg.attach(part2)

    # Sending the email message
    smtp = smtplib.SMTP(emailhost, emailport)
    # for Gmail connection used within Minuteman
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
    query = """
      --Find patrons in the first quarter of MLN libraries whose cards expired yesterday
      SELECT
        MIN(n.first_name),
        MIN(n.last_name),
        MIN(v.field_content) as email,
        TO_CHAR(p.expiration_date_gmt,'Mon DD, YYYY'),
        p.id
      FROM sierra_view.patron_record as p
      JOIN sierra_view.varfield v		
        ON p.id = v.record_id
        AND v.varfield_type_code = 'z'
      JOIN sierra_view.patron_record_fullname n
        ON p.id = n.patron_record_id
      WHERE p.expiration_date_gmt::DATE = (CURRENT_DATE - INTERVAL '1 day')
      AND p.ptype_code IN('1', '2', '3', '4', '5', '6', '7', '8', '10', '11', '12', '110', '301', '302', '303', '304', '305', '306', '307', '308', '310', '311', '312') 
      GROUP BY 5, 4
      """

    query_results = run_query(query)

    for rownum, row in enumerate(query_results):

        # emailto can send to multiple addresses by separating emails with commas
        emailto = [str(row[2])]
        emailsubject = "Your library card has expired"
        # Creating the email message
        email_text = """Dear {} {},
       
Your library card expired on {}.  Visit or contact any Minuteman library for information on how to renew your account and regain access to over 5 million items.
      
***This is an automated email***""".format(
            str(row[0]), str(row[1]), str(row[3])
        )

        email_html = """
    <html>
    <head></head>
    <body style="background-color:#FFFFFF;">
    <table style="width: 70%; margin-left: 15%; margin-right: 15%; border: 0; cellspacing: 0; cellpadding: 0; background-color: #FFFFFF;">
    <tr>
    <img src="https://www.minlib.net/sites/default/files/glazed_builder_images/clock.png" style="height: 135px; width: 135px; display: block; margin-left: auto; margin-right: auto;" alt="placeholder">
    <font face="Scala Sans, Calibri, Arial"; size="3">
    <p>Dear {} {},<br><br>
    
    Your library card expired on {}.<br>
    Visit or contact any <a href="https://www.minlib.net/our-libraries">Minuteman library</a> for information on how to renew your account and regain access to over 5 million items.<br><br>
    ***This is an automated email.  Do not reply.***<br><br>
    </font>
    </p>
    <img src="https://www.minlib.net/sites/default/files/glazed_builder_images/logo-print-small.jpg" style="height: 32px; width: 188px; display: block; margin-left: auto; margin-right: auto;" alt="Minuteman logo">
    </tr>
    </table>
    </body>
    </html>""".format(
            str(row[0]), str(row[1]), str(row[3])
        )
        send_email(emailsubject, email_text, email_html, emailto)


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
        email_subject = "expired patrons 1 script error"
        email_message = (
            "Your script failed with the following error:\n\n" + traceback.format_exc()
        )

        send_email_error(email_subject, email_message, emailto)
        raise
