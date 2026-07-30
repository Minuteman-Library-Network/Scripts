#!/usr/bin/env python3

# Run in py313

"""
Gather the current in house use count for a few specified item records
and emails that data to designated staff monthly.

Author: Jeremy Goldstein
Contact Info: jgoldstein@minlib.net
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

# function constructs and sends outgoing email given a subject, a recipient and body text in both txt and html forms
def send_email(subject, message, recipient):
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
              record_metadata.record_type_code||record_metadata.record_num||'a' AS item_rec_number, 
              bib_record_property.best_title AS title, 
              item_record.use3_count AS item_use_3
            FROM sierra_view.item_record
            JOIN sierra_view.record_metadata
              ON sierra_view.item_record.id = record_metadata.id
            JOIN sierra_view.bib_record_item_record_link
              ON item_record.id = bib_record_item_record_link.item_record_id
            JOIN sierra_view.bib_record_property
              ON bib_record_item_record_link.bib_record_id = bib_record_property.bib_record_id
            WHERE record_metadata.record_type_code||record_metadata.record_num IN ('i19472209','i19480658')
            """
    query_results = run_query(query)

    # Creating the email message
    # read config file with recipient list for email
    config_recipient = configparser.ConfigParser()
    config_recipient.read("C:\\Scripts\\Creds\\emails.ini")
    email_to = config_recipient["needham_puzzles_and_games"]["recipients"].split()
    email_subject = "Count Use for Puzzles and Board Games"
    email_message = """{}
    """.format(date.today())
    for rownum, row in enumerate(query_results):
        email_message += """
{} {}: Item Use 3 = {}
""".format(str(row[0]), str(row[1]), str(row[2]))
    email_message += """
***This is an automated email. Replies to minuteman@minlib.net will not be seen***"""

    send_email(email_subject, email_message, email_to)

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
        email_subject = "Needham puzzles and games script error"
        email_message = (
            "Your script failed with the following error:\n\n" + traceback.format_exc()
        )

        send_email(email_subject, email_message, emailto)
        raise
