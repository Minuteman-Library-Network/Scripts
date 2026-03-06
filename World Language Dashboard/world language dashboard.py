#!/usr/bin/env python3

"""
Jeremy Goldstein
Minuteman Library Network

Generates world language holdings for each library
and updates Google Sheet used as a data source by Looker Studio
"""
# run in py313

import configparser
import psycopg2
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email.utils import formatdate
from email import encoders
import traceback
import datetime
import pygsheets


# function takes a sql query as a parameter, connects to a database and returns the results
def runquery(query):
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


# log items that were corrected to an existing Google Sheet
def appendToSheet(spreadSheetId, data):
    if not data:
        raise ValueError("Data must not be empty.")

    gc = pygsheets.authorize(
        service_file="C:\\Scripts\\Creds\\GSheet updater creds.json"
    )

    sh = gc.open_by_key(spreadSheetId)
    wks = sh.sheet1  # or sh.worksheet_by_title("My Sheet")

    # clear current sheet, retaining header
    header = wks.get_row(1)
    wks.clear()
    wks.update_row(1, header)
    first_empty_row = 2
    rows_needed = first_empty_row + len(data) - 1

    # Expand the sheet if the data would exceed the current grid size
    if rows_needed > wks.rows:
        wks.add_rows(rows_needed - wks.rows)
    wks.update_values(f"A{first_empty_row}", data)


# converts psycopg2 fetchall() output to matrix required by pygsheets
def parse_pg_data(rows):

    def convert(val):
        if val is None:
            return ""
        if isinstance(val, (datetime.date, datetime.datetime)):
            return val.isoformat()  # e.g. "2026-03-04"
        return val  # int, float, str pass through as-is

    return [list(convert(val) for val in row) for row in rows]


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
    # enable creds file for referencing GSheets IDs
    config = configparser.ConfigParser()
    config.read("C:\\Scripts\\Creds\\config.ini")

    # query to gather total fines assessed for each checkout location the prior day
    query = """
      WITH language_limit AS (
        SELECT
          b.language_code,
          COUNT(i.id) AS item_count
  
        FROM sierra_view.item_record i
        JOIN sierra_view.bib_record_item_record_link l
          ON i.id = l.item_record_id
        JOIN sierra_view.bib_record b
          ON l.bib_record_id = b.id

        WHERE b.language_code != 'eng'

        GROUP BY 1
        HAVING COUNT(i.id) >= 100
      )

      SELECT
        n.name AS LANGUAGE,
        loc.name AS location,
        m.name AS FORMAT,
        COUNT(i.id) AS total_items,
        '=Vlookup(A'||ROW_NUMBER() OVER(ORDER BY n.name, loc.name)+1||',translation!$A$2:$C$150,3,false)' AS display_language
  
      FROM sierra_view.bib_record b
      JOIN sierra_view.bib_record_item_record_link l
        ON b.id = l.bib_record_id
      JOIN sierra_view.item_record i
        ON l.item_record_id = i.id
        AND SUBSTRING(i.location_code FROM 1 FOR 3) NOT IN ('','int','hpl','knp')
      JOIN language_limit ll
        ON b.language_code = ll.language_code
      JOIN sierra_view.language_property_myuser n
        ON b.language_code = n.code
      JOIN sierra_view.location_myuser loc
        ON SUBSTRING(i.location_code FROM 1 FOR 3) = loc.code::VARCHAR
      JOIN sierra_view.material_property_myuser m
        ON b.bcode2 = m.code

      GROUP BY 1,2,3
      ORDER BY
      1,2
      """

    # run query and append results to specified GSheet
    results = runquery(query)
    parsed_results = parse_pg_data(results)
    appendToSheet(config["gsheet"]["language_by_location"], parsed_results)

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
        email_subject = "world language dashboard script error"
        email_message = (
            "Your script failed with the following error:\n\n" + traceback.format_exc()
        )

        send_email_error(email_subject, email_message, emailto)
        raise
