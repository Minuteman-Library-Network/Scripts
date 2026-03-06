#!/usr/bin/env python3

"""
Jeremy Goldstein
Minuteman Library Network

Generates daily counts of items that could be paged for Wayland
Updates data in a Google sheet used as a data source by Looker Studio
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

# from oauth2client.service_account import ServiceAccountCredentials
# from googleapiclient.discovery import build


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

    first_empty_row = len(wks.get_all_values(include_tailing_empty_rows=False)) + 1
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

    # query to gather paged items at Wayland
    query = """
    WITH transit AS (
	  SELECT
		i.id AS id,
		i.location_code,
		SUBSTRING(SPLIT_PART(SPLIT_PART(v.field_content,'from ',2),' to',1)FROM 1 FOR 3) AS origin_loc,
		SUBSTRING(SPLIT_PART(v.field_content,'to ',2) FROM 1 FOR 3) AS destination_loc
	
	  FROM sierra_view.item_record i
	  JOIN sierra_view.varfield v
		ON i.id = v.record_id AND v.varfield_type_code = 'm' AND v.field_content LIKE '%IN TRANSIT%'
	  JOIN sierra_view.hold h
		ON i.id = h.record_id AND h.status = 't'

      WHERE i.item_status_code = 't'
	    AND (SUBSTRING(SPLIT_PART(SPLIT_PART(v.field_content,'from ',2),' to',1)FROM 1 FOR 3) = 'wyl'  OR SUBSTRING(SPLIT_PART(v.field_content,'to ',2) FROM 1 FOR 3) = 'wyl')
    )

    SELECT
	  TO_CHAR(CURRENT_DATE,'YYYY-MM-DD') AS "date",
	  l.name AS library,
	  COUNT(DISTINCT t.id) FILTER (WHERE t.origin_loc = l.code AND TO_TIMESTAMP(SPLIT_PART(v.field_content,': ',1),'DY MON DD YYYY HH:MIAM')::DATE != i.last_checkin_gmt::DATE) AS transit_from,
	  COUNT(DISTINCT t.id) FILTER (WHERE t.destination_loc = l.code AND t.location_code ~ '^wyl') AS transit_to,
	  CASE
		WHEN l.name != 'WAYLAND' THEN 0
		ELSE(
		  SELECT 
		    COUNT(h.id)
		  FROM sierra_view.hold h
		  JOIN sierra_view.item_record i
			ON h.record_id = i.id AND i.location_code ~ '^wyl'
			AND SUBSTRING(h.pickup_location_code,1,3) = 'wyl'
			AND h.status IN ('b','i')
			AND h.on_holdshelf_gmt::DATE = CURRENT_DATE
			--exclude items that were checked in instead of being paged
			AND abs(EXTRACT(EPOCH FROM (i.last_checkin_gmt - h.on_holdshelf_gmt))) > 5
			--exclude new items scanned to trigger initial holds
		  JOIN sierra_view.record_metadata rm
			ON i.id = rm.id
			AND h.on_holdshelf_gmt::DATE - rm.creation_date_gmt::DATE >=4
		) END AS placed_on_holdshelf

    FROM sierra_view.location_myuser l
    JOIN transit t
	  ON (l.code = t.origin_loc OR l.code = t.destination_loc)
      AND l.code NOT IN ('trn','mti')
    JOIN sierra_view.item_record i
	  ON t.id = i.id
    JOIN sierra_view.varfield v
	  ON t.id = v.record_id AND v.varfield_type_code = 'm' AND v.field_content LIKE '%IN TRANSIT%'

    GROUP BY 1,2,l.code
    ORDER BY 1,2
"""

    # run query and append results to specified GSheet
    results = runquery(query)
    parsed_results = parse_pg_data(results)
    appendToSheet(config["gsheet"]["wayland_paged_item_count"], parsed_results)


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
        email_subject = "Wayland Paged Items script error"
        email_message = (
            "Your script failed with the following error:\n\n" + traceback.format_exc()
        )

        send_email_error(email_subject, email_message, emailto)
        raise
