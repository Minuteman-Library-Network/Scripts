"""
Jeremy Goldstein
Minuteman Library Network

Script identifies instances in which a checkin failed, resulting in a Sierra item being simulataneously checked out an in transit
Once identified, script uses the Sierra API to check the item in again and clear the error
"""

# run in simian

from sierra_ils_utils import SierraAPI
import json
import configparser
import psycopg2
import os
from oauth2client.service_account import ServiceAccountCredentials
from googleapiclient.discovery import build
import gspread
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email.utils import formatdate
from email import encoders
import traceback


# function initializes a session using the Sierra API
def init_api():
    config = configparser.ConfigParser()
    config.read("C:\\Scripts\\Creds\\config.ini")
    """
    .ini file contains url/key/secret for the api in the following form
    [api]
    base_url = https://[local domain]/iii/sierra-api/v6
    client_key = [enter Sierra API key]
    client_secret = [enter Sierra API secret]
    """

    base_url = config["api"]["base_url"] + "/"
    # note sierra-ils-utils assumes base_url contains the trailing /, which the file I have been using did not contain so it is appended here
    client_key = config["api"]["client_key"]
    client_secret = config["api"]["client_secret"]

    # launch SierraAPI session
    sierra_api = SierraAPI(base_url, client_key, client_secret)
    sierra_api.request("GET", "info/token")

    return sierra_api


# function takes a sql query as a parameter, connects to a database and returns the results
def runquery(query):
    config = configparser.ConfigParser()
    config.read("C:\\Scripts\\Creds\\config.ini")

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


# Uses items/checkouts/ API endpoint to check item in again
def checkin_item(barcode, username, statgroup, sierra_api):
    url = (
        "items/checkouts/"
        + barcode
        + "?username="
        + username
        + "&statgroup="
        + statgroup
    )
    request = sierra_api.request("DELETE", url)
    request.raise_for_status()


# log items that were corrected to an existing Google Sheet
def appendToSheet(spreadSheetId, data):
    scopes = [
        "https://spreadsheets.google.com/feeds",
        "https://www.googleapis.com/auth/drive",
    ]
    creds = ServiceAccountCredentials.from_json_keyfile_name(
        "C:\\Scripts\\Creds\\GSheet updater creds.json", scopes
    )
    service = build("sheets", "v4", credentials=creds)
    sheet = service.spreadsheets()
    request = (
        service.spreadsheets()
        .values()
        .append(
            spreadsheetId=spreadSheetId,
            range="A1:Z1",
            valueInputOption="USER_ENTERED",
            body={"values": data},
        )
    )
    result = request.execute()


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
    # query to identify items that are simultaneously checked out and in transit, retrieve data needed for corrections and logging
    error_query = """\
            SELECT
              ip.barcode,
              u.name AS username,
              u.statistic_group_code_num AS checkin_stat_group_code,
              TO_TIMESTAMP(SPLIT_PART(v.field_content,': IN',1), 'Dy Mon DD YYYY  HH:MIAM')::VARCHAR AS checkin_time,
              so.name AS checkout_stat_group_name,
              i.checkout_statistic_group_code_num AS checkout_stat_group_code,
              o.checkout_gmt::VARCHAR AS checkout_time,
              v.field_content AS message,
              SPLIT_PART(SPLIT_PART(v.field_content,'from ',2),' to',1) AS origin_loc,
              SPLIT_PART(v.field_content,'to ',2) AS destination_loc,
              CASE
                WHEN h.id IS NOT NULL THEN true
	            ELSE FALSE
              END AS fulfilling_hold
  
             FROM sierra_view.item_record i
             JOIN sierra_view.checkout o
               ON i.id = o.item_record_id
             JOIN sierra_view.varfield v
               ON i.id = v.record_id
               AND v.varfield_type_code = 'm'
               AND v.field_content LIKE '%IN TRANSIT%'
             JOIN sierra_view.item_record_property ip
               ON i.id = ip.item_record_id
             JOIN sierra_view.statistic_group_myuser so
               ON i.checkout_statistic_group_code_num = so.code
             JOIN sierra_view.iii_user u
               ON SPLIT_PART(SPLIT_PART(v.field_content,'from ',2),' to',1) = u.name
             LEFT JOIN sierra_view.hold h
               ON i.id = h.record_id

             WHERE i.item_status_code IN ('t')
               --build in buffer to avoid catching items actively being checked out
               AND EXTRACT(EPOCH FROM (CURRENT_TIMESTAMP - TO_TIMESTAMP(SPLIT_PART(v.field_content,': IN',1), 'Dy Mon DD YYYY  HH:MIAM'))) > 120

             ORDER BY TO_TIMESTAMP(SPLIT_PART(v.field_content,': IN',1), 'Dy Mon DD YYYY  HH:MIAM')
            """

    item_errors = runquery(error_query)
    # log query results to preexisting Google sheet
    config = configparser.ConfigParser()
    config.read("C:\\Scripts\\Creds\\config.ini")
    appendToSheet(config["gsheet"]["correct_checkins"], item_errors)

    # initialize Sierra API
    sierra_api = init_api()
    # for each item in the error_query results, check it in again
    for rownum, row in enumerate(item_errors):
        checkin_item(str(row[0]), row[1], str(row[2]), sierra_api)


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
        email_subject = "correct checkin errors script error"
        email_message = (
            "Your script failed with the following error:\n\n" + traceback.format_exc()
        )

        send_email_error(email_subject, email_message, emailto)
        raise
