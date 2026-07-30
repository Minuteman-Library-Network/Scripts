#!/usr/bin/env python3

# Run in py313

"""
Identify bibs lacking 008 fields, which are required for populating filters in our catalog
Use Sierra API to fill in cursory form of field where needed

Author: Jeremy Goldstein
Contact Info: jgoldstein@minlib.net
"""

import requests
import json
import configparser
from base64 import b64encode
import psycopg2
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email.utils import formatdate
from email import encoders
import traceback

# function takes a sql query as a parameter, connects to a database and returns the results
def run_query(query):
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

def get_token():
    # config api    
    config = configparser.ConfigParser()
    # Must enter credentials information into api_info.ini file located in this same directory.  A template is provided
    config.read("C:\\Scripts\\Creds\\config.ini")
    base_url = config["api"]["base_url"]
    client_key = config["api"]["client_key"]
    client_secret = config["api"]["client_secret"]
    auth_string = b64encode((client_key + ':' + client_secret).encode('ascii')).decode('utf-8')
    header = {}
    header["authorization"] = 'Basic ' + auth_string
    header["Content-Type"] = 'application/x-www-form-urlencoded'
    body = {"grant_type": "client_credentials"}
    url = base_url + '/token'
    response = requests.post(url, data=json.dumps(body), headers=header)
    json_response = json.loads(response.text)
    token = json_response["access_token"]
    return token

def mod_bib(bib_id,language,token,s):
    config = configparser.ConfigParser()
    config.read("C:\\Scripts\\Creds\\config.ini")
    bibpatch = {"varFields": [{"fieldTag": "y","marcTag": "008","content": "||||||s                      000 | {} d".format(language)}]} 
    url = config["api"]["base_url"] + "/bibs/" + bib_id
    header = {"Authorization": "Bearer " + token, "Content-Type": "application/json;charset=UTF-8"}

    request = s.put(url, data=json.dumps(bibpatch), headers = header)

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
	
    config = configparser.ConfigParser()
    config.read("C:\\Scripts\\Creds\\config.ini")
    #open API session using sierra_ils_utils library
    
    query = r"""
    SELECT
        rm.record_num,
        b.language_code

    FROM sierra_view.bib_record b
    LEFT JOIN sierra_view.control_field f
        ON
        b.id = f.record_id AND f.control_num = 8
    JOIN sierra_view.record_metadata rm
        ON
        b.id = rm.id

    WHERE f.id IS NULL
    AND b.language_code !~ '(^eng|^zxx|^und)'
    AND b.language_code !=''
    AND b.is_suppressed = FALSE
    AND b.bcode3 NOT IN ('c', 'n', 'o', 'q', 'r', 'z')

    ORDER BY 1
    """
    
    
    query_results = run_query(query)
    
    
    #start up a requests session
    token = get_token()
    s = requests.Session()
    #for each row in query results call mod_bib
    for rownum, row in enumerate(query_results):
        mod_bib(str(row[0]),row[1],token,s)

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
        email_subject = "Missing 008 script error"
        email_message = (
            "Your script failed with the following error:\n\n" + traceback.format_exc()
        )

        send_email_error(email_subject, email_message, emailto)
        raise

