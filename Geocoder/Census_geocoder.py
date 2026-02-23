#!/usr/bin/env python3

# Run in geocoder

"""
Script to update census fields in patron records
after converting address to geoids using the Census bureau's batch geocoding API

Author: Jeremy Goldstein
Contact Info: jgoldstein@minlib.net

Takes roughly a bit over 4 hours to run
Due to time required to update patron records via the Sierra API
"""

import requests
import json
import os
import configparser
import psycopg2
import pandas as pd
import csv
import censusgeocode
from datetime import datetime
from datetime import timedelta
from datetime import date
from base64 import b64encode
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email.utils import formatdate
from email import encoders
import traceback


# function to generate access token for use with Sierra API
def get_token():
    # config api
    config = configparser.ConfigParser()
    config.read("C:\\Scripts\\Creds\\config.ini")
    base_url = config["api"]["base_url"]
    client_key = config["api"]["client_key"]
    client_secret = config["api"]["client_secret"]
    auth_string = b64encode((client_key + ":" + client_secret).encode("ascii")).decode(
        "utf-8"
    )
    header = {}
    header["authorization"] = "Basic " + auth_string
    header["Content-Type"] = "application/x-www-form-urlencoded"
    body = {"grant_type": "client_credentials"}
    url = base_url + "/token"
    response = requests.post(url, data=json.dumps(body), headers=header, verify=False)
    json_response = json.loads(response.text)
    token = json_response["access_token"]
    return token


# function to write/overwrite a census field within a patron record
def mod_patron(patronid, state, county, tract, block, token, s):
    config = configparser.ConfigParser()
    config.read("C:\\Scripts\\Creds\\config.ini")
    url = config["api"]["base_url"] + "/patrons/" + patronid
    header = {
        "Authorization": "Bearer " + token,
        "Content-Type": "application/json;charset=UTF-8",
    }
    payload = {
        "varFields": [
            {
                "fieldTag": "k",
                "content": "|s"
                + state
                + "|c"
                + county
                + "|t"
                + tract
                + "|b"
                + block
                + "|d"
                + format(date.today()),
            }
        ]
    }
    request = s.put(url, data=json.dumps(payload), headers=header)


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


# Function to write the results of a sql query to a specified csv file
def csv_writer(query_results, csv_file):

    with open(csv_file, "w", encoding="utf-8", newline="") as tempFile:
        myFile = csv.writer(tempFile, delimiter=",")
        myFile.writerows(query_results)
    tempFile.close()

    return csv_file


# Function takes a csv file containing address information and
# gather geoids from the census burearu's API, combining the two into a new output csv file
def geocode(csv_file):
    start_time = datetime.now()
    cg = censusgeocode.CensusGeocode(
        benchmark="Public_AR_Current", vintage="Current_Current"
    )

    print("sending batch file to process")
    result = cg.addressbatch(csv_file)

    print("---response from geocode %s seconds ---" % (datetime.now() - start_time))

    print("building data frame")
    df = pd.DataFrame(result, columns=result[0].keys())
    print("outputting to csv")
    path = "C:\Scripts\Geocoder\Temp Files"
    output_file = os.path.join(path, "output.csv")
    df.to_csv(output_file, mode="a", header=True)
    return output_file


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

    # tracking time it takes script to run as it can be a number of hours and may be useful to monitor
    start_time = datetime.now()
    print("querying Sierra database")
    query = """
    /*
    retrieves 10,000 patron record ids and addresses
    lacking a census field or due to be checked for potential address changes
    */
    SELECT
      DISTINCT a.id,
        a.street,
        a.city,
        a.region,
        a.zip 
    FROM (
      SELECT
        p.patron_record_id AS id,
        COALESCE(p.addr1,'') AS street,
        COALESCE(CASE 
          WHEN p.city IS NOT NULL THEN REGEXP_REPLACE(TRIM(p.city),'\s[a-zA-Z]{2}\s?(\d{5,})?\s?\-?\s?(\d{4,})?$','')
          WHEN p.city IS NULL AND p.addr1 ~ '\y([A-Za-z]+),?\s[A-Za-z]{2}(?:\s?\d{5,})?$' THEN (REGEXP_MATCH(TRIM(p.addr1), '\y([A-Za-z]+),?\s[A-Za-z]{2}(?:\s?\d{5,})?$'))[1]
        END,'') AS city,
        COALESCE(CASE
	      WHEN p.region = '' AND (LOWER(p.city) ~ '\sma$' OR pr.pcode3 BETWEEN '1' AND '200') THEN 'MA'
          WHEN p.region IS NULL AND p.city IS NULL AND TRIM(p.addr1) ~ '^.*\s(ma|Ma|mA|MA)(\s\d{5,})|$' THEN 'MA' 
          ELSE REGEXP_REPLACE(p.region,'\d|\-|\s|\\\\','','g')
	     END,'') AS region,
        COALESCE(CASE 
	      WHEN p.postal_code IS NULL AND TRIM(p.region) ~ '(MA)?\s?\d{5}' THEN SUBSTRING(TRIM(p.region),'\d{5,9}\s?\-?\s?\d{0,4}')
          WHEN p.postal_code IS NULL AND p.region = '' AND p.city ~ '(MA)?\s?\d{5}' THEN SUBSTRING(TRIM(p.city),'\d{5,9}\s?\-?\s?\d{0,4}$')
          WHEN p.postal_code IS NULL AND p.region IS NULL AND p.city IS NULL AND p.addr1 ~ '(MA)?\s?\d{5}' THEN SUBSTRING(TRIM(p.addr1),'\d{5,9}\s?\-?\s?\d{0,4}$')
          ELSE SUBSTRING(p.postal_code,'^\d{5}')
	     END,'') AS zip,
        s.content,
        rm.creation_date_gmt, 
        pr.ptype_code 
      FROM sierra_view.patron_record_address p 
      JOIN sierra_view.record_metadata rm 
        ON p.patron_record_id = rm.id
	    AND p.patron_record_address_type_id = '1' 
      LEFT JOIN sierra_view.subfield s 
        ON p.patron_record_id = s.record_id
	    AND s.field_type_code = 'k'
	    AND s.tag = 'd' 
      JOIN sierra_view.patron_record pr 
        ON p.patron_record_id = pr.id 

      WHERE LOWER(p.addr1) !~ '^p\.?\s?o'
        AND (((s.content IS NULL OR s.content !~ '^\d{4}\-\d{2}\-\d{2}$') AND pr.ptype_code NOT IN ('43','199','204','205','206','207','254'))
        OR TO_DATE(SUBSTRING(REGEXP_REPLACE(s.content,'[^0-9\-]','','g'),1,10),'YYYY-MM-DD') < rm.record_last_updated_gmt::DATE)
        /*
        Used when conducting a full update of patron records
        pr.ptype_code NOT IN ('43','199','204','205','206','207','254')
        AND pr.ptype_code IN ('1','3','4')
        */
      ORDER BY CASE
        WHEN s.content IS NULL THEN 1
	    ELSE 2
	    END, s.content 
      LIMIT 10000
    ) a
    """
    query_results = run_query(query)
    print("---sql query completed %s seconds ---" % (datetime.now() - start_time))

    csv_file = "/Scripts/Geocoder/Temp Files/patrons_to_geocode{}.csv".format(
        date.today()
    )
    file_to_send = csv_writer(query_results, csv_file)
    print("---csv generated %s seconds ---" % (datetime.now() - start_time))
    print("calling geocode function")
    output_file = geocode(file_to_send)
    # delete file_to_send once new output_file has been created
    os.remove(file_to_send)

    print("---geocode data retrieved %s seconds ---" % (datetime.now() - start_time))
    print("writing to Sierra patrons")
    with open(output_file, encoding="utf-8", newline="") as csv_file_temp:
        reader = csv.DictReader(csv_file_temp)

        # initialize Sierra API
        # track api token expiration time to know when it must be refreshed
        expiration_time = datetime.now() + timedelta(seconds=3600)
        token = get_token()
        s = requests.Session()

        # loop though csv list and write data to Sierra for each row
        for row in reader:
            # displaying patron ids, mostly as an indicator that the script is still running
            print(row["id"])
            # refresh API token if present one has expired
            if datetime.now() >= expiration_time:
                print("refreshing token")
                expiration_time = datetime.now() + timedelta(seconds=3600)
                token = get_token()
            mod_patron(
                row["id"],
                row["statefp"],
                row["countyfp"],
                row["tract"],
                row["block"],
                token,
                s,
            )

    print("---time to complete %s seconds ---" % (datetime.now() - start_time))

    # delete csv file
    os.remove(output_file)


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
        email_subject = "census geocoder script error"
        email_message = (
            "Your script failed with the following error:\n\n" + traceback.format_exc()
        )

        send_email_error(email_subject, email_message, emailto)
        raise
