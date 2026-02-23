#!/usr/bin/env python3

# Run in py38

"""
Script to automate creation of minimal MARC records
and subsequent upload to Ingram for IPage's holdings feature

Author: Jeremy Goldstein
Contact Info: jgoldstein@minlib.net
"""

import psycopg2
import re
import pymarc
import configparser
import os
import pysftp
from datetime import date
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email.utils import formatdate
from email import encoders
import traceback


# function takes a sql query as a parameter, connects to a database and returns the results
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


# create file of MARC records based on data returned from run_query()
def marc_writer(query_data, marc_file):

    # create a mrc file using the filename passed to the function
    os.makedirs(os.path.dirname(marc_file), exist_ok=True)

    # open file in write binary mode
    with open(marc_file, "wb") as f:
        # iterate through each row of the file
        for rownum, row in enumerate(query_data):

            # declare PyMARC record object
            item_load = pymarc.Record(to_unicode=True, force_utf8=True)

            # define data fields in CSV file
            ocn = row[1]
            isbn = row[2].split("|")

            # Clean up OCLC numbers with regular expression
            ocn = re.sub("[^0-9]", "", ocn)

            # write data to field variables
            field_001 = pymarc.Field(tag="001", data=ocn)
            item_load.add_ordered_field(field_001)
            for i in isbn:
                if i == "":
                    break
                field_020 = pymarc.Field(
                    tag="020",
                    indicators=pymarc.Indicators(" ", " "),
                    subfields=[pymarc.Subfield(code="a", value=i)],
                )
                item_load.add_ordered_field(field_020)

            # write date to file
            f.write(item_load.as_marc())

    return marc_file


# function to sftp a specified file
def sftp_file(file, library):
    config = configparser.ConfigParser()
    config.read("C:\\Scripts\\Creds\\config.ini")

    # set connection option to disable check for host key
    cnopts = pysftp.CnOpts()
    cnopts.hostkeys = None

    # open sftp connection
    srv = pysftp.Connection(
        host=config["ingram"]["host"],
        username=config["ingram"]["user_" + library],
        password=config["ingram"]["pw_" + library],
        cnopts=cnopts,
    )
    # upload specified file to root directory
    srv.put(file)

    # close connection
    srv.close()


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


def main(library):
    try:
        # run holdings query for specified library
        query = open(library + "_ingram_holdings.sql", "r").read()
        query_results = run_query(query)

        # generate marc file based on those query results
        marc_file_name = (
            "/Scripts/Ingram Holdings/Temp_Files/"
            + library
            + "_holdings{}.mrc".format(date.today())
        )
        marc_file = marc_writer(query_results, marc_file_name)

        # sftp file to Ingram
        sftp_file(
            "C:\\Scripts\\Ingram Holdings\\Temp_Files\\"
            + library
            + "_holdings{}.mrc".format(date.today()),
            library,
        )

        # delete file once script is complete
        os.remove(marc_file)
    except Exception:
        # read config file with recipient list for email
        config_recipient = configparser.ConfigParser()
        config_recipient.read("C:\\Scripts\\Creds\\emails.ini")
        emailto = config_recipient["script_error"]["recipients"].split()

        # craft email subject and message containing error message details from traceback
        email_subject = "Ingram Holdings " + library + " script error"
        email_message = (
            "Your script failed with the following error:\n\n" + traceback.format_exc()
        )

        send_email_error(email_subject, email_message, emailto)
        raise


if __name__ == "__main__":
    main("blm")
    main("con")
    main("fpl")
    main("nor")
    main("som")
    main("win")
