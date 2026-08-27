#!/usr/bin/env python3

# run in py313

"""
Jeremy Goldstein
Minuteman Library Network

Generates monthly list of new items for each library
Saves lists as HTML tables, which are upload to our intranet site for distribution to staff
"""

import psycopg2
import pandas as pd
import os
import paramiko
import configparser
import sys
import time
from datetime import date, timedelta
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email.utils import formatdate
from email import encoders
import traceback


# run sql query against Sierra database and return results
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

# convert sql query results into html table
def html_writer(query_results,html_file):
    # create dataframe from query results
    df = pd.DataFrame(query_results)
    df[2] = df.apply(lambda row: f'<a href="{row[3]}">{row[2]}</a>',axis=1)
    df = df.drop(columns=[3])
    html_table = df.to_html(index=False, header=False, render_links=True, escape=False)

    html = f'''<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<TITLE>New Items this Month</TITLE>
</head>
<BODY>

<CENTER>
<TABLE height="10%" width="75%">
  <TBODY>
  <TR align=middle bgColor=#33ccff>
    <TH vAlign=center>New Items this Month</TH>
  </TR>
  </TBODY>
</TABLE>

<TABLE height="10%" width="75%">
  <TBODY>
  <TR>
    <TH bgColor=#ccccc><FONT size=4>at this Library</FONT></TH>
  </TR>
  </TBODY>
</TABLE>
</CENTER>

<HR>
<P>

{html_table}

</BODY>
</html>'''

    with open(html_file, "w", encoding='utf-8') as f:
        f.write(html)

# upload report to SIC directory and optionally remove older files
def sftp_file(local_file, file_name, library):

    config = configparser.ConfigParser()
    config.read("C:\\Scripts\\Creds\\config.ini")

    # establish ssh client
    ssh = paramiko.SSHClient()
    ssh.set_missing_host_key_policy(paramiko.AutoAddPolicy())

    # connect to sftp server
    ssh.connect(
        hostname=config["sic"]["sic_host"],
        username=config["sic"]["sic_user"],
        password=config["sic"]["sic_pw"]
    )
    sftp = ssh.open_sftp()

    local_file = local_file
    remote_path = "/reports/Library-Specific Reports/{}/New Items".format(library)
    remote_path_to_file = remote_path + "/{}".format(file_name)
    sftp.put(local_file, remote_path_to_file)

    sftp.close()
    os.remove(local_file)


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

    
def main(library,libcode):

    try:
        query = r"""
            SELECT
              TRIM(REGEXP_REPLACE(ip.call_number,'\|.',' ','g')) as call_number,
              bp.best_author as author,
              bp.best_title as title,
              'https://catalog.minlib.net/Record/.'||rmb.record_type_code||rmb.record_num AS url
              
            FROM sierra_view.item_record i
            JOIN sierra_view.record_metadata rm
              ON i.id = rm.id
            JOIN sierra_view.item_record_property ip
              ON i.id = ip.item_record_id
            JOIN sierra_view.bib_record_item_record_link l
              ON i.id = l.item_record_id
            JOIN sierra_view.bib_record_property bp
              ON l.bib_record_id = bp.bib_record_id
            JOIN sierra_view.record_metadata rmb
              ON l.bib_record_id = rmb.id
            WHERE rm.creation_date_gmt >= DATE_TRUNC('month', CURRENT_DATE - INTERVAL '1 month')
              AND rm.creation_date_gmt < DATE_TRUNC('month', CURRENT_DATE)
              AND i.itype_code_num NOT IN ('241', '255', '242')
              AND i.item_message_code <> 'f'
              AND i.location_code ~ '^"""+libcode[0:2].lower()+"""'
              ORDER BY 1,2,3
            """
       
        query_results = run_query(query)
        # Name of Excel File
        file_name = "{}NewItems{}.htm".format(libcode, (date.today().replace(day=1) - timedelta(3)).strftime("%b%Y"))
        html_file = "/Scripts/New Items/Temp Files/{}".format(file_name)
        html_writer(query_results, html_file)
        sftp_file(html_file, file_name, library)

    except:
      # read config file with recipient list for email
      config_recipient = configparser.ConfigParser()
      config_recipient.read("C:\\Scripts\\Creds\\emails.ini")
      emailto = config_recipient["script_error"]["recipients"].split()

      # craft email subject and message containing error message details from traceback
      email_subject = "New Items HTML " + library + " script error"
      email_message = (
            "Your script failed with the following error:\n\n" + traceback.format_exc()
      )

      send_email_error(email_subject, email_message, emailto)
      raise
    


if __name__ == "__main__":
    # run for each library within Minuteman
    main('Acton','ACT')
    main('Arlington','ARL')
    main('Ashland','ASH')
    main('Bedford','BED')
    main('Belmont','BLM')
    main('Brookline','BRK')
    main('Cambridge','CAM')
    main('Concord','CON')
    main('Dedham','DDM')
    main('Dean','DEA')
    main('Dover','DOV')
    main('Framingham Public','FPL')
    main('Framingham State','FST')
    main('Franklin','FRK')
    main('Holliston','HOL')
    main('Lasell','LAS')
    main('Lexington','LEX')
    main('Lincoln','LIN')
    main('Maynard','MAY')
    main('Medfield','MLD')
    main('Medford','MED')
    main('Medway','MWY')
    main('Millis','MIL')
    main('Natick','NAT')
    main('Needham','NEE')
    main('Newton','NTN')
    main('Norwood','NOR')
    main('Olin','OLN')
    main('Regis','REG')
    main('Sherborn','SHR')
    main('Somerville','SOM')
    main('Stow','STO')
    main('Sudbury','SUD')
    main('Waltham','WLM')
    main('Watertown','WAT')
    main('Wayland','WYL')
    main('Wellesley','WEL')
    main('Weston','WSN')
    main('Westwood','WWD')
    main('Winchester','WIN')
    main('Woburn','WOB')