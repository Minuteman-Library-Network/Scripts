#!/usr/bin/env python3

# run in py313

"""
Jeremy Goldstein
Minuteman Library Network

Script used to generate a monthly report of withdrawn items counts
as a crosstab report by Item Type and Location Code
Report is produced as an Excel file
that is then uploaded to our staff intranet site for distribution, via sftp.
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

#convert sql query results into formatted excel file
def excel_writer_pandas(query_results,headers,excel_file):

    # create dataframe from query results
    df = pd.DataFrame(query_results, columns = headers)
    # convert dataframe to pivot table, making sure to preserve the sort order from the query results for the column order
    df = df.pivot_table(index = ['location_code','location_name'], columns = ['itype'], values = 'item_count', sort = False)
    # with pivot table complete sort rows on first column
    df = df.sort_index(level=0)
    # replace null entries with 0's 
    df = df.fillna(0)
    # write to Excel file using xlsxwriter for formatting column widths and the header row
    writer = pd.ExcelWriter(excel_file, engine='xlsxwriter')
    df.to_excel(writer)
    workbook = writer.book
    worksheet = writer.sheets['Sheet1']
    worksheet.set_column('A:A', 13.33)
    worksheet.set_column('B:B', 32.22)
    cell_format = workbook.add_format({'bold': True})
    worksheet.set_row(0, None, cell_format)
    writer.close()


# upload report to SIC directory and optionally remove older files
def sftp_file(local_file, file_name):

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
    remote_path = "/reports/Network-Wide Reports/Withdrawn Count"
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

    
def main():

    query = r"""
            WITH code_list AS (
              SELECT DISTINCT
                l.code AS location_code,
                l.name AS location_name,
                it.code AS itype_code,
                it.name AS itype_name
              FROM sierra_view.location_myuser l
              CROSS JOIN sierra_view.itype_property_myuser it

              WHERE it.name != ''
              ORDER BY 1,3
            )

            SELECT
              i.location_code,
              cl.location_name AS location_name,
              cl.itype_code||' '||cl.itype_name AS itype,
              COALESCE(COUNT(cl.*),0) AS item_count

            FROM code_list cl
            JOIN sierra_view.item_record i
              ON cl.location_code = i.location_code
              AND cl.itype_code = i.itype_code_num

            WHERE i.location_code !~ ('^(int|ceb)')
              AND i.itype_code_num NOT IN ('240','241','242')
              AND i.item_status_code = 'w'
              AND i.last_status_update::DATE >= CURRENT_DATE - INTERVAL '1 month'
  
            GROUP BY 1,2,cl.itype_code,cl.itype_name
            ORDER BY cl.itype_code,1
            """
       
    query_results, headers = run_query(query)
    # Name of Excel File
    file_name = "MLNWithdrawnCount{}.xlsx".format((date.today().replace(day=1) - timedelta(1)).strftime("%b%Y"))
    excel_file =  "/Scripts/Withdrawn By Location/Temp Files/{}".format(file_name)
    excel_writer_pandas(query_results, headers, excel_file)
    sftp_file(excel_file, file_name)


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
        email_subject = "New Items By Location script error"
        email_message = (
            "Your script failed with the following error:\n\n" + traceback.format_exc()
        )

        send_email_error(email_subject, email_message, emailto)
        raise