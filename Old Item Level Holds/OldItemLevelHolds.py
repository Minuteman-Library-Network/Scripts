#!/usr/bin/env python3

# run in py313

"""
Jeremy Goldstein
Minuteman Library Network

Generates monthly list of old item holds for each library
Saves lists as Excel documents, which are upload to our intranet site for distribution to staff
"""

import psycopg2
import xlsxwriter
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
    # return variable containing query results
    return rows

#convert sql query results into formatted excel file
def excel_writer(query_results,excel_file):

    #Creating the Excel file for staff
    workbook = xlsxwriter.Workbook(excel_file,{'remove_timezone': True})
    worksheet = workbook.add_worksheet()


    #Formatting our Excel worksheet
    worksheet.set_landscape()
    worksheet.hide_gridlines(0)

    #Formatting Cells
    eformat= workbook.add_format({'text_wrap': True, 'valign': 'top'})
    eformatlabel= workbook.add_format({'text_wrap': True, 'valign': 'top', 'bold': True})

    # Setting the column widths
    worksheet.set_column(0,0,14.14)
    worksheet.set_column(1,1,13.43)
    worksheet.set_column(2,2,26.57)
    worksheet.set_column(3,3,14)
    worksheet.set_column(4,4,55.29)
    worksheet.set_column(5,5,12.86)
    worksheet.set_column(6,6,16)
    worksheet.set_column(7,7,11.3)

    #Inserting a header
    worksheet.set_header('Old Item Level Holds')

    # Adding column labels
    worksheet.write(0,0,'Item Number', eformatlabel)
    worksheet.write(0,1,'Location Code', eformatlabel)
    worksheet.write(0,2,'Call Number', eformatlabel)
    worksheet.write(0,3,'Volume', eformatlabel)
    worksheet.write(0,4,'Title', eformatlabel)
    worksheet.write(0,5,'Status', eformatlabel)
    worksheet.write(0,6,'Barcode', eformatlabel)
    worksheet.write(0,7,'IType', eformatlabel)

    # Writing the report for staff to the Excel worksheet
    for rownum, row in enumerate(query_results):
        worksheet.write(rownum+1,0,row[0], eformat)
        worksheet.write(rownum+1,1,row[1], eformat)
        worksheet.write(rownum+1,2,row[2], eformat)
        worksheet.write(rownum+1,3,row[3], eformat)
        worksheet.write(rownum+1,4,row[4], eformat)
        worksheet.write(rownum+1,5,row[5], eformat)
        worksheet.write(rownum+1,6,row[6], eformat)
        worksheet.write(rownum+1,7,row[7], eformat)
    
    workbook.close()
    
    return excel_file


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
    remote_path = "/reports/Library-Specific Reports/{}/Old Item Level Holds".format(library)
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
                  DISTINCT id2reckey(i.id)||'a',
                  i.location_code,
                  REGEXP_REPLACE(ip.call_number,'\|[a-z]',''),
                  v.field_content,
                  b.best_title,
                  s.name,
                  ip.barcode,
                  it.name
                FROM sierra_view.hold h
                JOIN sierra_view.item_record i
                  ON h.record_id = i.id
                  AND i.item_status_code != '!'
                  AND i.itype_code_num NOT IN ('5','6','221','222','223','224','242')
                LEFT JOIN sierra_view.checkout o
                  ON i.id = o.item_record_id
                JOIN sierra_view.item_record_property ip
                  ON i.id = ip.item_record_id
                JOIN sierra_view.bib_record_item_record_link l
                  ON i.id = l.item_record_id
                JOIN sierra_view.bib_record_property b
                  ON l.bib_record_id = b.bib_record_id
                JOIN sierra_view.varfield v
                  ON i.id = v.record_id
                  AND v.varfield_type_code = 'v'
                JOIN sierra_view.itype_property_myuser it
                  ON i.itype_code_num = it.code
                JOIN sierra_view.item_status_property_myuser s
                  ON i.item_status_code = s.code
                
                WHERE h.status = '0'
                  AND h.is_frozen = 'false'
                  AND CURRENT_DATE > (h.placed_gmt::DATE + h.delay_days)
                  AND h.placed_gmt::DATE <(CURRENT_DATE - INTERVAL '2 weeks')
                  AND o.id IS NULL
                  AND i.location_code ~ '^""" + libcode[0:2].lower() + """'
                ORDER BY 2,6,3,4
                """
        query_results = run_query(query)
        #Name of Excel File
        file_name = "{}OldItemLevelHolds{}.xlsx".format(libcode,date.today().replace(day=1).strftime("%b%Y"))
        excel_file =  "/Scripts/Old Item Level Holds/Temp Files/{}".format(file_name)
        excel_writer(query_results,excel_file)
        sftp_file(excel_file, file_name, library)
    except:
      # read config file with recipient list for email
      config_recipient = configparser.ConfigParser()
      config_recipient.read("C:\\Scripts\\Creds\\emails.ini")
      emailto = config_recipient["script_error"]["recipients"].split()

      # craft email subject and message containing error message details from traceback
      email_subject = "Old Item Level Holds " + library + " script error"
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