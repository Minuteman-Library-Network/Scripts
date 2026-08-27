#!/usr/bin/env python3

# run in py313

"""
Jeremy Goldstein
Minuteman Library Network

Generates monthly list of new patrons for each library
Saves lists as Excel documents, which are upload to our intranet site for distribution to staff
"""

import psycopg2
import xlsxwriter
import os
import paramiko
import configparser
import sys
import time
from datetime import date, datetime, timedelta
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
    eformat= workbook.add_format({'text_wrap': True, 'valign': 'top', 'align': 'left', 'font_size': '8', 'font_name':'Arial'})
    eformatlabel= workbook.add_format({'text_wrap': True, 'valign': 'top', 'bold': True, 'font_size': '8', 'font_name':'Arial'})
    dateformat = workbook.add_format({'num_format': 'mm/dd/yyyy', 'align': 'left', 'font_size': '8', 'font_name':'Arial'})

    # Setting the column widths
    worksheet.set_column(0,0,6.3)
    worksheet.set_column(1,1,12.45)
    worksheet.set_column(2,2,5.15)
    worksheet.set_column(3,3,10.86)
    worksheet.set_column(4,4,8)
    worksheet.set_column(5,5,9)
    worksheet.set_column(6,6,20)
    worksheet.set_column(7,7,38)
    worksheet.set_column(8,8,22)

    #Inserting a header
    worksheet.set_header('New Patrons')

    # Adding column labels
    worksheet.write(0,0,'Home Library', eformatlabel)
    worksheet.write(0,1,'Barcode', eformatlabel)
    worksheet.write(0,2,'PType', eformatlabel)
    worksheet.write(0,3,'Creation Date', eformatlabel)
    worksheet.write(0,4,'MA Town', eformatlabel)
    worksheet.write(0,5,'MA Town Name', eformatlabel)    
    worksheet.write(0,6,'Name', eformatlabel)
    worksheet.write(0,7,'Address', eformatlabel)
    worksheet.write(0,8,'Email', eformatlabel)

    # Writing the report for staff to the Excel worksheet
    for rownum, row in enumerate(query_results):
        worksheet.write(rownum+1,0,row[0], eformat)
        worksheet.write(rownum+1,1,row[1], eformat)
        worksheet.write(rownum+1,2,row[2], eformat)
        worksheet.write(rownum+1,3,row[3], dateformat)
        worksheet.write(rownum+1,4,row[4], eformat)
        worksheet.write(rownum+1,5,row[5], eformat)
        worksheet.write(rownum+1,6,row[6], eformat)
        worksheet.write(rownum+1,7,row[7], eformat)
        worksheet.write(rownum+1,8,row[8], eformat)

    
    workbook.close()
    
    return excel_file


# upload report to SIC directory and optionally remove older files
def sftp_file(local_file, file_name, library):
    # read config file with Sierra login credentials
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
    remote_path = "/reports/Library-Specific Reports/{}/New Patrons".format(library)
    remote_path_to_file = remote_path + "/{}".format(file_name)
    sftp.put(local_file, remote_path_to_file)

    # remove old file
    cutoff_time = datetime.now() - timedelta(days=90)
    for entry in sftp.listdir_attr(remote_path):
        if not entry.st_mode & 0o40000:  # 0o40000 is S_IFDIR
            file_time = datetime.fromtimestamp(entry.st_mtime)
            if file_time < cutoff_time and entry.filename != "meta.json":
                full_path = remote_path + "/{}".format(entry.filename)
                sftp.remove(full_path)

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
                  pv.home_library_code,
                  pv.barcode,
                  pv.ptype_code,
                  rm.creation_date_gmt,
                  pv.pcode3,
                  ud.name,
                  prf.last_name || ', ' || prf.first_name || ' ' || prf.middle_name,
                  CONCAT(a.addr1,', ',a.city,', ',a.region,' ',a.postal_code),
                  (SELECT
                     e.field_content
                   FROM sierra_view.varfield_view e
                   WHERE pv.id = e.record_id
                     AND e.occ_num = 0
                     AND e.record_type_code = 'p'
                     AND e.varfield_type_code = 'z'
                  ),
                  pv.record_num AS PatronId
                FROM sierra_view.patron_view pv
                JOIN sierra_view.record_metadata rm
                  ON rm.id = pv.id
                  AND rm.creation_date_gmt >= date_trunc('month',CURRENT_DATE - INTERVAL '1 month')
                    AND rm.creation_date_gmt < date_trunc('month', CURRENT_DATE)
                JOIN sierra_view.patron_record pr
                  ON pr.record_id = pv.id
                JOIN sierra_view.user_defined_pcode3_myuser ud
                  ON pv.pcode3::text = ud.code
                JOIN sierra_view.patron_record_fullname prf
                  ON prf.patron_record_id = rm.id
                LEFT OUTER JOIN sierra_view.patron_record_address a
                  ON pv.id = a.patron_record_id
                  AND a.patron_record_address_type_id = 1
                WHERE pv.ptype_code <> 207
                  AND pv.home_library_code != 'none'
                  AND pv.home_library_code ~ '^""" + libcode[0:2].lower() + """'
                ORDER BY 1,2
                """
        query_results = run_query(query)
        # Name of Excel File
        file_name = "{}NewPatrons{}.xlsx".format(libcode, (date.today().replace(day=1) - timedelta(4)).strftime("%b%Y"))
        excel_file = "/Scripts/New Patrons/Temp Files/{}".format(file_name)
        excel_writer(query_results, excel_file)
        sftp_file(excel_file, file_name, library)

    except:
      # read config file with recipient list for email
      config_recipient = configparser.ConfigParser()
      config_recipient.read("C:\\Scripts\\Creds\\emails.ini")
      emailto = config_recipient["script_error"]["recipients"].split()

      # craft email subject and message containing error message details from traceback
      email_subject = "New Patrons " + library + " script error"
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