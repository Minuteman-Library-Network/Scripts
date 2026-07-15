#!/usr/bin/env python3

# run in py313

"""
Jeremy Goldstein
Minuteman Library Network

Generates list of bills paid through ECommerce items for each library
Saves lists as Excel documents, which are upload to our intranet site for distribution to staff
"""

import psycopg2
import xlsxwriter
import os
import pysftp
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
    eformat= workbook.add_format({'text_wrap': True, 'valign': 'top', 'align': 'left'})
    eformatlabel= workbook.add_format({'text_wrap': True, 'valign': 'top', 'bold': True})
    dateformat = workbook.add_format({'num_format': 'mm/dd/yyyy', 'align': 'left'})
    currencyformat = workbook.add_format({'num_format': '$#,##0.00', 'align': 'left'})

    # Setting the column widths
    worksheet.set_column(0,0,12.29)
    worksheet.set_column(1,1,12.29)
    worksheet.set_column(2,2,12.29)
    worksheet.set_column(3,3,12.29)
    worksheet.set_column(4,4,17.43)
    worksheet.set_column(5,5,8.14)
    worksheet.set_column(6,6,10)
    worksheet.set_column(7,7,78.71)
    worksheet.set_column(8,8,9.71)
    worksheet.set_column(9,9,16)    
    
    #Inserting a header
    worksheet.set_header('ECommerce Payments')

    # Adding column labels
    worksheet.write(0,0,'Amount Paid', eformatlabel)
    worksheet.write(0,1,'Charge Amount', eformatlabel)
    worksheet.write(0,2,'Processing Fee', eformatlabel)
    worksheet.write(0,3,'Billing Fee', eformatlabel)
    worksheet.write(0,4,'Charge Type', eformatlabel)
    worksheet.write(0,5,'Owning Location', eformatlabel)
    worksheet.write(0,6,'Date Paid', eformatlabel)
    worksheet.write(0,7,'Title', eformatlabel)
    worksheet.write(0,8,'Stat Group', eformatlabel)
    worksheet.write(0,9,'Owning Location', eformatlabel)    

    # Writing the report for staff to the Excel worksheet
    for rownum, row in enumerate(query_results):
        worksheet.write(rownum+1,0,row[0], currencyformat)
        worksheet.write(rownum+1,1,row[1], currencyformat)
        worksheet.write(rownum+1,2,row[2], currencyformat)
        worksheet.write(rownum+1,3,row[3], currencyformat)
        worksheet.write(rownum+1,4,row[4], eformat)
        worksheet.write(rownum+1,5,row[5], eformat)
        worksheet.write(rownum+1,6,row[6], dateformat)
        worksheet.write(rownum+1,7,row[7], eformat)
        worksheet.write(rownum+1,8,row[8], eformat)
        worksheet.write(rownum+1,9,row[9], eformat)
     
    workbook.close()
    
    return excel_file



# upload report to SIC directory and optionally remove older files
def sftp_file(local_file, library):

    config = configparser.ConfigParser()
    config.read("C:\\Scripts\\Creds\\config.ini")

    cnopts = pysftp.CnOpts()
    cnopts.hostkeys = None

    srv = pysftp.Connection(
        host=config["sic"]["sic_host"],
        username=config["sic"]["sic_user"],
        password=config["sic"]["sic_pw"],
        cnopts=cnopts,
    )

    local_file = local_file

    srv.cwd("/reports/Library-Specific Reports/" + library + "/ECommerce/")
    srv.put(local_file)

    os.remove(local_file)


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

def main(library, libcode):
    try:
        query = r"""
          SELECT
            fines_paid.paid_now_amt::MONEY,
            fines_paid.item_charge_amt::MONEY,
            fines_paid.processing_fee_amt::MONEY,
            fines_paid.billing_fee_amt::MONEY,
            CASE
              WHEN fines_paid.charge_type_code = '1' THEN 'Manual Charge'
	          WHEN charge_type_code = '2' THEN 'Overdue' 
              WHEN charge_type_code = '3' THEN 'Replacement'
	          WHEN charge_type_code = '4' THEN 'Adjustment (OverdueX)'
              WHEN charge_type_code = '5' THEN 'Lost Book'
	          WHEN charge_type_code = '6' THEN 'Overdue Renewed'
              WHEN charge_type_code = '7' THEN 'Rental'
	          WHEN charge_type_code = '8' THEN 'Rental Adjustment'
              WHEN charge_type_code = '9' THEN 'Debit'
	          WHEN charge_type_code = 'a' THEN 'Notice'
              WHEN charge_type_code = 'b' THEN 'Credit Card'
	          WHEN charge_type_code = 'p' THEN 'Program'
	          ELSE 'OTHER'
            END,
            fines_paid.charge_location_code,
            fines_paid.paid_date_gmt::DATE,
            fines_paid.description,
            fines_paid.tty_num,
            SUBSTRING(charge_location_code, 1, 2)

          FROM sierra_view.fines_paid

          WHERE DATE_TRUNC('month', paid_date_gmt) = DATE_TRUNC('month', CURRENT_DATE - INTERVAL '1 month')
            AND tty_num::VARCHAR ~ '(992)|0|3|8|9$'
            AND payment_status_code NOT IN ('0','3')
            AND fines_paid.paid_now_amt > '0'
            AND fines_paid.charge_location_code ~ '^{}'
  
          ORDER BY 10,7
          """.format(libcode[0:2].lower())
        query_results = run_query(query)

        # To calculate last month's name
        last_month = date.today().replace(day=1) - timedelta(1)
        # Name of Excel File
        excel_file = "/Scripts/Fines Paid/Temp Files/" + libcode + "ECommerce{}.xlsx".format(last_month.strftime("%b%Y"))
        excel_writer(query_results, excel_file)
        sftp_file("C:\\Scripts\\Fines Paid\\Temp Files\\" + libcode + "ECommerce{}.xlsx".format(last_month.strftime("%b%Y")),library)
    except:
        # read config file with recipient list for email
        config_recipient = configparser.ConfigParser()
        config_recipient.read("C:\\Scripts\\Creds\\emails.ini")
        emailto = config_recipient["script_error"]["recipients"].split()

        # craft email subject and message containing error message details from traceback
        email_subject = "ecommerce fines paid " + library + " script error"
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
    main('Pine Manor','PMC')
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
