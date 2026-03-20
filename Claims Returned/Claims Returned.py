#!/usr/bin/env python3

# run in py313

"""
Jeremy Goldstein
Minuteman Library Network

Generates list of claimed returned items for each library
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


# convert sql query results into formatted excel file
def excel_writer(query_results, excel_file):

    # Creating the Excel file for staff
    workbook = xlsxwriter.Workbook(excel_file, {"remove_timezone": True})
    worksheet = workbook.add_worksheet()

    # Formatting our Excel worksheet
    worksheet.set_landscape()
    worksheet.hide_gridlines(0)

    # Formatting Cells
    eformat = workbook.add_format(
        {
            "text_wrap": True,
            "valign": "top",
            "align": "left",
            "font_size": "8",
            "font_name": "Arial",
        }
    )
    eformatlabel = workbook.add_format(
        {
            "text_wrap": True,
            "valign": "top",
            "bold": True,
            "font_size": "8",
            "font_name": "Arial",
        }
    )
    dateformat = workbook.add_format(
        {
            "num_format": "mm/dd/yyyy",
            "align": "left",
            "font_size": "8",
            "font_name": "Arial",
        }
    )

    # Setting the column widths
    worksheet.set_column(0, 0, 7.43)
    worksheet.set_column(1, 1, 12.45)
    worksheet.set_column(2, 2, 17)
    worksheet.set_column(3, 3, 15)
    worksheet.set_column(4, 4, 45)
    worksheet.set_column(5, 5, 50)
    worksheet.set_column(6, 6, 10)
    worksheet.set_column(7, 7, 10.6)

    # Inserting a header
    worksheet.set_header("Items Claimed Returned")

    # Adding column labels
    worksheet.write(0, 0, "Location", eformatlabel)
    worksheet.write(0, 1, "Barcode", eformatlabel)
    worksheet.write(0, 2, "Call Num", eformatlabel)
    worksheet.write(0, 3, "Volume", eformatlabel)
    worksheet.write(0, 4, "Title", eformatlabel)
    worksheet.write(0, 5, "Messages", eformatlabel)
    worksheet.write(0, 6, "Status", eformatlabel)
    worksheet.write(0, 7, "Updated Date", eformatlabel)

    # Writing the report for staff to the Excel worksheet
    for rownum, row in enumerate(query_results):
        worksheet.write(rownum + 1, 0, row[0], eformat)
        worksheet.write(rownum + 1, 1, row[1], eformat)
        worksheet.write(rownum + 1, 2, row[2], eformat)
        worksheet.write(rownum + 1, 3, row[3], eformat)
        worksheet.write(rownum + 1, 4, row[4], eformat)
        worksheet.write(rownum + 1, 5, row[5], eformat)
        worksheet.write(rownum + 1, 6, row[6], eformat)
        worksheet.write(rownum + 1, 7, row[7], dateformat)

    workbook.close()

    return excel_file


# upload report to SIC directory and optionally remove older files
def sftp_file(local_file, library):

    config = configparser.ConfigParser()
    config.read("C:\\Scripts\\Creds\\config.ini")

    cnopts = pysftp.CnOpts()

    srv = pysftp.Connection(
        host=config["sic"]["sic_host"],
        username=config["sic"]["sic_user"],
        password=config["sic"]["sic_pw"],
        cnopts=cnopts,
    )

    local_file = local_file

    srv.cwd("/reports/Library-Specific Reports/" + library + "/Claims Returned/")
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

        this_month = date.today().replace(day=1)

        query = r"""
	    SELECT
          DISTINCT i.location_code,
	      iprop.barcode,
	      TRIM(regexp_replace(iprop.call_number,'\|.',' ','g')),
	      v.field_content,
	      bprop.best_title,
	      vm.field_content,
	      i.item_status_code,
	      rm.record_last_updated_gmt::DATE
	    FROM sierra_view.item_record i
	    JOIN sierra_view.item_record_property iprop
          ON i.id = iprop.item_record_id
	    JOIN sierra_view.bib_record_item_record_link bilink
          ON i.id = bilink.item_record_id
	    JOIN sierra_view.bib_record_property bprop
          ON bprop.bib_record_id = bilink.bib_record_id
	    JOIN sierra_view.record_metadata rm
          ON i.id = rm.id
	    JOIN sierra_view.itype_property ip
          ON i.itype_code_num = ip.code_num
	    JOIN sierra_view.itype_property_name it
          ON ip.id = it.itype_property_id
	    LEFT JOIN sierra_view.varfield v
          ON i.id = v.record_id
          AND v.varfield_type_code = 'v'
	    LEFT JOIN sierra_view.varfield vm
          ON i.id = vm.record_id
          AND vm.varfield_type_code = 'x'
          AND vm.field_content LIKE '%Claimed%'
	
        WHERE i.item_status_code = 'z'
	      AND i.location_code ~ '^{}'
        ORDER BY 1,3
        """.format(
            libcode[0:2].lower()
        )

        query_results = run_query(query)

        # Name of Excel File
        excel_file = (
            "/Scripts/Claims Returned/Temp Files/"
            + libcode
            + "ClaimsReturned{}.xlsx".format(this_month.strftime("%b%Y"))
        )
        excel_writer(query_results, excel_file)
        sftp_file(
            "C:\\Scripts\\Claims Returned\\Temp Files\\"
            + libcode
            + "ClaimsReturned{}.xlsx".format(this_month.strftime("%b%Y")),
            library,
        )
    except:
        # read config file with recipient list for email
        config_recipient = configparser.ConfigParser()
        config_recipient.read("C:\\Scripts\\Creds\\emails.ini")
        emailto = config_recipient["script_error"]["recipients"].split()

        # craft email subject and message containing error message details from traceback
        email_subject = "claims returned " + library + " script error"
        email_message = (
            "Your script failed with the following error:\n\n" + traceback.format_exc()
        )

        send_email_error(email_subject, email_message, emailto)
        raise


if __name__ == "__main__":
    # run for each library within Minuteman
    main("Acton", "ACT")
    main("Arlington", "ARL")
    main("Ashland", "ASH")
    main("Bedford", "BED")
    main("Belmont", "BLM")
    main("Brookline", "BRK")
    main("Cambridge", "CAM")
    main("Concord", "CON")
    main("Dedham", "DDM")
    main("Dean", "DEA")
    main("Dover", "DOV")
    main("Framingham Public", "FPL")
    main("Framingham State", "FST")
    main("Franklin", "FRK")
    main("Holliston", "HOL")
    main("Lasell", "LAS")
    main("Lexington", "LEX")
    main("Lincoln", "LIN")
    main("Maynard", "MAY")
    main("Medfield", "MLD")
    main("Medford", "MED")
    main("Medway", "MWY")
    main("Millis", "MIL")
    main("Natick", "NAT")
    main("Needham", "NEE")
    main("Newton", "NTN")
    main("Norwood", "NOR")
    main("Olin", "OLN")
    main("Regis", "REG")
    main("Sherborn", "SHR")
    main("Somerville", "SOM")
    main("Stow", "STO")
    main("Sudbury", "SUD")
    main("Waltham", "WLM")
    main("Watertown", "WAT")
    main("Wayland", "WYL")
    main("Wellesley", "WEL")
    main("Weston", "WSN")
    main("Westwood", "WWD")
    main("Winchester", "WIN")
    main("Woburn", "WOB")
