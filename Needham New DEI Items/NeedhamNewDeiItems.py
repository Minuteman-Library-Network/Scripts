#!/usr/bin/env python3

# Run in py313

"""
Create and email a monthly report of new items
that contain the text DEI in a note field.
Additionally upload a copy of that report to intranet site via sftp.

Author: Jeremy Goldstein
Contact Info: jgoldstein@minlib.net
"""

import psycopg2
import xlsxwriter
import smtplib
import os
import pysftp
import sys
import configparser
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email.utils import formatdate
from email import encoders
from datetime import date
import time
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
    workbook = xlsxwriter.Workbook(excel_file)
    worksheet = workbook.add_worksheet()


    #Formatting our Excel worksheet
    worksheet.set_landscape()
    worksheet.hide_gridlines(0)

    #Formatting Cells
    eformat= workbook.add_format({'text_wrap': True, 'valign': 'top', 'align': 'left', 'border': 1})
    eformatlabel= workbook.add_format({'text_wrap': True, 'valign': 'bottom', 'bold': True, 'fg_color': '#D9D9D9', 'border': 1})
    eformattotal= workbook.add_format({'text_wrap': True, 'valign': 'top', 'align': 'left', 'bold': True, 'border': 1})


    # Setting the column widths
    worksheet.set_column(0,0,26)
    worksheet.set_column(1,1,8.14)
    worksheet.set_column(2,2,8.14)
    worksheet.set_column(3,3,14.29)
    worksheet.set_column(4,4,8.14)
    worksheet.set_column(5,5,8.14)
    worksheet.set_column(6,6,10.71)
    worksheet.set_column(7,7,9.71)
    worksheet.set_column(8,8,10.14)
    worksheet.set_column(9,9,8.14)
    worksheet.set_column(10,10,12)
    worksheet.set_column(11,11,8.14)
    worksheet.set_column(12,12,9.29)
    worksheet.set_column(13,13,10.57)

    # Inserting a header
    worksheet.set_header('Needham New DEI Items')

    # Adding column labels
    worksheet.write(0,0,'Collection', eformatlabel)
    worksheet.write(0,1,'Asian', eformatlabel)
    worksheet.write(0,2,'Black', eformatlabel)
    worksheet.write(0,3,'Disabilities & Neurodiversity', eformatlabel)
    worksheet.write(0,4,'Equity & Social Issues', eformatlabel)
    worksheet.write(0,5,'Hispanic & Latino', eformatlabel)
    worksheet.write(0,6,'Indigenous', eformatlabel)
    worksheet.write(0,7,'LGBTQIA+ & Gender Studies', eformatlabel)
    worksheet.write(0,8,'Mental & Emotional Health', eformatlabel)
    worksheet.write(0,9,'Middle Eastern & North African', eformatlabel)
    worksheet.write(0,10,'Multicultural', eformatlabel)
    worksheet.write(0,11,'Religion', eformatlabel)
    worksheet.write(0,12,'Substance Abuse & Addiction', eformatlabel)
    worksheet.write(0,13,'Total Items Added', eformatlabel)


    # Writing the report for staff to the Excel worksheet
    for rownum, row in enumerate(query_results):
        if row[0] == 'TOTAL':
            worksheet.write(rownum+1,0,row[0], eformattotal)
            worksheet.write(rownum+1,1,row[1], eformattotal)
            worksheet.write(rownum+1,2,row[2], eformattotal)
            worksheet.write(rownum+1,3,row[3], eformattotal)
            worksheet.write(rownum+1,4,row[4], eformattotal)
            worksheet.write(rownum+1,5,row[5], eformattotal)
            worksheet.write(rownum+1,6,row[6], eformattotal)
            worksheet.write(rownum+1,7,row[7], eformattotal)
            worksheet.write(rownum+1,8,row[8], eformattotal)
            worksheet.write(rownum+1,9,row[9], eformattotal)
            worksheet.write(rownum+1,10,row[10], eformattotal)
            worksheet.write(rownum+1,11,row[11], eformattotal)
            worksheet.write(rownum+1,12,row[12], eformattotal)
            worksheet.write(rownum+1,13,row[13], eformattotal)
	
        else:
            worksheet.write(rownum+1,0,row[0], eformat)
            worksheet.write(rownum+1,1,row[1], eformat)
            worksheet.write(rownum+1,2,row[2], eformat)
            worksheet.write(rownum+1,3,row[3], eformat)
            worksheet.write(rownum+1,4,row[4], eformat)
            worksheet.write(rownum+1,5,row[5], eformat)
            worksheet.write(rownum+1,6,row[6], eformat)
            worksheet.write(rownum+1,7,row[7], eformat)
            worksheet.write(rownum+1,8,row[8], eformat)
            worksheet.write(rownum+1,9,row[9], eformat)
            worksheet.write(rownum+1,10,row[10], eformat)
            worksheet.write(rownum+1,11,row[11], eformat)
            worksheet.write(rownum+1,12,row[12], eformat)
            worksheet.write(rownum+1,13,row[13], eformattotal)
    
    workbook.close()

    return excel_file	

# function takes a file as a parameter and attaches that file to an outgoing email
def send_email(subject, message, attachment, recipient):
    # read config file with credentials for email account
    config = configparser.ConfigParser()
    config.read("C:\\Scripts\\Creds\\config.ini")
    # read config file with recipient list for email
    config_recipient = configparser.ConfigParser()
    config_recipient.read("C:\\Scripts\\Creds\\emails.ini")

    # These are variables for the email that will be sent, taken from .ini files referenced above
    emailhost = config["email"]["host"]
    emailuser = config["email"]["user"]
    emailpass = config["email"]["pw"]
    emailport = config["email"]["port"]
    emailfrom = config["email"]["sender"]
    # plain text of email message
    emailmessage = message

    # Creating the email message
    msg = MIMEMultipart()
    msg["From"] = emailfrom
    if type(recipient) is list:
        msg["To"] = ", ".join(recipient)
    else:
        msg["To"] = recipient
    msg["Date"] = formatdate(localtime=True)
    msg["Subject"] = subject
    msg.attach(MIMEText(emailmessage))
    part = MIMEBase("application", "octet-stream")
    part.set_payload(open(attachment, "rb").read())
    encoders.encode_base64(part)
    part.add_header(
        "Content-Disposition", "attachment; filename=%s" % attachment.rsplit("/", 1)[-1]
    )
    msg.attach(part)

    # Sending the email message
    smtp = smtplib.SMTP(emailhost, emailport)
    smtp.ehlo()
    smtp.starttls()
    smtp.login(emailuser, emailpass)
    smtp.sendmail(emailfrom, recipient, msg.as_string())
    smtp.quit()


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

    srv.cwd(
        "/reports/Library-Specific Reports/"
        + library
        + "/Custom/"
    )
    srv.put(local_file)

    #remove old file

    for fname in srv.listdir_attr():
        fullpath = "/reports/Library-Specific Reports/"+library+"/Custom/{}".format(fname.filename)
        #time tracked in seconds, st_mtime is time last modified
        name = str(fname.filename)
        if (name != 'meta.json') and  ((time.time() - fname.st_mtime) // (24 * 3600) >= 365):
            srv.remove(fullpath)

    srv.close()
    os.remove(local_file)

def main():

	# query to identify patron records with incorrect owed_amt fields
    query = r"""
		SELECT
          *
		FROM (
		  SELECT
		  CASE
			WHEN i.icode1 = '92' THEN 'A BIOGRAPHY'
			WHEN i.icode1 BETWEEN '10' AND '100' AND ip.call_number_norm ~ '^(?!express)\w+ \d' THEN 'A [LANGUAGE] NONFICTION'
			WHEN i.icode1 = '1' AND ip.call_number_norm ~ '^(?!express)\w+ fiction' THEN 'A [LANGUAGE] FICTION'
			WHEN i.icode1 BETWEEN '195' AND '237' AND ip.call_number_norm ~ '^j world' THEN 'J WORLD/[LANGUAGE]'
			WHEN i.icode1 BETWEEN '10' AND '100' THEN 'A NONFICTION'
			WHEN i.icode1 = '117' THEN 'A GENEALOGY'
			WHEN i.icode1 = '115' THEN 'A REF'
			WHEN i.icode1 = '124' THEN 'A ELL'
			WHEN i.icode1 = '1' THEN 'A FICTION'
			WHEN i.icode1 = '160' THEN 'A GRAPHIC'
			WHEN i.icode1 = '2' THEN 'A MYSTERY'
			WHEN i.icode1 = '3' THEN 'A SCIENCE FICTION'
			WHEN i.icode1 = '5' THEN 'A PAPERBACK'
			WHEN i.icode1 = '6' THEN 'A LARGE PRINT FICTION'
			WHEN i.icode1 = '103' THEN 'A LARGE PRINT NONFICTION'
			WHEN i.icode1 IN ('130','131') THEN 'A AUDIOBOOKS'
			WHEN i.icode1 = '129' THEN 'A MUSIC CDs'
			WHEN i.icode1 = '132' THEN 'A ELL AUDIOBOOKS/CDs'
			WHEN i.icode1 IN ('134','135') THEN 'A PLAYAWAYS'
			WHEN i.icode1 = '136' THEN 'A TABLET/eREADER'
			WHEN i.icode1 IN ('141','149','150') THEN 'A EQUIP'
			WHEN i.icode1 IN ('140','143','144','146') THEN 'A DVDS/BLU-RAYS'
			WHEN i.icode1 = '137' THEN 'A VIDEO GAMES'
			WHEN i.icode1 = '161' THEN 'Y FICTION'
			WHEN i.icode1 = '166' THEN 'Y GRAPHIC'
			WHEN i.icode1 = '190' THEN 'Y AUDIOBOOKS'
			WHEN i.icode1 = '192' THEN 'Y PLAYAWAYS'
			WHEN i.icode1 = '201' THEN 'J FICTION'
			WHEN i.icode1 = '205' THEN 'J ILLUSTRATED FICTION'
			WHEN i.icode1 = '196' THEN 'J READ-ALONG'
			WHEN i.icode1 = '198' THEN 'J MYSTERY'
			WHEN i.icode1 = '200' THEN 'J EASY CHAPTER'
			WHEN i.icode1 = '206' THEN 'J PICTURE BOOKS'
			WHEN i.icode1 = '207' THEN 'J BOARD BOOKS'
			WHEN i.icode1 = '208' THEN 'J PAPERBACK PIC'
			WHEN i.icode1 = '209' THEN 'J EASY READER'
			WHEN i.icode1 = '197' THEN 'J GRAPHIC NOVEL'
			WHEN i.icode1 BETWEEN '210' AND '219' THEN 'J NONFICTION'
			WHEN i.icode1 = '220' THEN 'J BIOGRAPHY'
			WHEN i.icode1 = '199' THEN 'J PARENTS SHELF'
			WHEN i.icode1 = '195' THEN 'J PHONICS KIT'
			WHEN i.icode1 = '234' THEN 'J DVD'
			WHEN i.icode1 = '229' THEN 'J COMPACT DISC'
			WHEN i.icode1 = '232' THEN 'J BOOK ON CD'
			WHEN i.icode1 = '235' THEN 'J PLAYAWAY'
			WHEN i.icode1 = '236' THEN 'J PLAYAWAY VIDEO'
			WHEN i.icode1 = '237' THEN 'J TABLET/EREADER'
			ELSE 'UNKNOWN'
		  END AS "collection",
		  COUNT(i.id) FILTER(WHERE LOWER(TRIM(REGEXP_REPLACE(v.field_content,'dei:\s?','','i'))) ~ 'asia') AS "Asian",
		  COUNT(i.id) FILTER(WHERE LOWER(TRIM(REGEXP_REPLACE(v.field_content,'dei:\s?','','i'))) ~ 'black') AS "Black",
		  COUNT(i.id) FILTER(WHERE LOWER(TRIM(REGEXP_REPLACE(v.field_content,'dei:\s?','','i'))) ~ '(disab)|(neuro)') AS "Disabilities & Neurodiversity",
		  COUNT(i.id) FILTER(WHERE LOWER(TRIM(REGEXP_REPLACE(v.field_content,'dei:\s?','','i'))) ~ '(equi)|(social)') AS "Equity & Social Issues",
		  COUNT(i.id) FILTER(WHERE LOWER(TRIM(REGEXP_REPLACE(v.field_content,'dei:\s?','','i'))) ~ '(spani)|(latin)') AS "Hispanic & Latino",
		  COUNT(i.id) FILTER(WHERE LOWER(TRIM(REGEXP_REPLACE(v.field_content,'dei:\s?','','i'))) ~ 'indig') AS "Indigenous",
		  COUNT(i.id) FILTER(WHERE LOWER(TRIM(REGEXP_REPLACE(v.field_content,'dei:\s?','','i'))) ~ '(lgbt)|(gender)') AS "LGBTQIA+ & Gender Studies",
		  COUNT(i.id) FILTER(WHERE LOWER(TRIM(REGEXP_REPLACE(v.field_content,'dei:\s?','','i'))) ~ '(mental)|(emotion)') AS "Mental & Emotional Health",
		  COUNT(i.id) FILTER(WHERE LOWER(TRIM(REGEXP_REPLACE(v.field_content,'dei:\s?','','i'))) ~ '(middle)|(north africa)') AS "Middle Eastern & North African",
		  COUNT(i.id) FILTER(WHERE LOWER(TRIM(REGEXP_REPLACE(v.field_content,'dei:\s?','','i'))) ~ 'multicult') AS "Multicultural",
		  COUNT(i.id) FILTER(WHERE LOWER(TRIM(REGEXP_REPLACE(v.field_content,'dei:\s?','','i'))) ~ 'religio') AS "Religion",
		  COUNT(i.id) FILTER(WHERE LOWER(TRIM(REGEXP_REPLACE(v.field_content,'dei:\s?','','i'))) ~ '(substance)|(addict)') AS "Substance Abuse & Addiction",
		  COUNT(i.id) AS total_items_added
		FROM sierra_view.item_record i
		JOIN sierra_view.record_metadata rm
		  ON i.id = rm.id
		JOIN sierra_view.varfield v
		  ON i.id = v.record_id
          AND v.varfield_type_code = 'x'
		JOIN sierra_view.item_record_property ip
		  ON i.id = ip.item_record_id

		WHERE rm.creation_date_gmt::DATE >= CURRENT_DATE - INTERVAL '1 month'
          AND i.location_code ~ '^nee'
          AND v.field_content LIKE 'DEI:%'

		GROUP BY 1

		UNION

		SELECT
		  'TOTAL' AS "Collection",
		  COUNT(i.id) FILTER(WHERE LOWER(TRIM(REGEXP_REPLACE(v.field_content,'dei:\s?','','i'))) ~ 'asia') AS "Asian",
		  COUNT(i.id) FILTER(WHERE LOWER(TRIM(REGEXP_REPLACE(v.field_content,'dei:\s?','','i'))) ~ 'black') AS "Black",
		  COUNT(i.id) FILTER(WHERE LOWER(TRIM(REGEXP_REPLACE(v.field_content,'dei:\s?','','i'))) ~ '(disab)|(neuro)') AS "Disabilities & Neurodiversity",
		  COUNT(i.id) FILTER(WHERE LOWER(TRIM(REGEXP_REPLACE(v.field_content,'dei:\s?','','i'))) ~ '(equi)|(social)') AS "Equity & Social Issues",
		  COUNT(i.id) FILTER(WHERE LOWER(TRIM(REGEXP_REPLACE(v.field_content,'dei:\s?','','i'))) ~ '(spani)|(latin)') AS "Hispanic & Latino",
		  COUNT(i.id) FILTER(WHERE LOWER(TRIM(REGEXP_REPLACE(v.field_content,'dei:\s?','','i'))) ~ 'indig') AS "Indigenous",
		  COUNT(i.id) FILTER(WHERE LOWER(TRIM(REGEXP_REPLACE(v.field_content,'dei:\s?','','i'))) ~ '(lgbt)|(gender)') AS "LGBTQIA+ & Gender Studies",
		  COUNT(i.id) FILTER(WHERE LOWER(TRIM(REGEXP_REPLACE(v.field_content,'dei:\s?','','i'))) ~ '(mental)|(emotion)') AS "Mental & Emotional Health",
		  COUNT(i.id) FILTER(WHERE LOWER(TRIM(REGEXP_REPLACE(v.field_content,'dei:\s?','','i'))) ~ '(middle)|(north africa)') AS "Middle Eastern & North African",
		  COUNT(i.id) FILTER(WHERE LOWER(TRIM(REGEXP_REPLACE(v.field_content,'dei:\s?','','i'))) ~ 'multicult') AS "Multicultural",
		  COUNT(i.id) FILTER(WHERE LOWER(TRIM(REGEXP_REPLACE(v.field_content,'dei:\s?','','i'))) ~ 'religio') AS "Religion",
		  COUNT(i.id) FILTER(WHERE LOWER(TRIM(REGEXP_REPLACE(v.field_content,'dei:\s?','','i'))) ~ '(substance)|(addict)') AS "Substance Abuse & Addiction",
		  COUNT(i.id) AS total_items_added
		
        FROM sierra_view.item_record i
		JOIN sierra_view.record_metadata rm
		  ON i.id = rm.id
		JOIN sierra_view.varfield v
		  ON i.id = v.record_id
          AND v.varfield_type_code = 'x'
		JOIN sierra_view.item_record_property ip
		  ON i.id = ip.item_record_id

		WHERE rm.creation_date_gmt::DATE >= CURRENT_DATE - INTERVAL '1 month'
          AND i.location_code ~ '^nee'
          AND v.field_content LIKE 'DEI:%'

		GROUP BY 1
		)a

		ORDER BY CASE WHEN a.Collection = 'TOTAL' THEN 2 ELSE 1 END,1
        """

    query_results = run_query(query)
    #Name of Excel File
    excel_file =  "/Scripts/Needham New DEI Items/Temp Files/Needham New DEI Items {}.xlsx".format(date.today())
    local_file = excel_writer(query_results, excel_file)

	# send email with attached file
	# read config file with recipient list for email
    config_recipient = configparser.ConfigParser()
    config_recipient.read("C:\\Scripts\\Creds\\emails.ini")
    email_recipient = config_recipient["needham_new_dei_items"]["recipients"].split()  
    email_subject = "Needham New DEI Items"
    email_message = """***This is an automated email***
    
    
    The Needham New DEI Items report has been attached."""
    send_email(email_subject, email_message, local_file, email_recipient)

    sftp_file(
            "C:\\Scripts\\Needham New DEI Items\\Temp Files\\Needham New DEI Items {}.xlsx".format(date.today()),
            "Needham",
        )

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
        email_subject = "Needham New DEI Items script error"
        email_message = (
            "Your script failed with the following error:\n\n" + traceback.format_exc()
        )

        send_email_error(email_subject, email_message, emailto)
        raise
