#!/usr/bin/env python3

#Run in py313
"""
Jeremy Goldstein
Minuteman Library Network

Generates report on the turnover rate for a number of 
defined collections at the Woburn library.
Produces separate reports for adults/teens and youth
"""

import psycopg2
import xlsxwriter
import os
import configparser
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email.utils import formatdate
from email import encoders
from datetime import date
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
    # close database connection
    conn.close()
    # return variables containing query results and column headers
    return rows

# convert sql query results into formatted excel file
def excel_writer(query_results, excel_file):

    #Creating the Excel file for staff
    workbook = xlsxwriter.Workbook(excel_file)
    worksheet = workbook.add_worksheet()

    #Formatting our Excel worksheet
    worksheet.set_landscape()
    worksheet.hide_gridlines(0)

    # Formatting Cells
    eformat= workbook.add_format({'text_wrap': True, 'valign': 'top'})
    eformatlabel= workbook.add_format({'text_wrap': True, 'valign': 'top', 'bold': True})

    # Setting the column widths
    worksheet.set_column(0,0,20.57)
    worksheet.set_column(1,1,12.57)
    worksheet.set_column(2,2,10.43)
    worksheet.set_column(3,3,11.57)

    # Inserting a header
    worksheet.set_header('Woburn monthly turnover')

    # Adding column labels
    worksheet.write(0,0,'Category Name', eformatlabel)
    worksheet.write(0,1,'Total Owned', eformatlabel)
    worksheet.write(0,2,'Total Circs', eformatlabel)
    worksheet.write(0,3,'Turnover', eformatlabel)

    # Writing the report for staff to the Excel worksheet
    for rownum, row in enumerate(query_results):
        worksheet.write(rownum+1,0,row[0], eformat)
        worksheet.write(rownum+1,1,row[1], eformat)
        worksheet.write(rownum+1,2,row[2], eformat)
        worksheet.write(rownum+1,3,row[3], eformat)
    
    workbook.close()

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

def main():

    adult_query = r"""
                  SELECT 
                    "Category Name",
                    "Total Owned",
                    "Total Circs",
                    ROUND(("Total Circs" * 1.00) / "Total Owned",2) AS Turnover
                  FROM (
                    SELECT
                      CASE
	                    WHEN i.icode1 IN ('127','129','131') THEN 'Audiobooks'
	                    WHEN i.icode1 IN ('92') THEN 'Biographies'
	                    WHEN i.icode1 IN ('148','149') THEN 'DVDs'
	                    WHEN i.icode1 = '147' THEN 'DVDs NF'
	                    WHEN i.icode1 IN ('6', '116', '117') THEN 'Large Print Fiction'
	                    WHEN i.icode1 = '103' THEN 'Large Print Non-fiction'
	                    WHEN i.icode1 IN ('145','159') THEN 'Video Games'
	                    WHEN i.icode1 IN ('167','169','153') THEN 'Graphic'
	                    WHEN i.icode1 IN ('160','161','164') THEN 'Teen Fiction'
	                    WHEN i.icode1 = '166' THEN 'Teen Non-fiction'
	                    WHEN i.icode1 IN ('163','165') THEN 'Teen Languages'
	                    WHEN i.icode1 = '0' THEN 'Bonnie No Scat'
	                    WHEN i.icode1 = '101' THEN 'Periodicals'
	                    WHEN i.icode1 IN ('126','143','158') THEN 'Adult LOT'
	                    WHEN i.icode1 = '138' THEN 'Equipment'
	                    WHEN i.icode1 IN ('1','2','3','5','9','108','109','110','111','115','123','133','134','248') THEN 'Fiction'
	                    WHEN (i.icode1 BETWEEN '10' AND '100') OR i.icode1 IN ('102','106','124','139') THEN 'Non-fiction'
                      END AS "Category Name",
                      COUNT(DISTINCT i.id) AS "Total Owned",
                      COUNT(DISTINCT C.id) AS "Total Circs"

                    FROM sierra_view.item_record i
                    JOIN sierra_view.bib_record_item_record_link l
                      ON i.id = l.item_record_id
                    LEFT JOIN sierra_view.circ_trans C
                      ON i.id = C.item_record_id
                      AND C.op_code IN ('o','r')
                      AND C.transaction_gmt >= NOW()::DATE - INTERVAL '1 month'

                    WHERE i.location_code ~ '^wob'
                    GROUP BY 1
                  ) a
                  
                  WHERE "Category Name" IS NOT NULL

                  UNION

                  SELECT
                    'total' "AS Category NAME",
                    COUNT(DISTINCT i.id) AS "Total Owned",
                    COUNT(DISTINCT C.id) AS "Total Circs",
                    ROUND((COUNT(DISTINCT C.id) * 1.00) / COUNT(DISTINCT i.id),2) AS Turnover

                  FROM sierra_view.item_record i
                  LEFT JOIN sierra_view.circ_trans C
                    ON i.id = C.item_record_id
                    AND C.op_code IN ('o','r')
                    AND C.transaction_gmt >= NOW()::DATE - INTERVAL '1 month'

                  WHERE i.location_code ~ '^wob' 
                    AND (i.icode1  BETWEEN '10' AND '100'
                      OR i.icode1 IN (
                        '127','129','131',
                        '92',
                        '148','149',
                        '147',
                        '6', '116', '117',
                        '103',
                        '145','159',
                        '167','169',
                        '153','160','161','164','165',
                        '166',
                        '0',
                        '101',
                        '126','143','158',
                        '138',
                        '1','2','3','5','9','108','109','110','111','115','123','133','134','248',
                        '102','106','124','139'
                      )
                    )

                  ORDER BY 1
                  """
    youth_query = r"""
                  SELECT 
                    "Category Name",
                    "Total Owned",
                    "Total Circs",
                    ROUND(("Total Circs" * 1.00) / "Total Owned",2) AS Turnover
                  FROM (
                    SELECT
                      CASE
	                    WHEN i.icode1 IN ('227') THEN 'Wonderbooks and Vox Books'
	                    WHEN i.icode1 IN ('150') THEN 'Playaways'
	                    WHEN i.icode1 IN ('232', '233') THEN 'Big Books'
	                    WHEN i.icode1 IN ('220') THEN 'Biographies'
	                    WHEN i.icode1 IN ('200') THEN 'Board Books '
	                    WHEN i.icode1 IN ('236') THEN 'DVDs, fiction '
	                    WHEN i.icode1 IN ('235') THEN 'DVDs, nonfic'
	                    WHEN i.icode1 IN ('209', '194') THEN 'Early Readers ' 
	                    WHEN i.icode1 IN ('231') THEN 'Decodables'
	                    WHEN i.icode1 IN ('201', '202', '204', '205') THEN 'Fiction'
	                    WHEN i.icode1 IN ('190') THEN 'Graphic'
	                    WHEN id2reckey(l.bib_record_id) = 'b3829074' THEN 'iPads'
	                    WHEN i.icode1 IN ('125') THEN 'Library of Things'
	                    WHEN i.icode1 IN ('198') THEN 'Video Games'
	                    WHEN i.icode1 IN ('234') THEN 'Launchpads'
	                    WHEN i.icode1 IN ('192', '193', '195', '197') THEN 'Middle Readers'
	                    WHEN i.icode1 BETWEEN '210' AND '219' THEN 'Nonfiction'
	                    WHEN i.icode1 IN ('206', '196', '191') THEN 'Picture Books'
	                    WHEN i.icode1 IN ('208') THEN 'PTR'
	                    WHEN i.icode1 IN ('237', '238', '239') THEN 'World'
                      END AS "Category Name",
                      COUNT(DISTINCT i.id) AS "Total Owned",
                      COUNT(DISTINCT C.id) AS "Total Circs"

                    FROM sierra_view.item_record i
                    JOIN sierra_view.bib_record_item_record_link l
                      ON i.id = l.item_record_id
                    LEFT JOIN sierra_view.circ_trans C
                      ON i.id = C.item_record_id
                      AND C.op_code IN ('o','r')
                      AND C.transaction_gmt >= NOW()::DATE - INTERVAL '1 month'

                    WHERE i.location_code ~ '^wob'
                    GROUP BY 1
                  ) a
                  
                  WHERE "Category Name" IS NOT NULL

                  UNION

                  SELECT
                    'total' "AS Category NAME",
                    COUNT(DISTINCT i.id) AS "Total Owned",
                    COUNT(DISTINCT C.id) AS "Total Circs",
                    ROUND((COUNT(DISTINCT C.id) * 1.00) / COUNT(DISTINCT i.id),2) AS Turnover

                  FROM sierra_view.item_record i
                  LEFT JOIN sierra_view.circ_trans C
                    ON i.id = C.item_record_id
                    AND C.op_code IN ('o','r')
                    AND C.transaction_gmt >= NOW()::DATE - INTERVAL '1 month'
                  JOIN sierra_view.bib_record_item_record_link l
                    ON i.id = l.item_record_id

                  WHERE i.location_code ~ '^wob' 
                    AND (id2reckey(l.bib_record_id) = 'b3829074'
                      OR i.icode1 IN ('173','50','233','220','200','236','235','209', '194','201', '202', '204', '205','190','125','192', '193', '195', '197','210','211','212','213','214','215','216','217','218','219','206', '196','208','199','237', '238'))

                  ORDER BY 1
                  """
    adult_query_results = run_query(adult_query)
    youth_query_results = run_query (youth_query)

    # generate excel file from those query results
    adult_excel_file = "/Scripts/Woburn Monthly Turnover/Temp Files/wob monthly turnover adult and teen{}.xlsx".format(date.today())
    excel_writer(adult_query_results, adult_excel_file)
    youth_excel_file = "/Scripts/Woburn Monthly Turnover/Temp Files/wob monthly turnover youth{}.xlsx".format(date.today())
    excel_writer(youth_query_results, youth_excel_file)

    # send email
    email_subject = "Woburn Monthly Turnover Report"
    email_message = """***This is an automated email***


The Woburn Monthly Turnover Report has been attached."""
    # read config file with recipient list for email
    config_recipient = configparser.ConfigParser()
    config_recipient.read("C:\\Scripts\\Creds\\emails.ini")
    adult_recipient = config_recipient["woburn_turnover"]["adult_recipients"].split()
    youth_recipient = config_recipient["woburn_turnover"]["youth_recipients"].split()
    send_email(email_subject, email_message, adult_excel_file, adult_recipient)
    send_email(email_subject, email_message, youth_excel_file, youth_recipient)

    # delete local file
    os.remove(adult_excel_file)
    os.remove(youth_excel_file)


# run main function and send error email to admin of script encounters an error
if __name__ == "__main__":
    try:
        main()
    except Exception:
        # read config file with recipient list for email
        config_recipient = configparser.ConfigParser()
        config_recipient.read("C:\\Scripts\\Creds\\emails.ini")
        emailto = config_recipient["script_error_extended"]["recipients"].split()

        # craft email subject and message containing error message details from traceback
        email_subject = "Woburn Monthly Turnover script error"
        email_message = (
            "Your script failed with the following error:\n\n" + traceback.format_exc()
        )

        send_email_error(email_subject, email_message, emailto)
        raise