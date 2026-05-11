#!/usr/bin/env python3

#Run in py313
"""
Jeremy Goldstein
Minuteman Library Network

Generates on the turnover rate for a number of 
defined collections at the Concord/Free Public Library's Fowler Branch
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

    #Formatting Cells
    eformat = workbook.add_format({'text_wrap': True, 'valign': 'top'})
    eformatlabel = workbook.add_format({'text_wrap': True, 'valign': 'top', 'bold': True})

    # Setting the column widths
    worksheet.set_column(0,0,15.14)
    worksheet.set_column(1,1,12.57)
    worksheet.set_column(2,2,10.43)
    worksheet.set_column(3,3,11.57)

    #Inserting a header
    worksheet.set_header('Concord monthly turnover')

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
def send_email(subject, message, attachment):
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
    emailto = config_recipient["concord_turnover"]["recipients"].split()
    # plain text of email message
    emailmessage = message

    # Creating the email message
    msg = MIMEMultipart()
    msg["From"] = emailfrom
    if type(emailto) is list:
        msg["To"] = ", ".join(emailto)
    else:
        msg["To"] = emailto
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
    smtp.sendmail(emailfrom, emailto, msg.as_string())
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
    # query to identify bills from the past week for childrens items belonging to Cambridge
    query = """
    SELECT 
      "Category Name",
      "Total Owned",
      "Total Circs",
      ROUND(("Total Circs" * 1.00) / "Total Owned",2) AS Turnover
      FROM(
      SELECT
        CASE
	      WHEN i.icode1 = '122' AND i.itype_code_num IN ('36','37') THEN 'Adult Books on CD'
	      WHEN i.icode1 IN ('117','155') AND i.itype_code_num IN ('27','23') THEN 'Adult DVDs, feature'
	      WHEN i.icode1 IN ('117','155') AND i.itype_code_num = '28' THEN 'Adult DVDs, feature'
	      WHEN i.icode1 IN ('117','155') AND i.itype_code_num IN ('19','20') THEN 'Adult DVDs, TV series'
	      WHEN i.icode1 = '123' AND i.itype_code_num BETWEEN '0' AND '2' THEN 'Adult Graphic'
	      WHEN i.icode1 = '122' AND i.itype_code_num = '33' THEN 'Adult Music'
	      WHEN i.icode1 = '2' AND i.itype_code_num BETWEEN '0' AND '2' THEN 'Adult Mysteries'
	      WHEN i.icode1 = '4' AND i.itype_code_num = '0' THEN 'Adult Fantasy'
	      WHEN i.icode1 = '4' AND i.itype_code_num = '4' THEN 'Adult New Fantasy'
	      WHEN i.icode1 = '5' AND i.itype_code_num = '1' THEN 'Adult Paperback'
	      WHEN i.icode1 = '239' AND i.itype_code_num = '3' THEN 'Adult Reference'
	      WHEN i.icode1 = '101' AND i.itype_code_num = '10' THEN 'Adult Periodicals'
	      WHEN i.icode1 = '3' AND i.itype_code_num BETWEEN '0' AND '2' THEN 'Adult Science Fiction'
	      WHEN (i.icode1 = '119' AND i.itype_code_num = '0') OR i.itype_code_num = '5' THEN 'Adult Speed Read'
	      WHEN i.icode1 IN ('1','5', '240','241') AND i.itype_code_num IN ('0','1') THEN 'Adult Fiction'
	      WHEN ((i.icode1 BETWEEN '10' AND '100') OR i.icode1 IN ('102','125')) AND i.itype_code_num IN('0','1','2','12','100','101') THEN 'Adult Nonfiction'
	      WHEN i.icode1 IN ('220','232') AND i.itype_code_num IN ('150','151') THEN 'Children''s Biographies'
	      WHEN i.icode1 = '205' AND i.itype_code_num IN ('150','151') THEN 'Children''s Board Books'
	      WHEN i.icode1 = '110' AND i.itype_code_num = '173' THEN 'Children''s Books on CD'
	      WHEN i.icode1 = '110' AND i.itype_code_num = '180' THEN 'Children''s Playaways'
	      WHEN i.icode1 = '118' AND i.itype_code_num = '167' THEN 'Children''s DVDs, fiction'
	      WHEN i.icode1 = '208' AND i.itype_code_num IN ('150','151') THEN 'Children''s Early Chapter'
	      WHEN i.icode1 = '209' AND i.itype_code_num IN ('150','151') THEN 'Children''s Early Readers'
	      WHEN i.icode1 IN ('201','202','204','236') AND i.itype_code_num IN ('150','151') THEN 'Children''s Fiction'
	      WHEN i.icode1 = '124' AND i.itype_code_num IN ('150','151') THEN 'Children''s Graphic'
	      WHEN i.icode1 = '235' AND i.itype_code_num IN ('150','151') THEN 'Children''s Tales'
	      WHEN i.icode1 = '7' THEN 'Children''s Holiday'
	      WHEN i.icode1 = '248' THEN 'Children''s Kits'
	      WHEN i.icode1 = '238' AND i.itype_code_num = '154' THEN 'Children''s Reference'
	      WHEN i.icode1 = '221' AND i.itype_code_num = '158' THEN 'Children''s Magazine'
	      WHEN i.icode1 = '110' AND i.itype_code_num = '171' THEN 'Children''s Music'
	      WHEN ((i.icode1 BETWEEN '210' AND '219') OR (i.icode1 BETWEEN '222' AND '231')) AND i.itype_code_num IN ('150','151') THEN 'Children''s Nonfiction'
	      WHEN i.icode1 = '206' AND i.itype_code_num IN ('150','151') THEN 'Children''s Picture Books'
	      WHEN i.icode1 = '234' THEN 'Children''s VOX'
	      WHEN i.icode1 = '118' AND i.itype_code_num = '168' THEN 'Children''s DVDs, Nonfiction'
	      WHEN i.icode1 IN ('161','162','163','165','242') AND i.itype_code_num BETWEEN '100' AND '102' THEN 'Teen Fiction'
	      WHEN i.icode1 = '164' AND i.itype_code_num IN ('100','101') THEN 'Teen Graphic'
	      WHEN i.icode1 IN ('1','2','3','4','5','240','241') AND i.itype_code_num = '4' THEN 'Adult New Fiction'
	      WHEN ((i.icode1 BETWEEN '10' AND '100') OR i.icode1 = '102') AND i.itype_code_num = '4' THEN 'Adult New Nonfiction'
	      WHEN i.icode1 = '122' AND i.itype_code_num = '125' THEN 'Teen Books on CD'
	      WHEN i.icode1 = '122' AND i.itype_code_num = '130' THEN 'Teen Playaways'
	      WHEN i.icode1 = '101' AND i.itype_code_num = '107' THEN 'Teen Magazine'
	      WHEN i.itype_code_num = '2' THEN 'Adult Large Print'
	      WHEN i.icode1 = '138' THEN 'Equipment'
	      WHEN i.icode1 = '245' THEN 'Children''s Parent/Teacher'
	      WHEN i.icode1 = '170' THEN 'Tween Fiction'
	      WHEN i.icode1 = '171' THEN 'Tween Graphic'
	      WHEN i.icode1 = '109' AND i.itype_code_num = '0' THEN 'Living/Learning'
	      WHEN i.icode1 = '114' AND i.itype_code_num = '4' THEN 'Adult Biographies (New)'
	      WHEN i.icode1 = '114' AND i.itype_code_num = '0' THEN 'Adult Biographies'
	      ELSE i.icode1::VARCHAR||' / '||i.itype_code_num
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

      WHERE i.location_code ~ '^co2'

      GROUP BY 1) a

    WHERE "Category Name" IS NOT NULL

    UNION

    SELECT
      'total' "AS Category Name",
      COUNT(DISTINCT i.id) AS "Total Owned",
      COUNT(DISTINCT C.id) AS "Total Circs",
      ROUND((COUNT(DISTINCT C.id) * 1.00) / COUNT(DISTINCT i.id),2) AS Turnover

    FROM sierra_view.item_record i
    JOIN sierra_view.bib_record_item_record_link l
      ON i.id = l.item_record_id
    LEFT JOIN sierra_view.circ_trans c
      ON i.id = c.item_record_id
      AND c.op_code IN ('o','r')
      AND c.transaction_gmt >= NOW()::DATE - INTERVAL '1 month'

    WHERE i.location_code ~ '^co2'

    ORDER BY 1
    """
    
    query_results = run_query(query)

    # generate excel file from those query results
    excel_file = "/Scripts/Concord Monthly Turnover/Temp Files/co2 monthly turnover{}.xlsx".format(date.today())
    excel_writer(query_results, excel_file)

    # send email
    email_subject = "Concord/Fowler Monthly Turnover Report"
    email_message = """***This is an automated email***


The monthly Concord/Fowler Turnover Report has been attached."""
    send_email(email_subject, email_message, excel_file)

    # delete local file
    os.remove(excel_file)


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
        email_subject = "Concord Monthly Turnover script error"
        email_message = (
            "Your script failed with the following error:\n\n" + traceback.format_exc()
        )

        send_email_error(email_subject, email_message, emailto)
        raise