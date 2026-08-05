#!/usr/bin/env python3

#Run in py313
"""
Jeremy Goldstein
Minuteman Library Network

Generates report of lost in transit items for Sherborn
and emails results as an Excel file to designated staff
"""

import psycopg2
import xlsxwriter
import os
import smtplib
import configparser
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

#convert sql query results into formatted excel file
def excel_writer(query_results, excel_file):

    #Creating the Excel file for staff
    workbook = xlsxwriter.Workbook(excel_file,{'remove_timezone': True})
    worksheet = workbook.add_worksheet('Incoming')
    worksheet1 = workbook.add_worksheet('Outgoing')


    #Formatting our Excel worksheet
    worksheet.set_landscape()
    worksheet.hide_gridlines(0)
    worksheet1.set_landscape()
    worksheet1.hide_gridlines(0)

    #Formatting Cells
    eformat= workbook.add_format({'text_wrap': True, 'valign': 'top', 'align': 'left', 'font_size': '8', 'font_name':'Arial'})
    eformatlabel= workbook.add_format({'text_wrap': True, 'valign': 'top', 'bold': True})
    dateformat = workbook.add_format({'num_format': 'mm/dd/yyyy', 'align': 'left'})
   

    # Setting the column widths
    worksheet.set_column(0,0,7.71)
    worksheet.set_column(1,1,16.3)
    worksheet.set_column(2,2,17.43)
    worksheet.set_column(3,3,25.86)
    worksheet.set_column(4,4,42.71)
    worksheet.set_column(5,5,17)
    worksheet.set_column(6,6,6.57)
    worksheet.set_column(7,7,11.29)
    worksheet.set_column(8,8,7.71)
    worksheet1.set_column(0,0,7.71)
    worksheet1.set_column(1,1,16.3)
    worksheet1.set_column(2,2,17.43)
    worksheet1.set_column(3,3,25.86)
    worksheet1.set_column(4,4,42.71)
    worksheet1.set_column(5,5,17)
    worksheet1.set_column(6,6,6.57)
    worksheet1.set_column(7,7,11.29)
    worksheet1.set_column(8,8,7.71)
      

    # Adding column labels
    worksheet.write(0,0,'Location', eformatlabel)
    worksheet.write(0,1,'Item Barcode', eformatlabel)
    worksheet.write(0,2,'Call Number', eformatlabel)
    worksheet.write(0,3,'Author', eformatlabel)
    worksheet.write(0,4,'Title', eformatlabel)
    worksheet.write(0,5,'In Transit Date', eformatlabel)
    worksheet.write(0,6,'IType', eformatlabel)
    worksheet.write(0,7,'IType Name', eformatlabel)
    worksheet.write(0,8,'Origin', eformatlabel)
    worksheet1.write(0,0,'Location', eformatlabel)
    worksheet1.write(0,1,'Item Barcode', eformatlabel)
    worksheet1.write(0,2,'Call Number', eformatlabel)
    worksheet1.write(0,3,'Author', eformatlabel)
    worksheet1.write(0,4,'Title', eformatlabel)
    worksheet1.write(0,5,'In Transit Date', eformatlabel)
    worksheet1.write(0,6,'IType', eformatlabel)
    worksheet1.write(0,7,'IType Name', eformatlabel)
    worksheet1.write(0,8,'Destination', eformatlabel)
    
    # Writing the report for staff to the Excel worksheet
    row0 = 1
    row1 = 1
    
    for rownum, row in enumerate(query_results):
        if row[10] == 'shr':
            worksheet.write(row0,0,row[0], eformat)
            worksheet.write(row0,1,row[1], eformat)
            worksheet.write(row0,2,row[2], eformat)
            worksheet.write(row0,3,row[3], eformat)
            worksheet.write(row0,4,row[4], eformat)
            worksheet.write(row0,5,row[5], dateformat)
            worksheet.write(row0,6,row[6], eformat)
            worksheet.write(row0,7,row[7], eformat)
            worksheet.write(row0,8,row[9], eformat)
            row0 += 1
        elif row[9] == 'shr':
            worksheet1.write(row1,0,row[0], eformat)
            worksheet1.write(row1,1,row[1], eformat)
            worksheet1.write(row1,2,row[2], eformat)
            worksheet1.write(row1,3,row[3], eformat)
            worksheet1.write(row1,4,row[4], eformat)
            worksheet1.write(row1,5,row[5], dateformat)
            worksheet1.write(row1,6,row[6], eformat)
            worksheet1.write(row1,7,row[7], eformat)
            worksheet1.write(row1,8,row[10], eformat)
            row1 += 1
            
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


def main():
   
    query = r"""
    WITH transit AS (
      SELECT
      i.id AS id,
      SUBSTRING(SPLIT_PART(SPLIT_PART(v.field_content,'from ',2),' to',1)FROM 1 FOR 3) AS origin_loc,
      SUBSTRING(SPLIT_PART(v.field_content,'to ',2) FROM 1 FOR 3) AS destination_loc,
      to_timestamp(SPLIT_PART(v.field_content,': IN',1),'Dy Mon dd yyyy hh:miAM') AS placed_in_transit
    FROM sierra_view.item_record i
    JOIN sierra_view.varfield v
      ON i.id = v.record_id
      AND v.varfield_type_code = 'm'
      AND v.field_content LIKE '%IN TRANSIT%'
    WHERE
    i.item_status_code = 't')

    SELECT
		  DISTINCT i.location_code,
		  iprop.barcode,
		  TRIM(regexp_replace(iprop.call_number,'\|.',' ','g')) AS call_number,
		  bprop.best_author,
		  bprop.best_title,
		  t.placed_in_transit,
		  i.itype_code_num,
		  it.name,
		  record_metadata.record_last_updated_gmt,
		  t.origin_loc,
		  t.destination_loc
	  FROM transit t
	  JOIN sierra_view.item_record i
	    ON t.id = i.id
	  JOIN sierra_view.item_record_property iprop
	    ON i.id = iprop.item_record_id
	  JOIN sierra_view.bib_record_item_record_link bilink
	    ON bilink.item_record_id = i.id
	  JOIN sierra_view.bib_record_property bprop
	    ON bilink.bib_record_id = bprop.bib_record_id
	  JOIN sierra_view.record_metadata
	    ON record_metadata.id = i.id
	  JOIN sierra_view.itype_property ip
	    ON i.itype_code_num = ip.code_num
	  JOIN sierra_view.itype_property_name it
	    ON ip.id = it.itype_property_id 
	  WHERE i.item_status_code = 't'
		  AND ((current_date - i.last_status_update::date) > 14 OR i.last_status_update IS NULL)
		  AND (t.origin_loc = 'shr' OR t.destination_loc = 'shr')
	  ORDER BY 1,3
    """

    query_results = run_query(query)

    # generate excel file from those query results
    excel_file = "/Scripts/Sherborn Reports/Temp Files/SHRInTransitItems{}.xlsx".format(date.today().strftime("%b%Y"))
    excel_file = excel_writer(query_results, excel_file)
    

    email_subject = "SHR Weekly In Transit Items"
    email_message = """***This is an automated email***


    The Sherborn Weekly In Transit Items Report has been attached."""
	  # read config file with recipient list for email
    config_recipient = configparser.ConfigParser()
    config_recipient.read("C:\\Scripts\\Creds\\emails.ini")
    email_recipient = config_recipient["sherborn_in_transit"]["recipients"].split()  
    send_email(email_subject, email_message, excel_file, email_recipient)

    os.remove(excel_file)

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
        email_subject = "Sherborn In Transit Script Error"
        email_message = (
            "Your script failed with the following error:\n\n" + traceback.format_exc()
        )

        send_email_error(email_subject, email_message, emailto)
        raise
