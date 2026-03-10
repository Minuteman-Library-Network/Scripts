#!/usr/bin/env python3

#Run in py313
"""
Jeremy Goldstein
Minuteman Library Network

Generates monthly report of checkouts of LoT items at each Cambridge location
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

#connect to Sierra-db and store results of an sql query
def runquery(location):

    # import configuration file containing our connection string
    # config.ini looks like the following
    #[sql]
    #connection_string = dbname='iii' user='PUT_USERNAME_HERE' host='sierra-db.library-name.org' password='PUT_PASSWORD_HERE' port=1032

    config = configparser.ConfigParser()
    config.read('C:\\Scripts\\Creds\\config.ini')
    
    query = r"""
            SELECT
			t.transaction_gmt::DATE AS date,
			COUNT(t.id) FILTER(WHERE i.icode1 = '165') AS scat_165,
			COUNT(t.id) FILTER(WHERE i.icode1 = '166') AS scat_166,
			COUNT(t.id) FILTER(WHERE i.icode1 = '167') AS scat_167,
			COUNT(t.id) FILTER(WHERE i.icode1 = '168') AS scat_168,
			COUNT(t.id) FILTER(WHERE i.icode1 = '170') AS scat_170,
			COUNT(t.id) FILTER(WHERE i.icode1 = '171') AS scat_171,
			COUNT(t.id) FILTER(WHERE i.icode1 = '171') AS scat_172,
			COUNT(t.id) FILTER(WHERE i.icode1 = '180') AS scat_180,
			COUNT(t.id) FILTER(WHERE i.icode1 = '185') AS scat_185,
			COUNT(t.id) FILTER(WHERE i.icode1 = '186') AS scat_186,
			COUNT(t.id) AS total

			FROM
			sierra_view.circ_trans t
			JOIN
			sierra_view.item_record i
			ON
			t.item_record_id = i.id
			AND i.icode1 IN ('165','166','167','168','170','171','172','180','185','186')
			AND i.location_code = """ + location + """
            WHERE
			t.op_code = 'o'
			AND
			--t.transaction_gmt::DATE >= '2021-09-01' AND t.transaction_gmt::DATE < '2021-10-01'
			t.transaction_gmt::DATE BETWEEN CURRENT_DATE - INTERVAL '1 month' AND CURRENT_DATE
			GROUP BY 1
			ORDER BY 1
            """
      
    try:
	    # variable connection string should be defined in the imported config file
        conn = psycopg2.connect( config['sql']['connection_string'] )
    except:
        print("unable to connect to the database")
        clear_connection()
        return
        
    #Opening a session and querying the database for weekly new items
    cursor = conn.cursor()
    cursor.execute(query)
    #For now, just storing the data in a variable. We'll use it later.
    rows = cursor.fetchall()
    conn.close()
    
    return rows
'''
convert sql query results into formatted excel file
unlike similar functions in other scripts, runquery is called within excelWriter
'''
def excelWriter():
    #Name of Excel File
    excelfile =  '/Scripts/Cambridge Billed Items/Temp Files/CAMCheckoutsByScat{}.xlsx'.format(date.today())

    #Creating the Excel file for staff
    workbook = xlsxwriter.Workbook(excelfile,{'remove_timezone': True})
    worksheet = workbook.add_worksheet('cam')
    worksheet1 = workbook.add_worksheet('cam4')
    worksheet2 = workbook.add_worksheet('cam5')
    worksheet3 = workbook.add_worksheet('cam6')
    worksheet4 = workbook.add_worksheet('cam7')
    worksheet5 = workbook.add_worksheet('cam8')
    worksheet6 = workbook.add_worksheet('cam9')

    #Formatting our Excel worksheet
    worksheet.set_landscape()
    worksheet.hide_gridlines(0)

    #Formatting Cells
    eformat= workbook.add_format({'text_wrap': True, 'valign': 'bottom', 'font_name': 'Arial', 'font_size': '10', 'text_wrap': True, 'top': 1, 'bottom': 1})
    eformatlabel= workbook.add_format({'text_wrap': True, 'valign': 'bottom', 'bold': True, 'font_name': 'Arial', 'font_size': '10', 'text_wrap': True, 'top': 1, 'bottom': 1})
    dateformat = workbook.add_format({'num_format': 'mm/dd/yyyy', 'align': 'left', 'font_size': '10', 'font_name':'Arial', 'top': 1, 'bottom': 1, 'valign': 'bottom'})
   
    # Setting the column widths
    worksheet.set_column(0,0,12.43)
    worksheet.set_column(1,1,8.43)
    worksheet.set_column(2,2,8.43)
    worksheet.set_column(3,3,8.43)
    worksheet.set_column(4,4,8.43)
    worksheet.set_column(5,5,8.43)
    worksheet.set_column(6,6,8.43)
    worksheet.set_column(7,7,8.43)
    worksheet.set_column(8,8,8.43)
    worksheet.set_column(9,9,8.43)
    worksheet.set_column(10,10,8.43)
    worksheet.set_column(11,11,8.43)
    worksheet1.set_column(0,0,12.43)
    worksheet1.set_column(1,1,8.43)
    worksheet1.set_column(2,2,8.43)
    worksheet1.set_column(3,3,8.43)
    worksheet1.set_column(4,4,8.43)
    worksheet1.set_column(5,5,8.43)
    worksheet1.set_column(6,6,8.43)
    worksheet1.set_column(7,7,8.43)
    worksheet1.set_column(8,8,8.43)
    worksheet1.set_column(9,9,8.43)
    worksheet1.set_column(10,10,8.43)
    worksheet1.set_column(11,11,8.43)
    worksheet2.set_column(0,0,12.43)
    worksheet2.set_column(1,1,8.43)
    worksheet2.set_column(2,2,8.43)
    worksheet2.set_column(3,3,8.43)
    worksheet2.set_column(4,4,8.43)
    worksheet2.set_column(5,5,8.43)
    worksheet2.set_column(6,6,8.43)
    worksheet2.set_column(7,7,8.43)
    worksheet2.set_column(8,8,8.43)
    worksheet2.set_column(9,9,8.43)
    worksheet2.set_column(10,10,8.43)
    worksheet2.set_column(11,11,8.43)
    worksheet3.set_column(0,0,12.43)
    worksheet3.set_column(1,1,8.43)
    worksheet3.set_column(2,2,8.43)
    worksheet3.set_column(3,3,8.43)
    worksheet3.set_column(4,4,8.43)
    worksheet3.set_column(5,5,8.43)
    worksheet3.set_column(6,6,8.43)
    worksheet3.set_column(7,7,8.43)
    worksheet3.set_column(8,8,8.43)
    worksheet3.set_column(9,9,8.43)
    worksheet3.set_column(10,10,8.43)
    worksheet3.set_column(11,11,8.43)
    worksheet4.set_column(0,0,12.43)
    worksheet4.set_column(1,1,8.43)
    worksheet4.set_column(2,2,8.43)
    worksheet4.set_column(3,3,8.43)
    worksheet4.set_column(4,4,8.43)
    worksheet4.set_column(5,5,8.43)
    worksheet4.set_column(6,6,8.43)
    worksheet4.set_column(7,7,8.43)
    worksheet4.set_column(8,8,8.43)
    worksheet4.set_column(9,9,8.43)
    worksheet4.set_column(10,10,8.43)
    worksheet4.set_column(11,11,8.43)
    worksheet5.set_column(0,0,12.43)
    worksheet5.set_column(1,1,8.43)
    worksheet5.set_column(2,2,8.43)
    worksheet5.set_column(3,3,8.43)
    worksheet5.set_column(4,4,8.43)
    worksheet5.set_column(5,5,8.43)
    worksheet5.set_column(6,6,8.43)
    worksheet5.set_column(7,7,8.43)
    worksheet5.set_column(8,8,8.43)
    worksheet5.set_column(9,9,8.43)
    worksheet5.set_column(10,10,8.43)
    worksheet5.set_column(11,11,8.43)
    worksheet6.set_column(0,0,12.43)
    worksheet6.set_column(1,1,8.43)
    worksheet6.set_column(2,2,8.43)
    worksheet6.set_column(3,3,8.43)
    worksheet6.set_column(4,4,8.43)
    worksheet6.set_column(5,5,8.43)
    worksheet6.set_column(6,6,8.43)
    worksheet6.set_column(7,7,8.43)
    worksheet6.set_column(8,8,8.43)
    worksheet6.set_column(9,9,8.43)
    worksheet6.set_column(10,10,8.43)
    worksheet6.set_column(11,11,8.43)

    #Inserting a header
    worksheet.set_header('Cam Monthly Checkouts By Scat')

    # Adding column labels
    worksheet.write(0,0,'Date', eformatlabel)
    worksheet.write(0,1,'Scat 165', eformatlabel)
    worksheet.write(0,2,'Scat 166', eformatlabel)
    worksheet.write(0,3,'Scat 167', eformatlabel)
    worksheet.write(0,4,'Scat 168', eformatlabel)
    worksheet.write(0,5,'Scat 170', eformatlabel)
    worksheet.write(0,6,'Scat 171', eformatlabel)
    worksheet.write(0,7,'Scat 172', eformatlabel)
    worksheet.write(0,8,'Scat 180', eformatlabel)
    worksheet.write(0,9,'Scat 185', eformatlabel)
    worksheet.write(0,10,'Scat 186', eformatlabel)
    worksheet.write(0,11,'Total', eformatlabel)
    worksheet1.write(0,0,'Date', eformatlabel)
    worksheet1.write(0,1,'Scat 165', eformatlabel)
    worksheet1.write(0,2,'Scat 166', eformatlabel)
    worksheet1.write(0,3,'Scat 167', eformatlabel)
    worksheet1.write(0,4,'Scat 168', eformatlabel)
    worksheet1.write(0,5,'Scat 170', eformatlabel)
    worksheet1.write(0,6,'Scat 171', eformatlabel)
    worksheet1.write(0,7,'Scat 172', eformatlabel)
    worksheet1.write(0,8,'Scat 180', eformatlabel)
    worksheet1.write(0,9,'Scat 185', eformatlabel)
    worksheet1.write(0,10,'Scat 186', eformatlabel)
    worksheet1.write(0,11,'Total', eformatlabel)
    worksheet2.write(0,0,'Date', eformatlabel)
    worksheet2.write(0,1,'Scat 165', eformatlabel)
    worksheet2.write(0,2,'Scat 166', eformatlabel)
    worksheet2.write(0,3,'Scat 167', eformatlabel)
    worksheet2.write(0,4,'Scat 168', eformatlabel)
    worksheet2.write(0,5,'Scat 170', eformatlabel)
    worksheet2.write(0,6,'Scat 171', eformatlabel)
    worksheet2.write(0,7,'Scat 172', eformatlabel)
    worksheet2.write(0,8,'Scat 180', eformatlabel)
    worksheet2.write(0,9,'Scat 185', eformatlabel)
    worksheet2.write(0,10,'Scat 186', eformatlabel)
    worksheet2.write(0,11,'Total', eformatlabel)
    worksheet3.write(0,0,'Date', eformatlabel)
    worksheet3.write(0,1,'Scat 165', eformatlabel)
    worksheet3.write(0,2,'Scat 166', eformatlabel)
    worksheet3.write(0,3,'Scat 167', eformatlabel)
    worksheet3.write(0,4,'Scat 168', eformatlabel)
    worksheet3.write(0,5,'Scat 170', eformatlabel)
    worksheet3.write(0,6,'Scat 171', eformatlabel)
    worksheet3.write(0,7,'Scat 172', eformatlabel)
    worksheet3.write(0,8,'Scat 180', eformatlabel)
    worksheet3.write(0,9,'Scat 185', eformatlabel)
    worksheet3.write(0,10,'Scat 186', eformatlabel)
    worksheet3.write(0,11,'Total', eformatlabel)
    worksheet4.write(0,0,'Date', eformatlabel)
    worksheet4.write(0,1,'Scat 165', eformatlabel)
    worksheet4.write(0,2,'Scat 166', eformatlabel)
    worksheet4.write(0,3,'Scat 167', eformatlabel)
    worksheet4.write(0,4,'Scat 168', eformatlabel)
    worksheet4.write(0,5,'Scat 170', eformatlabel)
    worksheet4.write(0,6,'Scat 171', eformatlabel)
    worksheet4.write(0,7,'Scat 172', eformatlabel)
    worksheet4.write(0,8,'Scat 180', eformatlabel)
    worksheet4.write(0,9,'Scat 185', eformatlabel)
    worksheet4.write(0,10,'Scat 186', eformatlabel)
    worksheet4.write(0,11,'Total', eformatlabel)
    worksheet5.write(0,0,'Date', eformatlabel)
    worksheet5.write(0,1,'Scat 165', eformatlabel)
    worksheet5.write(0,2,'Scat 166', eformatlabel)
    worksheet5.write(0,3,'Scat 167', eformatlabel)
    worksheet5.write(0,4,'Scat 168', eformatlabel)
    worksheet5.write(0,5,'Scat 170', eformatlabel)
    worksheet5.write(0,6,'Scat 171', eformatlabel)
    worksheet5.write(0,7,'Scat 172', eformatlabel)
    worksheet5.write(0,8,'Scat 180', eformatlabel)
    worksheet5.write(0,9,'Scat 185', eformatlabel)
    worksheet5.write(0,10,'Scat 186', eformatlabel)
    worksheet5.write(0,11,'Total', eformatlabel)
    worksheet6.write(0,0,'Date', eformatlabel)
    worksheet6.write(0,1,'Scat 165', eformatlabel)
    worksheet6.write(0,2,'Scat 166', eformatlabel)
    worksheet6.write(0,3,'Scat 167', eformatlabel)
    worksheet6.write(0,4,'Scat 168', eformatlabel)
    worksheet6.write(0,5,'Scat 170', eformatlabel)
    worksheet6.write(0,6,'Scat 171', eformatlabel)
    worksheet6.write(0,7,'Scat 172', eformatlabel)
    worksheet6.write(0,8,'Scat 180', eformatlabel)
    worksheet6.write(0,9,'Scat 185', eformatlabel)
    worksheet6.write(0,10,'Scat 186', eformatlabel)
    worksheet6.write(0,11,'Total', eformatlabel)

    # Writing the report for staff to the Excel worksheet
    for rownum, row in enumerate(runquery("'camnn'")):
        worksheet.write(rownum+1,0,row[0], dateformat)
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
    for rownum, row in enumerate(runquery("'ca4nn'")):
        worksheet1.write(rownum+1,0,row[0], dateformat)
        worksheet1.write(rownum+1,1,row[1], eformat)
        worksheet1.write(rownum+1,2,row[2], eformat)
        worksheet1.write(rownum+1,3,row[3], eformat)
        worksheet1.write(rownum+1,4,row[4], eformat)
        worksheet1.write(rownum+1,5,row[5], eformat)
        worksheet1.write(rownum+1,6,row[6], eformat)
        worksheet1.write(rownum+1,7,row[7], eformat)
        worksheet1.write(rownum+1,8,row[8], eformat)
        worksheet1.write(rownum+1,9,row[9], eformat)
        worksheet1.write(rownum+1,10,row[10], eformat)
        worksheet1.write(rownum+1,11,row[11], eformat)
    for rownum, row in enumerate(runquery("'ca5nn'")):
        worksheet2.write(rownum+1,0,row[0], dateformat)
        worksheet2.write(rownum+1,1,row[1], eformat)
        worksheet2.write(rownum+1,2,row[2], eformat)
        worksheet2.write(rownum+1,3,row[3], eformat)
        worksheet2.write(rownum+1,4,row[4], eformat)
        worksheet2.write(rownum+1,5,row[5], eformat)
        worksheet2.write(rownum+1,6,row[6], eformat)
        worksheet2.write(rownum+1,7,row[7], eformat)
        worksheet2.write(rownum+1,8,row[8], eformat)
        worksheet2.write(rownum+1,9,row[9], eformat)
        worksheet2.write(rownum+1,10,row[10], eformat)
        worksheet2.write(rownum+1,11,row[11], eformat)
    for rownum, row in enumerate(runquery("'ca6nn'")):
        worksheet3.write(rownum+1,0,row[0], dateformat)
        worksheet3.write(rownum+1,1,row[1], eformat)
        worksheet3.write(rownum+1,2,row[2], eformat)
        worksheet3.write(rownum+1,3,row[3], eformat)
        worksheet3.write(rownum+1,4,row[4], eformat)
        worksheet3.write(rownum+1,5,row[5], eformat)
        worksheet3.write(rownum+1,6,row[6], eformat)
        worksheet3.write(rownum+1,7,row[7], eformat)
        worksheet3.write(rownum+1,8,row[8], eformat)
        worksheet3.write(rownum+1,9,row[9], eformat)
        worksheet3.write(rownum+1,10,row[10], eformat)
        worksheet3.write(rownum+1,11,row[11], eformat)
    for rownum, row in enumerate(runquery("'ca7nn'")):
        worksheet4.write(rownum+1,0,row[0], dateformat)
        worksheet4.write(rownum+1,1,row[1], eformat)
        worksheet4.write(rownum+1,2,row[2], eformat)
        worksheet4.write(rownum+1,3,row[3], eformat)
        worksheet4.write(rownum+1,4,row[4], eformat)
        worksheet4.write(rownum+1,5,row[5], eformat)
        worksheet4.write(rownum+1,6,row[6], eformat)
        worksheet4.write(rownum+1,7,row[7], eformat)
        worksheet4.write(rownum+1,8,row[8], eformat)
        worksheet4.write(rownum+1,9,row[9], eformat)
        worksheet4.write(rownum+1,10,row[10], eformat)
        worksheet4.write(rownum+1,11,row[11], eformat)
    for rownum, row in enumerate(runquery("'ca8nn'")):
        worksheet5.write(rownum+1,0,row[0], dateformat)
        worksheet5.write(rownum+1,1,row[1], eformat)
        worksheet5.write(rownum+1,2,row[2], eformat)
        worksheet5.write(rownum+1,3,row[3], eformat)
        worksheet5.write(rownum+1,4,row[4], eformat)
        worksheet5.write(rownum+1,5,row[5], eformat)
        worksheet5.write(rownum+1,6,row[6], eformat)
        worksheet5.write(rownum+1,7,row[7], eformat)
        worksheet5.write(rownum+1,8,row[8], eformat)
        worksheet5.write(rownum+1,9,row[9], eformat)
        worksheet5.write(rownum+1,10,row[10], eformat)
        worksheet5.write(rownum+1,11,row[11], eformat)
    for rownum, row in enumerate(runquery("'ca9nn'")):
        worksheet6.write(rownum+1,0,row[0], dateformat)
        worksheet6.write(rownum+1,1,row[1], eformat)
        worksheet6.write(rownum+1,2,row[2], eformat)
        worksheet6.write(rownum+1,3,row[3], eformat)
        worksheet6.write(rownum+1,4,row[4], eformat)
        worksheet6.write(rownum+1,5,row[5], eformat)
        worksheet6.write(rownum+1,6,row[6], eformat)
        worksheet6.write(rownum+1,7,row[7], eformat)
        worksheet6.write(rownum+1,8,row[8], eformat)
        worksheet6.write(rownum+1,9,row[9], eformat)
        worksheet6.write(rownum+1,10,row[10], eformat)
        worksheet6.write(rownum+1,11,row[11], eformat)
     
    workbook.close()
    return excelfile


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
    emailto = config_recipient["cambridge_checkouts_by_scat"]["recipients"].split()
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
	
    excel_file = excelWriter()

    # send email
    email_subject = "Cambridge Monthly Checkouts By Scat"
    email_message = """***This is an automated email***


    The cam monthly checkouts by scat report has been attached."""
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
        emailto = config_recipient["script_error"]["recipients"].split()

        # craft email subject and message containing error message details from traceback
        email_subject = "annual reports: record count script error"
        email_message = (
            "Your script failed with the following error:\n\n" + traceback.format_exc()
        )

        send_email_error(email_subject, email_message, emailto)
        raise
