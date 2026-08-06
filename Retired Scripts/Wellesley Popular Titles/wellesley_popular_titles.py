#!/usr/bin/env python3

"""
Jeremy Goldstein
Minuteman Library Network
Generates Popular Titles at Wellesley report and emails results as an Excel file
"""

import psycopg2
import xlsxwriter
import smtplib
import os
import configparser
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email.utils import formatdate
from email import encoders
from datetime import date

#Name of Excel File
excelfile =  'C:\\SQL Reports\\Wellesley Popular Titles\\WelPopularTitles{}.xlsx'.format(date.today())

# These are variables for the email that will be sent.
# Make sure to use your own library's email server (emaihost)
emailhost = ''
emailuser = ''
emailpass = ''
emailport = ''
emailsubject = 'Wellesley Popular Titles'
emailmessage = '''***This is an automated email***


The Wellesley Popular Titles report has been attached.'''
# Enter your own email information
emailfrom= ''
# emailto can send to multiple addresses by separating emails with commas
emailto = ['']

config = configparser.ConfigParser()
config.read('C:\\SQL Reports\\creds\\app_SIC.ini')
      
try:
	# variable connection string should be defined in the imported config file
    conn = psycopg2.connect( config['db']['connection_string'] )
except:
    print("unable to connect to the database")
    clear_connection()
        
#Opening a session and querying the database for weekly new items
cursor = conn.cursor()
cursor.execute(open("wellesley_popular_titles.sql","r").read())
#For now, just storing the data in a variable. We'll use it later.
rows = cursor.fetchall()
conn.close()

#Creating the Excel file for staff
workbook = xlsxwriter.Workbook(excelfile)
worksheet = workbook.add_worksheet()


#Formatting our Excel worksheet
worksheet.set_landscape()
worksheet.hide_gridlines(0)

#Formatting Cells
eformat= workbook.add_format({'text_wrap': True, 'valign': 'top'})
eformatlabel= workbook.add_format({'text_wrap': True, 'valign': 'top', 'bold': True})
eformaturl= workbook.add_format({'text_wrap': True, 'valign': 'top', 'font_color': 'blue'})


# Setting the column widths
worksheet.set_column(0,0,17)
worksheet.set_column(1,1,57.43)
worksheet.set_column(2,2,34.43)
worksheet.set_column(3,3,10.43)
worksheet.set_column(4,4,51.14)
worksheet.set_column(5,5,6)
worksheet.set_column(6,6,22.29)
worksheet.set_column(7,7,14.57)
worksheet.set_column(8,8,18)
worksheet.set_column(9,9,12.43)
worksheet.set_column(10,10,16.29)

#Inserting a header
worksheet.set_header('Wellesley Popular Titles')

# Adding column labels
worksheet.write(0,0,'Category', eformatlabel)
worksheet.write(0,1,'Title', eformatlabel)
worksheet.write(0,2,'Author', eformatlabel)
worksheet.write(0,3,'Bnumber', eformatlabel)
worksheet.write(0,4,'URL', eformatlabel)
worksheet.write(0,5,'Rank', eformatlabel)
worksheet.write(0,6,'WEL Ptype Transactions', eformatlabel)
worksheet.write(0,7,'Checkout Total', eformatlabel)
worksheet.write(0,8,'WEL Checkout Total', eformatlabel)
worksheet.write(0,9,'Holds Placed', eformatlabel)
worksheet.write(0,10,'WEL Holds Placed', eformatlabel)


# Writing the report for staff to the Excel worksheet
for rownum, row in enumerate(rows):
    worksheet.write(rownum+1,0,row[0], eformat)
    worksheet.write(rownum+1,1,row[1], eformat)
    worksheet.write(rownum+1,2,row[2], eformat)
    worksheet.write(rownum+1,3,row[3], eformat)
    worksheet.write_url(rownum+1,4,row[4], eformaturl, string=row[4])
    worksheet.write(rownum+1,5,row[5], eformat)
    worksheet.write(rownum+1,6,row[6], eformat)
    worksheet.write(rownum+1,7,row[7], eformat)
    worksheet.write(rownum+1,8,row[8], eformat)
    worksheet.write(rownum+1,9,row[9], eformat)
    worksheet.write(rownum+1,10,row[10], eformat)
    
workbook.close()

#Creating the email message
msg = MIMEMultipart()
msg['From'] = emailfrom
if type(emailto) is list:
    msg['To'] = ', '.join(emailto)
else:
    msg['To'] = emailto
msg['Date'] = formatdate(localtime = True)
msg['Subject'] = emailsubject
msg.attach (MIMEText(emailmessage))
part = MIMEBase('application', "octet-stream")
part.set_payload(open(excelfile,"rb").read())
encoders.encode_base64(part)
part.add_header('Content-Disposition','attachment; filename=%s' % excelfile)
msg.attach(part)

#Sending the email message
smtp = smtplib.SMTP(emailhost, emailport)
#for Google connection
smtp.ehlo()
smtp.starttls()
smtp.login(emailuser, emailpass)
smtp.sendmail(emailfrom, emailto, msg.as_string())
smtp.quit()

os.remove(excelfile)
