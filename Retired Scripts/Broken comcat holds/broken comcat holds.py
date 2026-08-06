#!/usr/bin/env python3

"""Create and email a list of new items

Author: Gem Stone-Logan
Contact Info: gem.stone-logan@mountainview.gov or gemstonelogan@gmail.com
"""

import psycopg2
import xlsxwriter
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email.utils import formatdate
from email import encoders
from datetime import date

#Name of Excel File
excelfile =  'C:\\SQL Reports\\Broken comcat holds\\BrokenComcatHolds{}.xlsx'.format(date.today())

# These are variables for the email that will be sent.
# Make sure to use your own library's email server (emaihost)
emailhost = ''
emailuser = ''
emailpass = ''
emailport = ''
emailsubject = 'Broken Comcat holds'
emailmessage = '''***This is an automated email***


The e-mail Field Problem report has been attached.'''
# Enter your own email information
emailfrom= ''
# emailto can send to multiple addresses by separating emails with commas
emailto = ['']

#Connecting to Sierra PostgreSQL database
conn = psycopg2.connect("dbname='' user='' host='' port='' password='' sslmode='require'")

#Opening a session and querying the database for weekly new items
cursor = conn.cursor()
cursor.execute(open("broken comcat holds.sql","r").read())
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
eformatdate= workbook.add_format({'num_format': 'mm/dd/yy'})
eformatlabel= workbook.add_format({'text_wrap': True, 'valign': 'top', 'bold': True})


# Setting the column widths
worksheet.set_column(0,0,23)
worksheet.set_column(1,1,12.29)
worksheet.set_column(2,2,11)
worksheet.set_column(3,3,19.14)
worksheet.set_column(4,4,8.43)
worksheet.set_column(5,5,10.14)
worksheet.set_column(6,6,11.71)
worksheet.set_column(7,7,8.43)
worksheet.set_column(8,8,20.29)

#Inserting a header
worksheet.set_header('Broken Comcat Holds')

# Adding column labels
worksheet.write(0,0,'Record_last_updated', eformatlabel)
worksheet.write(0,1,'record_num', eformatlabel)
worksheet.write(0,2,'pnumber', eformatlabel)
worksheet.write(0,3,'hold_placed', eformatlabel)
worksheet.write(0,4,'is_frozen', eformatlabel)
worksheet.write(0,5,'delay_days', eformatlabel)
worksheet.write(0,6,'expires', eformatlabel)
worksheet.write(0,7,'status', eformatlabel)
worksheet.write(0,8,'Pickup_location_code', eformatlabel)

# Writing the report for staff to the Excel worksheet
for rownum, row in enumerate(rows):
    worksheet.write(rownum+1,0,row[0], eformatdate)
    worksheet.write(rownum+1,1,row[1], eformat)
    worksheet.write(rownum+1,2,row[2], eformat)
    worksheet.write(rownum+1,3,row[3], eformatdate)
    worksheet.write(rownum+1,4,row[4], eformat)
    worksheet.write(rownum+1,5,row[5], eformat)
    worksheet.write(rownum+1,6,row[6], eformat)
    worksheet.write(rownum+1,7,row[7], eformat)
    worksheet.write(rownum+1,8,row[8], eformat)
    
    
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
