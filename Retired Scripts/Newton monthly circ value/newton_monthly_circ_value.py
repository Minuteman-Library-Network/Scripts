#!/usr/bin/env python3


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
excelfile =  'C:\\SQL Reports\\Newton monthly circ value\\NtnMonthlyCircValue{}.xlsx'.format(date.today())

# These are variables for the email that will be sent.
# Make sure to use your own library's email server (emaihost)
emailhost = ''
emailuser = ''
emailpass = ''
emailport = ''
emailsubject = 'Newton Monthly Circ Value'
emailmessage = '''***This is an automated email***


The e-mail Field Problem report has been attached.'''
# Enter your own email information
emailfrom= ''
# emailto can send to multiple addresses by separating emails with commas
emailto = ['']

#Connecting to Sierra PostgreSQL database
conn = psycopg2.connect("dbname='iii' user='mlnsql' host='sierra-db.minlib.net' port='1032' password='1234' sslmode='require'")

#Opening a session and querying the database for weekly new items
cursor = conn.cursor()
cursor.execute(open("newton_monthly_circ_value.sql","r").read())
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
eformat2= workbook.add_format({'num_format': 'mm/dd/yy hh:mm:ss'})


# Setting the column widths
worksheet.set_column(0,0,11.86)
worksheet.set_column(1,1,11.14)
worksheet.set_column(2,2,14.14)
worksheet.set_column(3,3,16.86)
worksheet.set_column(4,4,16.86)

#Inserting a header
worksheet.set_header('Newton Monthly Circ Value')

# Adding column labels
worksheet.write(0,0,'Value', eformatlabel)
worksheet.write(0,1,'Circ_Count', eformatlabel)
worksheet.write(0,2,'Value_Per_Circ', eformatlabel)
worksheet.write(0,3,'Start_Time', eformatlabel)
worksheet.write(0,4,'End_Time', eformatlabel)


# Writing the report for staff to the Excel worksheet
for rownum, row in enumerate(rows):
    worksheet.write(rownum+1,0,row[0], eformat)
    worksheet.write(rownum+1,1,row[1], eformat)
    worksheet.write(rownum+1,2,row[2], eformat)
    worksheet.write(rownum+1,3,row[3], eformat2)
    worksheet.write(rownum+1,4,row[4], eformat2)
    
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
