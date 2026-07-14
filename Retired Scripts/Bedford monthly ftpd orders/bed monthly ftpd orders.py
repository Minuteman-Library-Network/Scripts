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

excelfile =  'bed monthly ftpd orders{}.xlsx'.format(date.today())

# These are variables for the email that will be sent.
# Make sure to use your own library's email server (emaihost)
emailhost = ''
emailuser = ''
emailpass = ''
emailport = ''
emailsubject = 'BED Monthly ftpd orders'
emailmessage = '''***This is an automated email***


The Bedford Monthly FTPd Orders Report has been attached.'''
# Enter your own email information
emailfrom= ''
# emailto can send to multiple addresses by separating emails with commas
emailto = ['']

#Connecting to Sierra PostgreSQL database
conn = psycopg2.connect("dbname='' user='' host='' port='1032' password='' sslmode='require'")

query = """\
        SELECT
          o.vendor_record_code AS vendor_code,
          COUNT(DISTINCT o.id) AS total_records,
          SUM(cmf.copies) AS total_copies

        FROM sierra_view.order_record o
        JOIN sierra_view.subfield sent
          ON o.id = sent.record_id
          AND sent.field_type_code = 'b'
          AND sent.tag = 'b'
          AND TO_DATE(SUBSTRING(sent.content, '\d{2}\-\d{2}\-\d{4}'),'MM-DD-YYYY') IS NOT NULL
        JOIN sierra_view.order_record_cmf cmf
          ON o.id = cmf.order_record_id
          AND cmf.location_code != 'multi'
        JOIN sierra_view.accounting_unit a
          ON o.accounting_unit_code_num = a.code_num
        JOIN sierra_view.fund_master f
          ON cmf.fund_code::INT = f.code_num
          AND a.id = f.accounting_unit_id

        WHERE
          o.accounting_unit_code_num = 4
          AND TO_DATE(SUBSTRING(sent.content, '\d{2}\-\d{2}\-\d{4}'),'MM-DD-YYYY') >= CURRENT_DATE - INTERVAL '1 month'

        GROUP BY 1
        """

#Opening a session and querying the database for weekly new items
cursor = conn.cursor()
cursor.execute(query)
#For now, just storing the data in a variable. We'll use it later.
rows = cursor.fetchall()
conn.close()

#Creating the Excel file for staff
workbook = xlsxwriter.Workbook(excelfile, {'remove_timezone': True})
worksheet = workbook.add_worksheet()


#Formatting our Excel worksheet
worksheet.set_landscape()
worksheet.hide_gridlines(0)

#Formatting Cells
eformat= workbook.add_format({'text_wrap': True, 'valign': 'top'})
eformatlabel= workbook.add_format({'text_wrap': True, 'valign': 'top', 'bold': True})


# Setting the column widths
worksheet.set_column(0,0,13.29)
worksheet.set_column(1,1,13.29)
worksheet.set_column(2,2,13.29)

#Inserting a header
worksheet.set_header('Bedford Monthly FTPd Orders')

# Adding column labels
worksheet.write(0,0,'VendorCode', eformatlabel)
worksheet.write(0,1,'TotalRecords', eformatlabel)
worksheet.write(0,2,'TotalCopies', eformatlabel)

# Writing the report for staff to the Excel worksheet
for rownum, row in enumerate(rows):
    worksheet.write(rownum+1,0,row[0], eformat)
    worksheet.write(rownum+1,1,row[1], eformat)
    worksheet.write(rownum+1,2,row[2], eformat)
    
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
