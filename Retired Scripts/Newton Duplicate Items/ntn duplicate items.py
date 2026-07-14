#!/usr/bin/env python3

"""Create and email a list of new items

Author: Gem Stone-Logan
Contact Info: gem.stone-logan@mountainview.gov or gemstonelogan@gmail.com
"""

import psycopg2
import xlsxwriter
import os
import pysftp
import configparser
import sys
import time
from datetime import date


#Name of Excel File
excelfile =  'NTNDuplicateItems{}.xlsx'.format(date.today())

config = configparser.ConfigParser()
config.read('C:\\SQL Reports\\creds\\app_SIC.ini')

#Connecting to Sierra PostgreSQL database
conn = psycopg2.connect( config['db']['connection_string'] )

#Opening a session and querying the database for weekly new items
cursor = conn.cursor()
cursor.execute(open("ntn duplicate items.sql","r").read())
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


# Setting the column widths
worksheet.set_column(0,0,11.14)
worksheet.set_column(1,1,66.14)
worksheet.set_column(2,2,23.86)
worksheet.set_column(3,3,60.43)
worksheet.set_column(4,4,15.43)
worksheet.set_column(5,5,10.43)
worksheet.set_column(6,6,17.43)
worksheet.set_column(7,7,13.14)
worksheet.set_column(8,8,8)
#Inserting a header
worksheet.set_header('Newton Duplicate Items')

# Adding column labels
worksheet.write(0,0,'Bib_Number', eformatlabel)
worksheet.write(0,1,'Title', eformatlabel)
worksheet.write(0,2,'Author', eformatlabel)
worksheet.write(0,3,'Call_Number', eformatlabel)
worksheet.write(0,4,'Mattype', eformatlabel)
worksheet.write(0,5,'Item_Count', eformatlabel)
worksheet.write(0,6,'Last_YTD_Checkouts', eformatlabel)
worksheet.write(0,7,'YTD_Checkouts', eformatlabel)
worksheet.write(0,8,'Turnover', eformatlabel)


# Writing the report for staff to the Excel worksheet
for rownum, row in enumerate(rows):
    worksheet.write(rownum+1,0,row[0], eformat)
    worksheet.write(rownum+1,1,row[1], eformat)
    worksheet.write(rownum+1,2,row[2], eformat)
    worksheet.write(rownum+1,3,row[3], eformat)
    worksheet.write(rownum+1,4,row[4], eformat)
    worksheet.write(rownum+1,5,row[5], eformat)
    worksheet.write(rownum+1,6,row[6], eformat)
    worksheet.write(rownum+1,7,row[7], eformat)
    worksheet.write(rownum+1,8,row[8], eformat)
    
workbook.close()

cnopts = pysftp.CnOpts()

srv = pysftp.Connection(host = config['sic']['sic_host'], username = config['sic']['sic_user'], password= config['sic']['sic_pw'], cnopts=cnopts)

local_file = excelfile

srv.cwd('/reports/Library-Specific Reports/Newton/Custom/')
srv.put(local_file)

#remove old file

for fname in srv.listdir_attr():
    fullpath = '/reports/Library-Specific Reports/Cambridge/New Titles/{}'.format(fname.filename)
    #time tracked in seconds, st_mtime is time last modified
    name = str(fname.filename)
    if (name != 'meta.json') and ((time.time() - fname.st_mtime) // (24 * 3600) >= 1095):
        srv.remove(fullpath)

srv.close()
os.remove(local_file)
