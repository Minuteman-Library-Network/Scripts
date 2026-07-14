#!/usr/bin/env python3

#Run in py313
"""
Jeremy Goldstein
Minuteman Library Network
Script used to generate a monthly report of new patron records the include common data entry errors such as invalid barcodes or malformed addresses.
Reports are produced as Excel files that are then uploaded to our staff intrenet site for distribution, via sftp.
"""

import psycopg2
import xlsxwriter
import os
import pysftp
import configparser
import sys
import time
from datetime import date
import re
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email.utils import formatdate
from email import encoders
import traceback

# run sql query against Sierra database and return results
def run_query(query,ptypes):
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
    cursor.execute(query, (tuple(ptypes),))
    # Storing the results in a variable. We'll use it later.
    rows = cursor.fetchall()
    # close database connection
    conn.close()
    # return variable containing query results
    return rows

#convert sql query results into formatted excel file
def academic_excel_writer(query_results,excel_file):
	
    #Creating the Excel file for staff
    workbook = xlsxwriter.Workbook(excel_file,{'remove_timezone': True})
    worksheet = workbook.add_worksheet()

    #Formatting our Excel worksheet
    worksheet.set_landscape()
    worksheet.hide_gridlines(0)

    #Formatting Cells
    eformat= workbook.add_format({'text_wrap': True, 'valign': 'top'})
    eformatlabel= workbook.add_format({'text_wrap': True, 'valign': 'top', 'bold': True})

    #Setting the column widths
    worksheet.set_column(0,0,15)
    worksheet.set_column(1,1,18.86)
    worksheet.set_column(2,2,5.43)
    worksheet.set_column(3,3,18.43)
    worksheet.set_column(4,4,6.43)
    worksheet.set_column(5,5,28.71)
    worksheet.set_column(6,6,12.43)
    worksheet.set_column(7,7,8.29)
    worksheet.set_column(8,8,12.55)
    worksheet.set_column(9,9,12.55)

    #Inserting a header
    worksheet.set_header('new patrons with data errors')

    #Adding column labels
    worksheet.write(0,0,'Record_Number', eformatlabel)
    worksheet.write(0,1,'Barcode', eformatlabel)
    worksheet.write(0,2,'PType', eformatlabel)
    worksheet.write(0,3,'Home_Library_Code', eformatlabel)
    worksheet.write(0,4,'Blank', eformatlabel)
    worksheet.write(0,5,'Street_Address', eformatlabel)
    worksheet.write(0,6,'City', eformatlabel)
    worksheet.write(0,7,'Zip_Code', eformatlabel)
    worksheet.write(0,8,'Phone_Num', eformatlabel)
    worksheet.write(0,9,'Alt_Phone', eformatlabel)

    #Writing the report for staff to the Excel worksheet
    for rownum, row in enumerate(query_results):
        worksheet.write(rownum+1,0,row[0], eformat)
        if not re.search('^\d{14}',str(row[1])):
            worksheet.write(rownum+1,1,row[1], eformatlabel)
        else:
            worksheet.write(rownum+1,1,row[1], eformat)
        worksheet.write(rownum+1,2,row[2], eformat)
        if not str(row[3]).endswith('z'):
            worksheet.write(rownum+1,3,row[3], eformatlabel)
        else:
            worksheet.write(rownum+1,3,row[3], eformat)
        if (not re.search('^[\dPp]',str(row[5]))):
    	    worksheet.write(rownum+1,5,row[5], eformatlabel)
        else:
            worksheet.write(rownum+1,5,row[5], eformat)
        worksheet.write(rownum+1,6,row[6], eformat)
        if (not re.search('^\d{5}',str(row[7])) and not re.search('^\d{5}([\-]\d{4})',str(row[8]))):
            worksheet.write(rownum+1,7,row[7], eformatlabel)
        else:
            worksheet.write(rownum+1,7,row[7], eformat)
        if (not re.search('^\d{3}([\-]\d{3})([\-]\d{4})',str(row[8])) and not re.search('^\d{3}([ ]\d{3})([ ]\d{4})',str(row[8]))):
            worksheet.write(rownum+1,8,row[8], eformatlabel)
        else:
            worksheet.write(rownum+1,8,row[8], eformat)
        if (not re.search('^\d{3}([\-]\d{3})([\-]\d{4})',str(row[9])) and not re.search('^\d{3}([ ]\d{3})([ ]\d{4})',str(row[9]))):
            worksheet.write(rownum+1,9,row[9], eformatlabel)
        else:
            worksheet.write(rownum+1,9,row[9], eformat)
    
    workbook.close()
    
    return excel_file

def excel_writer(query_results,excel_file,ma_town):
	
    workbook = xlsxwriter.Workbook(excel_file,{'remove_timezone': True})
    worksheet = workbook.add_worksheet()

    #Formatting our Excel worksheet
    worksheet.set_landscape()
    worksheet.hide_gridlines(0)

    #Formatting Cells
    eformat= workbook.add_format({'text_wrap': True, 'valign': 'top'})
    eformatlabel= workbook.add_format({'text_wrap': True, 'valign': 'top', 'bold': True})

    #Setting the column widths
    worksheet.set_column(0,0,15)
    worksheet.set_column(1,1,18.86)
    worksheet.set_column(2,2,5.43)
    worksheet.set_column(3,3,10.86)
    worksheet.set_column(4,4,18.43)
    worksheet.set_column(5,5,6.43)
    worksheet.set_column(6,6,28.71)
    worksheet.set_column(7,7,12.43)
    worksheet.set_column(8,8,8.29)
    worksheet.set_column(9,9,12.55)
    worksheet.set_column(10,10,12.55)

    #Inserting a header
    worksheet.set_header('new patrons with data errors')

    #Adding column labels
    worksheet.write(0,0,'Record_Number', eformatlabel)
    worksheet.write(0,1,'Barcode', eformatlabel)
    worksheet.write(0,2,'PType', eformatlabel)
    worksheet.write(0,3,'Mass_Town', eformatlabel)
    worksheet.write(0,4,'Home_Library_Code', eformatlabel)
    worksheet.write(0,5,'Blank', eformatlabel)
    worksheet.write(0,6,'Street_Address', eformatlabel)
    worksheet.write(0,7,'City', eformatlabel)
    worksheet.write(0,8,'Zip_Code', eformatlabel)
    worksheet.write(0,9,'Phone_Num', eformatlabel)
    worksheet.write(0,10,'Alt_Phone', eformatlabel)

    #Writing the report for staff to the Excel worksheet
    for rownum, row in enumerate(query_results):
        worksheet.write(rownum+1,0,row[0], eformat)
        if not re.search('^\d{14}',str(row[1])):
            worksheet.write(rownum+1,1,row[1], eformatlabel)
        else:
            worksheet.write(rownum+1,1,row[1], eformat)
        worksheet.write(rownum+1,2,row[2], eformat)
        if str(row[3]) != ma_town:
            worksheet.write(rownum+1,3,row[3], eformatlabel)
        else:
            worksheet.write(rownum+1,3,row[3], eformat)
        if not str(row[4]).endswith('z'):
            worksheet.write(rownum+1,4,row[4], eformatlabel)
        else:
            worksheet.write(rownum+1,4,row[4], eformat)
        if (not re.search('^[\dPp]',str(row[6]))):
    	    worksheet.write(rownum+1,6,row[6], eformatlabel)
        else:
            worksheet.write(rownum+1,6,row[6], eformat)
        worksheet.write(rownum+1,7,row[7], eformat)
        if (not re.search('^\d{5}',str(row[8])) and not re.search('^\d{5}([\-]\d{4})',str(row[8]))):
            worksheet.write(rownum+1,8,row[8], eformatlabel)
        else:
            worksheet.write(rownum+1,8,row[8], eformat)
        if (not re.search('^\d{3}([\-]\d{3})([\-]\d{4})',str(row[9])) and not re.search('^\d{3}([ ]\d{3})([ ]\d{4})',str(row[9]))):
            worksheet.write(rownum+1,9,row[9], eformatlabel)
        else:
            worksheet.write(rownum+1,9,row[9], eformat)
        if (not re.search('^\d{3}([\-]\d{3})([\-]\d{4})',str(row[10])) and not re.search('^\d{3}([ ]\d{3})([ ]\d{4})',str(row[10]))):
            worksheet.write(rownum+1,10,row[10], eformatlabel)
        else:
            worksheet.write(rownum+1,10,row[10], eformat)
    
    workbook.close()
    
    return excel_file

#upload report to SIC directory and optionally remove older files
def sftp_file(local_file,library):

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

    srv.cwd('/reports/Library-Specific Reports/'+library+'/New Patrons with Data Errors/')
    srv.put(local_file)

    #remove old file

    for fname in srv.listdir_attr():
        fullpath = '/reports/Library-Specific Reports/'+library+'/New Patrons with Data Errors/{}'.format(fname.filename)
        #time tracked in seconds, st_mtime is time last modified
        name = str(fname.filename)
        if (name != 'meta.json') and ((time.time() - fname.st_mtime) // (24 * 3600) >= 90):
            srv.remove(fullpath)

    srv.close()
    os.remove(local_file)


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
    
def main(library,libcode,ptypes,ma_town = None):
    try:
        ma_town_filter = ""
        if ma_town:
            ma_town_filter = "OR p.pcode3 != '{}'".format(ma_town)
        query = r"""
        SELECT
          id2reckey(p.id)||'a' AS record_num,
          p.barcode,
          p.ptype_code AS ptype,
          p.pcode3 AS mass_town,
          p.home_library_code,
          p.patron_agency_code_num AS agency,
          a.addr1 AS street_address,
          a.city,
          a.postal_code AS zip_code,
          t.phone_number,
          u.phone_number AS alt_phont_num

        FROM sierra_view.patron_view AS p
        JOIN sierra_view.record_metadata m
          ON p.id = m.id
        JOIN sierra_view.patron_record_address as a
          ON p.id = a.patron_record_id
          AND a.patron_record_address_type_id = '1'
        --telephone field
        LEFT JOIN sierra_view.patron_record_phone AS t
          ON p.id = t.patron_record_id
          AND t.patron_record_phone_type_id = '1'
        --alt telephone field
        LEFT JOIN sierra_view.patron_record_phone AS u
          ON p.id = u.patron_record_id
          AND u.patron_record_phone_type_id = '2'

        WHERE m.creation_date_gmt > (CURRENT_DATE - INTERVAL '1 month')
          AND p.ptype_code IN %s
          AND (
            (p.home_library_code !~ 'z$' AND p.home_library_code IS NOT NULL) --failed to use pickup location code
        """ + ma_town_filter + """--ma town does not match ptype, excluded for academic libraries
            OR a.addr1 IS NULL
            OR a.addr1 !~'^[\dPp]'--address doesn't start with a # or PO
            OR a.city IS NULL
            OR (a.postal_code !~ '^\d{5}' AND a.postal_code !~'^\d{5}([\-]\d{4})') --zipcode not ##### or #####-####
            OR p.barcode IS NULL
            OR (p.barcode !~ '^\d{14}' AND p.ptype_code < '301') --barcode not 14 digits
            OR (t.phone_number IS NOT NULL AND (t.phone_number !~'^\d{3}[\-]\d{3}[\-]\d{4}' AND t.phone_number !~'^\d{3}[ ]\d{3}[ ]\d{4}')) -- phone not ###-###-#### or ### ### ####
            OR (u.phone_number IS NOT NULL AND (u.phone_number !~'^\d{3}[\-]\d{3}[\-]\d{4}' AND u.phone_number !~'^\d{3}[ ]\d{3}[ ]\d{4}')) -- alt phone not ###-###-#### or ### ### ####
            )
        ORDER BY 2,1
        """
        query_results = run_query(query,ptypes)
        #Name of Excel File
        excel_file =  "/Scripts/New Patrons With Data Errors/Temp Files/" + libcode + "NewPatronsWithDataErrors{}.xlsx".format(date.today())
        if ma_town == None:
            academic_excel_writer(query_results,excel_file)
        else:    
            excel_writer(query_results,excel_file,ma_town)
        sftp_file("C:\\Scripts\\New Patrons With Data Errors\\Temp Files\\" + libcode + "NewPatronsWithDataErrors{}.xlsx".format(date.today()),library)

    except:
        # read config file with recipient list for email
        config_recipient = configparser.ConfigParser()
        config_recipient.read("C:\\Scripts\\Creds\\emails.ini")
        emailto = config_recipient["script_error"]["recipients"].split()

        # craft email subject and message containing error message details from traceback
        email_subject = "new patrons with data errors: " + library + " script error"
        email_message = (
            "Your script failed with the following error:\n\n" + traceback.format_exc()
        )

        send_email_error(email_subject, email_message, emailto)
        raise

if __name__ == "__main__":
    # run for each library within Minuteman
    main('Acton','ACT',['1','301'],'1')
    main('Arlington','ARL',['2','302'],'3')
    main('Ashland','ASH',['3','303'],'4')
    main('Bedford','BED',['4','304'],'7')
    main('Belmont','BLM',['5','305'],'9')
    main('Brookline','BRK',['6','306'],'19')
    main('Cambridge','CAM',['7','307'],'21')
    main('Concord','CON',['8','308'],'27')
    main('Dedham','DDM',['10','110','310'],'29')
    main('Dean','DEA',['9','159'])
    main('Dover','DOV',['11','311'],'30')
    main('Framingham Public','FPL',['12','312'],'36')
    main('Framingham State','FST',['13','163'])
    main('Franklin','FRK',['14','314'],'37')
    main('Holliston','HOL',['15','115','315'],'43')
    main('Lasell','LAS',['16','116','166'])
    main('Lexington','LEX',['17','117','317'],'50')
    main('Lincoln','LIN',['18','318'],'51')
    main('Maynard','MAY',['20','120','320'],'61')
    main('Medfield','MLD',['21','121','321'],'62')
    main('Medford','MED',['22','122','322'],'63')
    main('Medway','MWY',['23','123'],'64')
    main('Millis','MIL',['24','324'],'69')
    main('Natick','NAT',['26','326'],'75')
    main('Needham','NEE',['27','327'],'76')
    main('Newton','NTN',['29','129','329'],'79')
    main('Norwood','NOR',['30','130','330'],'83')
    main('Olin','OLN',['47','147','197'])
    main('Regis','REG',['45','195'])
    main('Sherborn','SHR',['46','346'],'95')
    main('Somerville','SOM',['31','331'],'98')
    main('Stow','STO',['32','332'],'102')
    main('Sudbury','SUD',['33','133','333'],'103')
    main('Waltham','WLM',['34','334'],'113')
    main('Watertown','WAT',['35','335'],'114')
    main('Wayland','WYL',['36','136','336'],'115')
    main('Wellesley','WEL',['37','137','337'],'116')
    main('Weston','WSN',['38','338'],'119')
    main('Westwood','WWD',['39','339'],'120')
    main('Winchester','WIN',['40','340'],'124')
    main('Woburn','WOB',['41','341'],'126')
