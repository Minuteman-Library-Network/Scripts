#!/usr/bin/env python3

#olin report waiting on accounting unit assignment

"""Create and upload daily local pickup hold paging list

Based on code from Gem Stone-Logan
"""

import psycopg2
import xlsxwriter
import os
import pysftp
import configparser
import sys
import time
from datetime import date

#convert sql query results into formatted excel file
def excelWriter(query_results,excelfile):

    #Creating the Excel file for staff
    workbook = xlsxwriter.Workbook(excelfile,{'remove_timezone': True})
    worksheet = workbook.add_worksheet()


    #Formatting our Excel worksheet
    worksheet.set_landscape()
    worksheet.hide_gridlines(0)

    #Formatting Cells
    eformat= workbook.add_format({'text_wrap': True, 'valign': 'top'})
    eformatlabel= workbook.add_format({'text_wrap': True, 'valign': 'top', 'bold': True, 'align': 'center'})

    # Setting the column widths
    worksheet.set_column(0,0,10.86)
    worksheet.set_column(1,1,14.43)
    worksheet.set_column(2,2,31.71)
    worksheet.set_column(3,3,39.14)
    worksheet.set_column(4,4,80.43)
    worksheet.set_column(5,5,14.14)
    worksheet.set_column(6,6,18.43)
    worksheet.set_column(7,7,10)

    #Inserting a header
    worksheet.set_header('Local pickup hold paging list')

    # Adding column labels
    worksheet.write(0,0,'Bib Number', eformatlabel)
    worksheet.write(0,1,'Barcode', eformatlabel)
    worksheet.write(0,2,'Call Number', eformatlabel)
    worksheet.write(0,3,'Author', eformatlabel)
    worksheet.write(0,4,'Title', eformatlabel)
    worksheet.write(0,5,'Pickup Location', eformatlabel)
    worksheet.write(0,6,'Item Location', eformatlabel)
    worksheet.write(0,7,'Itype', eformatlabel)


    # Writing the report for staff to the Excel worksheet
    for rownum, row in enumerate(query_results):
        worksheet.write(rownum+1,0,row[1], eformat)
        worksheet.write(rownum+1,1,row[2], eformat)
        worksheet.write(rownum+1,2,row[3], eformat)
        worksheet.write(rownum+1,3,row[4], eformat)
        worksheet.write(rownum+1,4,row[5], eformat)
        worksheet.write(rownum+1,5,row[6], eformat)
        worksheet.write(rownum+1,6,row[7], eformat)
        worksheet.write(rownum+1,7,row[8], eformat)
    
    workbook.close()
    
    return excelfile

#convert sql query results into formatted excel file per custom request for Westwood
def excelWriterWWD(query_results,excelfile):

    #Creating the Excel file for staff
    workbook = xlsxwriter.Workbook(excelfile,{'remove_timezone': True})
    worksheet = workbook.add_worksheet()


    #Formatting our Excel worksheet
    worksheet.set_landscape()
    worksheet.hide_gridlines(0)
    #Inserting a header
    worksheet.set_header('Local pickup hold paging list')
    worksheet.set_page_view()
    worksheet.set_footer('&CPage &P of &N')

    #Formatting Cells
    eformat= workbook.add_format({'text_wrap': True, 'valign': 'top', 'font_size': '14'})
    eformatlabel= workbook.add_format({'text_wrap': True, 'valign': 'top', 'bold': True, 'font_size': '14', 'align': 'center'})

    # Setting the column widths
    worksheet.set_column(0,0,25.71)
    worksheet.set_column(1,1,24.57)
    worksheet.set_column(2,2,33)
    worksheet.set_column(3,3,64)
    worksheet.set_column(4,4,21.43)
    #worksheet.set_column(5,5,14.14)
    #worksheet.set_column(6,6,18.43)
    #worksheet.set_column(7,7,10)

    # Adding column labels
    #worksheet.write(0,0,'Bib Number', eformatlabel)
    worksheet.write(0,0,'Barcode', eformatlabel)
    worksheet.write(0,1,'Call Number', eformatlabel)
    worksheet.write(0,2,'Author', eformatlabel)
    worksheet.write(0,3,'Title', eformatlabel)
    #worksheet.write(0,5,'Pickup Location', eformatlabel)
    worksheet.write(0,4,'Item Location', eformatlabel)
    #worksheet.write(0,7,'Itype', eformatlabel)


    # Writing the report for staff to the Excel worksheet
    for rownum, row in enumerate(query_results):
        #worksheet.write(rownum+1,0,row[1], eformat)
        worksheet.write(rownum+1,0,row[2], eformat)
        worksheet.write(rownum+1,1,row[3], eformat)
        worksheet.write(rownum+1,2,row[4], eformat)
        worksheet.write(rownum+1,3,row[5], eformat)
        #worksheet.write(rownum+1,5,row[6], eformat)
        worksheet.write(rownum+1,4,row[7], eformat)
        #worksheet.write(rownum+1,7,row[8], eformat)
    
    workbook.close()
    
    return excelfile

#connect to Sierra-db and store results of an sql query
def runquery(query):

    # import configuration file containing our connection string
    # app.ini looks like the following
    #[db]
    #connection_string = dbname='iii' user='PUT_USERNAME_HERE' host='sierra-db.library-name.org' password='PUT_PASSWORD_HERE' port=1032

    config = configparser.ConfigParser()
    config.read('C:\\SQL Reports\\creds\\app_SIC.ini')
      
    try:
	    # variable connection string should be defined in the imported config file
        conn = psycopg2.connect( config['db']['connection_string'] )
    except:
        print("unable to connect to the database")
        clear_connection()
        return
        
    #Opening a session and querying the database for weekly new items
    cursor = conn.cursor()
    cursor.execute(open(query,"r").read())
    #For now, just storing the data in a variable. We'll use it later.
    rows = cursor.fetchall()
    conn.close()
    
    return rows

#upload report to SIC directory and optionally remove older files
def ftp_file(local_file,library,libcode):

    config = configparser.ConfigParser()
    config.read('C:\\SQL Reports\\creds\\app_SIC.ini')
    
    cnopts = pysftp.CnOpts()

    srv = pysftp.Connection(host = config['sic']['sic_host'], username = config['sic']['sic_user'], password= config['sic']['sic_pw'], cnopts=cnopts)

    local_file = local_file

    srv.cwd('/reports/Library-Specific Reports/'+library+'/Local Pickup Hold Paging List/')
    
    #remove old file
    fullpath = '/reports/Library-Specific Reports/'+library+'/Local Pickup Hold Paging List/'+libcode+'LocalPickupHoldPagingList.xlsx'
    srv.remove(fullpath)

    srv.put(local_file)


    srv.close()
    os.remove(local_file)

def main(library,libcode):
	
    tempFile = runquery(libcode.lower()+" local pickup hold paging list.sql")
    #Name of Excel File
    excelfile =  libcode+'LocalPickupHoldPagingList.xlsx'
    if libcode.startswith('WW'):
        local_file = excelWriterWWD(tempFile,excelfile)
    else:
        local_file = excelWriter(tempFile,excelfile)
    ftp_file(local_file,library,libcode)

main('Acton','ACT')
main('Arlington','ARL')
main('Arlington','AR2')
main('Ashland','ASH')
main('Bedford','BED')
main('Belmont','BLM')
main('Brookline','BRK')
main('Brookline','BR2')
main('Brookline','BR3')
main('Cambridge','CAM')
main('Cambridge','CA3')
main('Cambridge','CA4')
main('Cambridge','CA5')
main('Cambridge','CA6')
main('Cambridge','CA7')
main('Cambridge','CA8')
main('Cambridge','CA9')
main('Concord','CON')
main('Concord','CO2')
main('Dedham','DDM')
main('Dedham','DD2')
main('Dean','DEA')
main('Dover','DOV')
main('Framingham Public','FPL')
main('Framingham Public','FP2')
main('Framingham State','FST')
main('Franklin','FRK')
main('Holliston','HOL')
main('Lasell','LAS')
main('Lexington','LEX')
main('Lincoln','LIN')
main('Maynard','MAY')
main('Medfield','MLD')
main('Medford','MED')
main('Medway','MWY')
main('Millis','MIL')
main('Natick','NAT')
main('Natick','NA2')
main('Needham','NEE')
main('Newton','NTN')
main('Norwood','NOR')
main('Olin','OLN')
main('Pine Manor','PMC')
main('Regis','REG')
main('Sherborn','SHR')
main('Somerville','SOM')
main('Somerville','SO2')
main('Somerville','SO3')
main('Stow','STO')
main('Sudbury','SUD')
main('Waltham','WLM')
main('Watertown','WAT')
main('Wayland','WYL')
main('Wellesley','WEL')
main('Wellesley','WE2')
main('Wellesley','WE3')
main('Weston','WSN')
main('Westwood','WWD')
main('Westwood','WW2')
main('Winchester','WIN')
main('Woburn','WOB')
