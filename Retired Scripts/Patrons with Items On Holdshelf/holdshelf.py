#!/usr/bin/env python3

"""Create montly collection dev by scat report and place in reports folder

Based on code from Gem Stone-Logan
"""

import psycopg2
import xlsxwriter
import os
import shutil
import configparser
from datetime import date

#connect to Sierra-db and store results of an sql query
def runquery(pickup_location):

    # import configuration file containing our connection string
    # app.ini looks like the following
    #[db]
    #connection_string = dbname='iii' user='PUT_USERNAME_HERE' host='sierra-db.library-name.org' password='PUT_PASSWORD_HERE' port=1032

    config = configparser.ConfigParser()
    config.read('C:\\SQL Reports\\creds\\app.ini')
    query = "SELECT DISTINCT id2reckey(h.patron_record_id)||'a' AS patron_number,COALESCE(s.content,'') AS email,p.last_name||', '||p.first_name||' '||p.middle_name AS name,COALESCE(ph1.phone_number,'') AS phone,COALESCE(ph2.phone_number,'') AS phone2,STRING_AGG(b.best_title, ', ') AS titles FROM sierra_view.hold h LEFT JOIN sierra_view.subfield s ON h.patron_record_id = s.record_id AND s.field_type_code = 'z' AND s.occ_num = 0 JOIN sierra_view.patron_record_fullname p ON h.patron_record_id = p.patron_record_id AND p.display_order = '0' LEFT JOIN sierra_view.patron_record_phone ph1 ON h.patron_record_id = ph1.patron_record_id AND ph1.display_order = 0 AND ph1.patron_record_phone_type_id = '1' LEFT JOIN sierra_view.patron_record_phone ph2 ON h.patron_record_id = ph2.patron_record_id AND ph2.display_order = 0 AND ph2.patron_record_phone_type_id = '2' JOIN sierra_view.bib_record_item_record_link l ON h.record_id = l.item_record_id JOIN sierra_view.bib_record_property b ON l.bib_record_id = b.bib_record_id WHERE h.pickup_location_code ~ '^"+pickup_location+"' AND h.status IN ('b','i') GROUP BY 1,2,3,4,5 ORDER BY 3"
      
    try:
	    # variable connection string should be defined in the imported config file
        conn = psycopg2.connect( config['db']['connection_string'] )
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

def excelWriter(query_results,excelfile):

    #Creating the Excel file for staff
    workbook = xlsxwriter.Workbook(excelfile,{'remove_timezone': True})
    worksheet = workbook.add_worksheet()


    #Formatting our Excel worksheet
    worksheet.set_landscape()
    worksheet.hide_gridlines(0)

    #Formatting Cells
    eformat= workbook.add_format({'text_wrap': True, 'valign': 'top'})
    eformatlabel= workbook.add_format({'text_wrap': True, 'valign': 'top', 'bold': True})

    # Setting the column widths
    worksheet.set_column(0,0,15.57)
    worksheet.set_column(1,1,33.86)
    worksheet.set_column(2,2,22.29)
    worksheet.set_column(3,3,12.29)
    worksheet.set_column(4,4,12.29)
    worksheet.set_column(5,5,115)

    #Inserting a header
    worksheet.set_header('Patrons With Items On Holdshelf')

    # Adding column labels
    worksheet.write(0,0,'Patron_number', eformatlabel)
    worksheet.write(0,1,'Email', eformatlabel)
    worksheet.write(0,2,'Name', eformatlabel)
    worksheet.write(0,3,'Phone', eformatlabel)
    worksheet.write(0,4,'Phone2', eformatlabel)
    worksheet.write(0,5,'Titles', eformatlabel)

    # Writing the report for staff to the Excel worksheet
    for rownum, row in enumerate(query_results):
        worksheet.write(rownum+1,0,row[0], eformat)
        worksheet.write(rownum+1,1,row[1], eformat)
        worksheet.write(rownum+1,2,row[2], eformat)
        worksheet.write(rownum+1,3,row[3], eformat)
        worksheet.write(rownum+1,4,row[4], eformat)
        worksheet.write(rownum+1,5,row[5], eformat)
    
    workbook.close()
    
    return excelfile

def remove_old_file(library,libcode):
    
    os.chdir('W:/'+library)
    for fname in os.listdir("."):
        if os.path.isfile(fname) and fname.startswith(libcode+"PatronsWithItemsOnHoldshelf"):
            os.remove(fname)
    os.chdir('C:/SQL Reports/Patrons with Items On Holdshelf')

def main(library,libcode):
	
    remove_old_file(library,libcode)
    tempFile = runquery(libcode.lower())
    #Name of Excel File
    excelfile =  'W:\\'+library+'\\'+libcode+'PatronsWithItemsOnHoldshelf{}.xlsx'.format(date.today())
    excelWriter(tempFile,excelfile)

main('norwood','NOR')
#main('holliston','HOL')
#main('medfield','MLD')
