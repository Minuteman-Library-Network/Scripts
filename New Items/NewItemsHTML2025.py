#!/usr/bin/env python3

# experiment Liza began but never got working

"""Create and email monthly New Items Reports

Based on code from Gem Stone-Logan
"""

import psycopg2
import xlsxwriter
import os
import pysftp
import re
import shutil
import configparser
import sys
import time
from datetime import date, timedelta

#To calculate last month's name
last_month = date.today().replace(day=1) - timedelta(1)

#convert sql query results into formatted html file
def text_to_html(input_file, output_file):
    # Read the text file
    with open(input_file, 'r', encoding='utf-8') as file:
        text = file.read()

    # Basic HTML structure
    html_content = f"""
    <!DOCTYPE html>
    <html lang="en">
    <head>
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <title>Converted Text</title>
        <style>
            body {{
                font-family: Arial, sans-serif;
                line-height: 1.6;
                margin: 20px;
            }}
        </style>
    </head>
    <body>
        <pre>{text}</pre>
    </body>
    </html>
    """

    # Write to the output HTML file
    with open(output_file, 'w', encoding='utf-8') as file:
        file.write(html_content)

    print(f"HTML file saved as {output_file}")


#connect to Sierra-db and store results of an sql query
def runquery(libcode):

    # import configuration file containing our connection string
    # app.ini looks like the following
    #[db]
    #connection_string = dbname='iii' user='PUT_USERNAME_HERE' host='sierra-db.library-name.org' password='PUT_PASSWORD_HERE' port=1032

    config = configparser.ConfigParser()
    config.read('C:\\SQL Reports\\creds\\app.ini')
      
    try:
	    # variable connection string should be defined in the imported config file
        conn = psycopg2.connect( config['db']['connection_string'] )
    except:
        print("unable to connect to the database")
        clear_connection()
        return
        
    
    query = "SELECT DISTINCT SUBSTRING(i.location_code, 1, 5), " \
            "iprop.barcode, it.name, i.icode1, TRIM(regexp_replace(iprop.call_number,'\|.',' ','g')), " \
            "v.field_content, bprop.best_author, bprop.best_title " \
            "FROM sierra_view.item_view i " \
            "JOIN sierra_view.bib_record_item_record_link bilink ON bilink.item_record_id = i.id " \
            "JOIN sierra_view.bib_record_property bprop ON bprop.bib_record_id = bilink.bib_record_id " \
            "JOIN sierra_view.item_record_property iprop ON iprop.item_record_id = i.id " \
            "JOIN sierra_view.record_metadata ON record_metadata.id = i.id "  \
            "JOIN sierra_view.itype_property ip ON i.itype_code_num = ip.code_num " \
            "JOIN sierra_view.itype_property_name it ON ip.id = it.itype_property_id " \
            "LEFT JOIN sierra_view.varfield v ON v.record_id = i.id AND v.varfield_type_code = 'v' " \
            "WHERE record_metadata.creation_date_gmt >= date_trunc('month', current_date - interval '1' month) and record_metadata.creation_date_gmt < date_trunc('month', current_date) " \
            "AND i.itype_code_num NOT IN ('241', '255', '242', '10', '107', '158') " \
            "AND i.item_message_code <> 'f' AND i.location_code ~ '^"+libcode.lower()+"' ORDER BY 1, 5"
    #Opening a session and querying the database for monthly new items
    cursor = conn.cursor()
    cursor.execute(query)
    #For now, just storing the data in a variable. We'll use it later.
    rows = cursor.fetchall()
    conn.close()
    
    return rows
    
#upload report to SIC directory and optionally remove older files
def ftp_file(local_file,library):

    config = configparser.ConfigParser()
    config.read('C:\\SQL Reports\\creds\\app_SIC.ini')

    cnopts = pysftp.CnOpts()

    srv = pysftp.Connection(host = config['sic']['sic_host'], username = config['sic']['sic_user'], password= config['sic']['sic_pw'], cnopts=cnopts)

    local_file = local_file

    srv.cwd('/reports/Library-Specific Reports/'+library+'/New Items/')
    srv.put(local_file)

    #remove old file

#    for fname in srv.listdir_attr():
#        fullpath = '/reports/Library-Specific Reports/'+library+'/New Items/{}'.format(fname.filename)
#        #time tracked in seconds, st_mtime is time last modified
#        name = str(fname.filename)
#        if (name != 'meta.json') and  ((time.time() - fname.st_mtime) // (24 * 3600) >= 90):
#            srv.remove(fullpath)

#    srv.close()
#    os.remove(local_file)

def main(library,libcode):
	
    tempFile = runquery(libcode[0:2])
    
    #archive old Excel Files
#    os.chdir('W:/'+library)
#    for fname in os.listdir("."):
#        if os.path.isfile(fname) and fname.startswith(libcode+"Withdrawn"):
#            shutil.move(os.path.join('W:/'+library+'/',fname),'W:/'+library+'/archive')
#            os.chdir('W:/'+library)
#    os.chdir('C:/SQL Reports/New Items')
   
    #Name of Excel File
    htmlfile =  libcode+'NewItems{}.html'.format(last_month.strftime("%b%Y"))
    local_file = text_to_html(tempFile,htmlfile)
#    ftp_file(local_file, library)

main('Acton','ACT')
main('Arlington','ARL')
main('Ashland','ASH')
main('Bedford','BED')
main('Belmont','BLM')
main('Brookline','BRK')
main('Cambridge','CAM')
main('Concord','CON')
main('Dedham','DDM')
main('Dean','DEA')
main('Dover','DOV')
main('Framingham Public','FPL')
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
main('Needham','NEE')
main('Newton','NTN')
main('Norwood','NOR')
main('Olin','OLN')
main('Pine Manor','PMC')
main('Regis','REG')
main('Sherborn','SHR')
main('Somerville','SOM')
main('Stow','STO')
main('Sudbury','SUD')
main('Waltham','WLM')
main('Watertown','WAT')
main('Wayland','WYL')
main('Wellesley','WEL')
main('Weston','WSN')
main('Westwood','WWD')
main('Winchester','WIN')
main('Woburn','WOB')
