import psycopg2
import os
import configparser
from datetime import datetime
import csv

def benchmark(server):
    csvFile = 'SQL_benchmark.csv'

    config = configparser.ConfigParser()
    config.read('C:\\SQL Reports\\creds\\api_info.ini')

    query = """\
            SELECT *
            FROM sierra_view.item_record
            LIMIT 1000000
            """

    start_time = datetime.now()

    try:
	    # variable connection string should be defined in the imported config file
        if server == 'production':
            conn = psycopg2.connect( config['api']['connection_string'] )
        if server == 'test':
            conn = psycopg2.connect( config['test']['connection_string'] )
    except:
        print("unable to connect to the database")
        clear_connection()
        
    cursor = conn.cursor()
    cursor.execute(query)
    rows = cursor.fetchall()
    conn.close()

    end_time = datetime.now()
    duration = end_time - start_time
    
    with open(csvFile, 'a', encoding = 'utf-8', newline = '') as tempFile: 
        myFile = csv.writer(tempFile, delimiter = ',')
        myFile.writerow([server,start_time,end_time,duration])
    tempFile.close()

benchmark('production')
benchmark('test')
