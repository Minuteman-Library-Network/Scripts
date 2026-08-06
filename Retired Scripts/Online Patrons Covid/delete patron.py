import requests
import json
import configparser
from base64 import b64encode
import psycopg2
import csv
from datetime import date

def get_token():
    # config api    
    config = configparser.ConfigParser()
    config.read('api_info.ini')
    base_url = config['api']['base_url']
    client_key = config['api']['client_key']
    client_secret = config['api']['client_secret']
    auth_string = b64encode((client_key + ':' + client_secret).encode('ascii')).decode('utf-8')
    header = {}
    header["authorization"] = 'Basic ' + auth_string
    header["Content-Type"] = 'application/x-www-form-urlencoded'
    body = {"grant_type": "client_credentials"}
    url = base_url + '/token'
    response = requests.post(url, data=json.dumps(body), headers=header)
    json_response = json.loads(response.text)
    token = json_response["access_token"]
    return token

def mod_patron(patron_id,token,s):
    config = configparser.ConfigParser()
    config.read('api_info.ini')
    url = config['api']['base_url'] + "/patrons/" + patron_id
    header = {"Authorization": "Bearer " + token, "Content-Type": "application/json;charset=UTF-8"}
    #True/False must be titlecase
    #payload = {"patronType": patron_type}
    request = s.delete(url, headers = header)
    
def main():
    config = configparser.ConfigParser()
    config.read('api_info.ini')
        #Connecting to Sierra PostgreSQL database
    query = """\
    SELECT
    m.record_num,
    m.creation_date_gmt::DATE AS created_date,
    n.last_name||', '||n.first_name||' '||n.middle_name AS name,
    z.index_entry AS email
    FROM
    sierra_view.patron_record AS p
    JOIN
    sierra_view.patron_record_phone AS t
    ON
    p.id = t.patron_record_id AND t.patron_record_phone_type_id = '1'
    JOIN
    sierra_view.phrase_entry z
    ON
    p.id = z.record_id AND z.varfield_type_code = 'z'
    JOIN
    sierra_view.phrase_entry w
    ON
    p.id = w.record_id AND w.varfield_type_code = 'w'
    JOIN
    sierra_view.patron_record_address as a
    ON p.id = a.patron_record_id
    JOIN
    sierra_view.record_metadata m
    ON
    p.id = m.id and m.record_type_code = 'p'
    JOIN
    sierra_view.patron_record_fullname n
    ON
    p.id = n.patron_record_id AND n.middle_name ~ '^[A-Za-z]{2,}$' AND n.first_name ~ '^[A-Za-z]{2,}$' AND n.last_name ~ '^[A-Za-z]{2,}$'
    WHERE
    p.patron_agency_code_num = '47'
    AND p.ptype_code = '207'
    AND a.addr1 ~'^[A-Za-z]+$'
    AND a.city ~'^[A-Za-z]+$'
    AND a.postal_code IS NULL
    AND a.region = ''
    AND t.phone_number ~ '\d{10}'
    AND m.creation_date_gmt::DATE >= '2022-07-18'
    ORDER BY 1
    --LIMIT 1
    """
    conn = psycopg2.connect("dbname='iii' user='" + config['api']['sql_user'] + "' host='" + config['api']['sql_host'] + "' port='1032' password='" + config['api']['sql_pass'] + "' sslmode='require'")

    #Opening a session and querying the database for weekly new items
    cursor = conn.cursor()
    cursor.execute(query)
    #For now, just storing the data in a variable. We'll use it later.
    rows = cursor.fetchall()
    conn.close()
    
    csvFile = 'Deleted_Patrons.csv'
    with open(csvFile,'a', encoding='utf-8', newline='') as tempFile:
        myFile = csv.writer(tempFile, delimiter='|')
        myFile.writerows(rows)
    tempFile.close()
    
    
    token = get_token()
    s = requests.Session()
    #does not work for item level holds
    for rownum, row in enumerate(rows):
        mod_patron(str(row[0]),token,s)
        print(row[0])       
                    
main()

