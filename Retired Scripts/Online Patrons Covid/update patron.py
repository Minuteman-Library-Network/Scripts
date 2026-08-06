import requests
import json
import configparser
from base64 import b64encode
import psycopg2

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

def mod_patron(patron_id,patron_type):
    config = configparser.ConfigParser()
    config.read('api_info.ini')
    token = get_token()
    url = config['api']['base_url'] + "/patrons/" + patron_id
    header = {"Authorization": "Bearer " + token, "Content-Type": "application/json;charset=UTF-8"}
    #True/False must be titlecase
    payload = {"patronType": patron_type}
    request = requests.put(url, data=json.dumps(payload), headers = header)
    
def main():
    config = configparser.ConfigParser()
    config.read('api_info.ini')
        #Connecting to Sierra PostgreSQL database
    query = "SELECT rm.record_num, CASE WHEN LOWER(a.city) ~ 'acton' THEN 1 WHEN LOWER(a.city) ~ 'arlington' THEN 2 WHEN LOWER(a.city) ~ 'ashland' THEN 3 WHEN (LOWER(a.city) ~ '^bedford' OR LOWER(a.city) ~ 'hanscom') THEN 4 WHEN LOWER(a.city) ~ 'belmont' THEN 5 WHEN LOWER(a.city) ~ 'brookline' THEN 6 WHEN LOWER(a.city) ~ 'cambridge' THEN 7 WHEN LOWER(a.city) ~ 'concord' THEN 8 WHEN LOWER(a.city) ~ 'dedham' THEN 10 WHEN LOWER(a.city) ~ '^dover' THEN 11 WHEN LOWER(a.city) ~ 'framingham' THEN 12 WHEN LOWER(a.city) ~ 'franklin' THEN 14 WHEN LOWER(a.city) ~ 'holliston' THEN 15 WHEN LOWER(a.city) ~ 'lexington' THEN 17 WHEN LOWER(a.city) ~ 'lincoln' THEN 18 WHEN LOWER(a.city) ~ 'maynard' THEN 20 WHEN LOWER(a.city) ~ 'medfield' THEN 21 WHEN LOWER(a.city) ~ 'medford' THEN 22 WHEN LOWER(a.city) ~ 'medway' THEN 23 WHEN LOWER(a.city) ~ 'millis' THEN 24 WHEN LOWER(a.city) ~ 'natick' THEN 26 WHEN LOWER(a.city) ~ 'needham' THEN 27 WHEN (LOWER(a.city) ~ 'newton' OR LOWER(a.city) ~ 'auburndale' OR LOWER(a.city) ~ 'chestnut' OR LOWER(a.city) ~ 'nonatum' OR LOWER(a.city) ~ 'waban') THEN 29 WHEN LOWER(a.city) ~ 'norwood' THEN 30 WHEN LOWER(a.city) ~ 'sherborn' THEN 46 WHEN LOWER(a.city) ~ 'somerville' THEN 31 WHEN LOWER(a.city) ~ 'stow' THEN 32 WHEN LOWER(a.city) ~ 'sudbury' THEN 33 WHEN LOWER(a.city) ~ 'waltham' THEN 34 WHEN LOWER(a.city) ~ 'watertown' THEN 35 WHEN LOWER(a.city) ~ 'wayland' THEN 36 WHEN LOWER(a.city) ~ 'wellesley' THEN 37 WHEN LOWER(a.city) ~ 'weston' THEN 38 WHEN LOWER(a.city) ~ 'westwood' THEN 39 WHEN LOWER(a.city) ~ 'winchester' THEN 40 WHEN LOWER(a.city) ~ 'woburn' THEN 41 END AS ptype FROM sierra_view.patron_record p JOIN sierra_view.patron_record_address a ON p.id = a.patron_record_id JOIN sierra_view.record_metadata rm ON p.id = rm.id WHERE p.ptype_code = '207' AND ( LOWER(a.city) ~ 'acton' OR LOWER(a.city) ~ 'arlington' OR LOWER(a.city) ~ 'ashland' OR LOWER(a.city) ~ 'auburndale' OR (LOWER(a.city) ~ 'bedford' AND LOWER(a.city) !~ 'new') OR LOWER(a.city) ~ 'belmont' OR LOWER(a.city) ~ 'brookline' OR LOWER(a.city) ~ 'cambridge' OR LOWER(a.city) ~ 'chestnut' OR (LOWER(a.city) ~ 'concord' AND a.postal_code != '03301') OR LOWER(a.city) ~ 'dedham' OR LOWER(a.city) ~ 'dover' OR LOWER(a.city) ~ 'framingham' OR LOWER(a.city) ~ 'franklin' OR LOWER(a.city) ~ 'holliston' OR LOWER(a.city) ~ 'hanscom' OR LOWER(a.city) ~ 'lexington' OR (LOWER(a.city) ~ 'lincoln' AND a.postal_code !~ '~685') OR LOWER(a.city) ~ 'medfield' OR LOWER(a.city) ~ 'medford' OR LOWER(a.city) ~ 'millis' OR LOWER(a.city) ~ 'maynard' OR LOWER(a.city) ~ 'medway' OR LOWER(a.city) ~ 'natick' OR LOWER(a.city) ~ 'needham' OR LOWER(a.city) ~ 'newton' OR LOWER(a.city) ~ 'nonatum' OR LOWER(a.city) ~ 'norwood' OR LOWER(a.city) ~ 'somerville' OR LOWER(a.city) ~ 'sudbury' OR LOWER(a.city) ~ 'sherborn' OR LOWER(a.city) ~ 'stow(\W|$)' OR LOWER(a.city) ~ 'waltham' OR LOWER(a.city) ~ 'watertown' OR LOWER(a.city) ~ 'wayland' OR LOWER(a.city) ~ 'wellesley' OR LOWER(a.city) ~ 'westwood' OR LOWER(a.city) ~ 'weston' OR LOWER(a.city) ~ 'winchester' OR LOWER(a.city) ~ 'woburn' OR LOWER(a.city) ~ 'waban') ORDER BY 1"
    conn = psycopg2.connect("dbname='iii' user='" + config['api']['sql_user'] + "' host='" + config['api']['sql_host'] + "' port='1032' password='" + config['api']['sql_pass'] + "' sslmode='require'")

    #Opening a session and querying the database for weekly new items
    cursor = conn.cursor()
    cursor.execute(query)
    #For now, just storing the data in a variable. We'll use it later.
    rows = cursor.fetchall()
    conn.close()
    #does not work for item level holds
    for rownum, row in enumerate(rows):
        mod_patron(str(row[0]),row[1])
        print(row[0])
        print(row[1])           
                    
main()

