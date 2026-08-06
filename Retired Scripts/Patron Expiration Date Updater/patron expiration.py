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

def mod_patron(patron_id,exp_date):
    config = configparser.ConfigParser()
    config.read('api_info.ini')
    token = get_token()
    url = config['api']['base_url'] + "/patrons/" + patron_id
    header = {"Authorization": "Bearer " + token, "Content-Type": "application/json;charset=UTF-8"}
    #True/False must be titlecase
    payload = {"expirationDate": exp_date}
    request = requests.put(url, data=json.dumps(payload), headers = header)
    
def main():
    config = configparser.ConfigParser()
    config.read('api_info.ini')
        #Connecting to Sierra PostgreSQL database
    query = """\
    SELECT
	rm.record_num, 
	((rm.creation_date_gmt + INTERVAL '10 YEAR')::DATE)::VARCHAR
	FROM
	sierra_view.record_metadata rm
	WHERE
	rm.record_type_code = 'p'
	AND rm.record_num = '2358623'
    """
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

