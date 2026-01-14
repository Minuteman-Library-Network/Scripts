#!/usr/bin/env python3

# Run in py38

"""
Jeremy Goldstein
jgoldstein@minlib.net
Minuteman Library Network

Script used to batch update all holds with a given pickup location and a status of on hold
to a different specified pickup location.
Holds that cannot be updated due to Sierra limitations will simply be skipped past
"""

import requests
import json
import configparser
from base64 import b64encode
import psycopg2
from datetime import datetime
from datetime import timedelta


# function to generate access token for use with Sierra API
def get_token():
    # config api
    config = configparser.ConfigParser()
    config.read("C:\\Scripts\\Creds\\config.ini")
    base_url = config["api"]["base_url"]
    client_key = config["api"]["client_key"]
    client_secret = config["api"]["client_secret"]
    auth_string = b64encode((client_key + ":" + client_secret).encode("ascii")).decode(
        "utf-8"
    )
    header = {}
    header["authorization"] = "Basic " + auth_string
    header["Content-Type"] = "application/x-www-form-urlencoded"
    body = {"grant_type": "client_credentials"}
    url = base_url + "/token"
    response = requests.post(url, data=json.dumps(body), headers=header, verify=False)
    json_response = json.loads(response.text)
    token = json_response["access_token"]
    return token


# function takes a sql query as a parameter, connects to a database and returns the results
def run_query(query):
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
    cursor.execute(query)
    # Storing the results in a variable. We'll use it later.
    rows = cursor.fetchall()
    # close database connection
    conn.close()
    # return variable containing query results
    return rows


# function will update the pickup location for a hold to a specified value, so long as Sierra's rules permit it
def mod_hold(hold_id, is_frozen, new_location, token, s):
    config = configparser.ConfigParser()
    config.read("C:\\Scripts\\Creds\\config.ini")
    url = config["api"]["base_url"] + "/patrons/holds/" + hold_id
    header = {
        "Authorization": "Bearer " + token,
        "Content-Type": "application/json;charset=UTF-8",
    }
    payload = {"pickupLocation": new_location, "freeze": is_frozen}
    request = s.put(url, data=json.dumps(payload), headers=header, verify=False)


def main():
    config = configparser.ConfigParser()
    config.read("api_info.ini")
    old_location = input("Enter location code to search, or type 'q' to quit.\n")
    new_location = input("enter location code you wish to change to.\n")
    while old_location != "q":
        if not old_location.endswith("z"):
            print("Invalid location code entered, please try again")
            break
        if not new_location.endswith("z"):
            print("Invalid location code entered, please try again")
            break
        confirm = input(
            "This will change holds with location "
            + old_location
            + " to "
            + new_location
            + ".  Do you wish to proceed?  type 'y' to continue.\n"
        )
        if not confirm == "y":
            print("\nThis program will now quit.  Goodbye.")
            break
        query = (
            """
        select
          id,
          is_frozen
        FROM sierra_view.hold
        WHERE pickup_location_code = '"""
            + old_location
            + """'
        AND status = '0'
        """
        )

        holds_list = run_query(query)

        # initialize Sierra API
        # track api token expiration time to know when it must be refreshed
        expiration_time = datetime.now() + timedelta(seconds=3600)
        token = get_token()
        s = requests.Session()

        # does not work for item level holds
        for hold_id, is_frozen in holds_list:
            print("hold id: " + str(hold_id))
            print("is_frozen: " + str(is_frozen))
            # refresh API token if present one has expired
            if datetime.now() >= expiration_time:
                print("refreshing token")
                expiration_time = datetime.now() + timedelta(seconds=3600)
                token = get_token()
            mod_hold(str(hold_id), is_frozen, new_location, token, s)

        old_location = input("\nEnter another location code, or press 'q' to quit.\n")

    print("\nThis program will now quit.  Goodbye.")


main()
