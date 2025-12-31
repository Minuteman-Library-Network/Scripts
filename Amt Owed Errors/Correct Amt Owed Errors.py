#!/usr/bin/env python3

# Run in py38

"""
Identify patrons with amt owed discrepancies in their patron records
due to a mismatch between the patron_record.amt_owed and the actual amount of active fines
Correct those errors by creating a manual charge in the difference and waiving it via the API

Author: Jeremy Goldstein
Contact Info: jgoldstein@minlib.net
"""

import requests
import json
import configparser
from base64 import b64encode
import psycopg2

# make sure to remove verify=False from all request calls, in place to handle expired certificate on test
# also temporary limit on error query from testing

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

# function to create a manual charge on a patron account
def manual_charge(patron_id, amount, location):
    config = configparser.ConfigParser()
    config.read("C:\\Scripts\\Creds\\config.ini")

    token = get_token()
    url = config["api"]["base_url"] + "/patrons/" + patron_id + "/fines/charge"
    header = {
        "Authorization": "Bearer " + token,
        "Content-Type": "application/json;charset=UTF-8",
    }
    payload = {"amount": amount, "reason": "Residual fine", "location": location}
    request = requests.post(url, data=json.dumps(payload), headers=header, verify=False)

# function to waive a fine
def clear_fine(patron_id, invoiceNumber):
    config = configparser.ConfigParser()
    config.read("C:\\Scripts\\Creds\\config.ini")

    token = get_token()
    url = config["api"]["base_url"] + "/patrons/" + patron_id + "/fines/payment"
    header = {
        "Authorization": "Bearer " + token,
        "Content-Type": "application/json;charset=UTF-8",
    }
    payload = {
        "payments": [
            {"amount": 0, "paymentType": 2, "invoiceNumber": "" + invoiceNumber + ""}
        ]
    }
    request = requests.put(url, data=json.dumps(payload), headers=header, verify=False)


def main():
    # query to identify patrons with amt owed errors
    error_query = """\
            SELECT
              rm.record_num,
              (
                p.owed_amt * 100
                - (
                  SUM(
                    COALESCE(f.item_charge_amt * 100, 0)
                    + COALESCE(f.processing_fee_amt * 100, 0)
                    + COALESCE(f.billing_fee_amt * 100, 0)
                    - COALESCE(f.paid_amt * 100, 0)
                  )
                )
              )::INT AS FineDiscrepancy,
              p.home_library_code AS location
  
            FROM sierra_view.record_metadata rm
            JOIN sierra_view.patron_record p
              ON p.id = rm.id
            LEFT JOIN sierra_view.fine f
              ON f.patron_record_id = p.id

            GROUP BY rm.record_num, p.owed_amt,3
            HAVING p.owed_amt != SUM(COALESCE(f.item_charge_amt, 0.00) + COALESCE(f.processing_fee_amt, 0.00) + COALESCE(f.billing_fee_amt, 0.00) - COALESCE(f.paid_amt, 0.00))
            """
    # query to retrieve data created by manual_charge function, in order to waive the charge
    manual_charge_query = """\
        SELECT
          rm.record_num,
          f.invoice_num::varchar

        FROM sierra_view.fine f
        JOIN sierra_view.record_metadata rm
            ON f.patron_record_id = rm.id

        WHERE f.assessed_gmt::DATE = CURRENT_DATE
            AND f.charge_code = '1'
            AND f.description = 'Residual fine' 
        """

    # identify patrons with amt owed errors and create manual charges in the amount of those discrepancies
    amt_owed_errors = run_query(error_query)
    for rownum, row in enumerate(amt_owed_errors):
        manual_charge(str(row[0]), row[1], row[2])

    # Find the newly created manual charges and waive them
    fines_to_clear = run_query(manual_charge_query)
    for rownum, row in enumerate(fines_to_clear):
        clear_fine(str(row[0]), row[1])


main()
