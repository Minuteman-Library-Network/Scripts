#!/usr/bin/env python3
# run in py313

"""
Jeremy Goldstein
Minuteman Library Network

Create monthly collection dev by scat report for each library
and sftp file to eachc library's reports folder on our intranet site
"""

import psycopg2
import xlsxwriter
import os
import pysftp
import configparser
import sys
import time
from datetime import date, timedelta
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email.utils import formatdate
from email import encoders
import traceback


# run sql query against Sierra database and return results
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


# convert sql query results into formatted excel file
def excel_writer(query_results, excelfile):
    workbook = xlsxwriter.Workbook(excelfile, {"remove_timezone": True})
    worksheet = workbook.add_worksheet()

    # Formatting our Excel worksheet
    worksheet.set_landscape()
    worksheet.hide_gridlines(0)

    # Formatting Cells
    eformat = workbook.add_format(
        {"text_wrap": True, "valign": "top", "align": "center"}
    )
    eformatlabel = workbook.add_format(
        {"text_wrap": True, "valign": "top", "bold": True, "align": "center"}
    )
    eformatprice = workbook.add_format(
        {"text_wrap": True, "valign": "top", "align": "center"}
    )
    eformatpercent = workbook.add_format(
        {"text_wrap": True, "valign": "top", "align": "center"}
    )

    eformatprice.set_num_format(0x07)
    eformatpercent.set_num_format(0x0A)

    # Setting the column widths
    worksheet.set_column(0, 0, 5.29)
    worksheet.set_column(1, 1, 10.29)
    worksheet.set_column(2, 2, 10.14)
    worksheet.set_column(3, 3, 10.14)
    worksheet.set_column(4, 4, 10.14)
    worksheet.set_column(5, 5, 8.57)
    worksheet.set_column(6, 6, 10.86)
    worksheet.set_column(7, 7, 8.57)
    worksheet.set_column(8, 8, 10.86)
    worksheet.set_column(9, 9, 8.57)
    worksheet.set_column(10, 10, 10.86)
    worksheet.set_column(11, 11, 8.57)
    worksheet.set_column(12, 12, 10.86)
    worksheet.set_column(13, 13, 8.57)
    worksheet.set_column(14, 14, 10.43)
    worksheet.set_column(15, 15, 11)
    worksheet.set_column(16, 16, 8.71)
    worksheet.set_column(17, 17, 8.86)
    worksheet.set_column(18, 18, 10)
    worksheet.set_column(19, 19, 10)

    # Inserting a header
    worksheet.set_header("Collection Development By Scat")

    # Adding column labels
    worksheet.write(0, 0, "Scat", eformatlabel)
    worksheet.write(0, 1, "Item Total", eformatlabel)
    worksheet.write(0, 2, "Checkout Total", eformatlabel)
    worksheet.write(0, 3, "Renewal Total", eformatlabel)
    worksheet.write(0, 4, "Circulation Total", eformatlabel)
    worksheet.write(0, 5, "AVG Price", eformatlabel)
    worksheet.write(0, 6, "Have Circed Within 1 Year", eformatlabel)
    worksheet.write(0, 7, "% Circed Within 1 Year", eformatlabel)
    worksheet.write(0, 8, "Have Circed Within 3 Years", eformatlabel)
    worksheet.write(0, 9, "% Circed Within 3 Years", eformatlabel)
    worksheet.write(0, 10, "Have Circed Within 5 Years", eformatlabel)
    worksheet.write(0, 11, "% Circed Within 5 Years", eformatlabel)
    worksheet.write(0, 12, "Have Circed", eformatlabel)
    worksheet.write(0, 13, "% Have Circed", eformatlabel)
    worksheet.write(0, 14, "Have 0 Circs", eformatlabel)
    worksheet.write(0, 15, "% 0 Circs", eformatlabel)
    worksheet.write(0, 16, "Cost Per Circ", eformatlabel)
    worksheet.write(0, 17, "Turnover", eformatlabel)
    worksheet.write(0, 18, "Relative Item Total", eformatlabel)
    worksheet.write(0, 19, "Relative Circulation", eformatlabel)

    # Writing the report for staff to the Excel worksheet
    for rownum, row in enumerate(query_results):
        worksheet.write(rownum + 1, 0, row[0], eformat)
        worksheet.write(rownum + 1, 1, row[1], eformat)
        worksheet.write(rownum + 1, 2, row[2], eformat)
        worksheet.write(rownum + 1, 3, row[3], eformat)
        worksheet.write(rownum + 1, 4, row[4], eformat)
        worksheet.write(rownum + 1, 5, row[5], eformatprice)
        worksheet.write(rownum + 1, 6, row[6], eformat)
        worksheet.write(rownum + 1, 7, row[7], eformatpercent)
        worksheet.write(rownum + 1, 8, row[8], eformat)
        worksheet.write(rownum + 1, 9, row[9], eformatpercent)
        worksheet.write(rownum + 1, 10, row[10], eformat)
        worksheet.write(rownum + 1, 11, row[11], eformatpercent)
        worksheet.write(rownum + 1, 12, row[12], eformat)
        worksheet.write(rownum + 1, 13, row[13], eformatpercent)
        worksheet.write(rownum + 1, 14, row[14], eformat)
        worksheet.write(rownum + 1, 15, row[15], eformatpercent)
        worksheet.write(rownum + 1, 16, row[16], eformatprice)
        worksheet.write(rownum + 1, 17, row[17], eformat)
        worksheet.write(rownum + 1, 18, row[18], eformatpercent)
        worksheet.write(rownum + 1, 19, row[19], eformatpercent)

    workbook.close()

    return excelfile


# upload report to SIC directory and optionally remove older files
def sftp_file(local_file, library):

    config = configparser.ConfigParser()
    config.read("C:\\Scripts\\Creds\\config.ini")

    cnopts = pysftp.CnOpts()

    srv = pysftp.Connection(
        host=config["sic"]["sic_host"],
        username=config["sic"]["sic_user"],
        password=config["sic"]["sic_pw"],
        cnopts=cnopts,
    )

    local_file = local_file

    srv.cwd(
        "/reports/Library-Specific Reports/"
        + library
        + "/Collection Development By Scat/"
    )
    srv.put(local_file)

    # remove old file

    for fname in srv.listdir_attr():
        fullpath = (
            "/reports/Library-Specific Reports/"
            + library
            + "/Collection Development By Scat/{}".format(fname.filename)
        )
        # time tracked in seconds, st_mtime is time last modified
        name = str(fname.filename)
        if (name != "meta.json") and (
            (time.time() - fname.st_mtime) // (24 * 3600) >= 1095
        ):
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


def main(library, libcode):

    try:
        query = (
            r"""
        SELECT
          icode1 AS "Scat",
          COUNT (id) AS "Item total",
          SUM(checkout_total) AS "Total_Checkouts",
          SUM(renewal_total) AS "Total_Renewals",
          SUM(checkout_total) + SUM(renewal_total) AS "Total_Circulation",
          ROUND(AVG(price) FILTER(WHERE price>'0' AND price <'10000'),2) AS "AVG_price",
          COUNT (id) FILTER(WHERE last_checkout_gmt >= (CURRENT_DATE - INTERVAL '1 year')) AS "have_circed_within_1_year",
          ROUND(CAST(COUNT(id) FILTER(WHERE last_checkout_gmt >= (CURRENT_DATE - INTERVAL '1 year')) AS NUMERIC (12,2)) / CAST(COUNT (id) AS NUMERIC (12,2)), 6) AS "Percentage_1_year",
          COUNT (id) FILTER(WHERE last_checkout_gmt >= (CURRENT_DATE - INTERVAL '3 years')) AS "have_circed_within_3_years",
          ROUND(CAST(COUNT(id) FILTER(WHERE last_checkout_gmt >= (CURRENT_DATE - INTERVAL '3 years')) AS NUMERIC (12,2)) / CAST(COUNT (id) AS NUMERIC (12,2)), 6) AS "Percentage_3_years",
          COUNT (id) FILTER(WHERE last_checkout_gmt >= (CURRENT_DATE - INTERVAL '5 years')) AS "have_circed_within_5_years",
          ROUND(CAST(COUNT(id) FILTER(WHERE last_checkout_gmt >= (CURRENT_DATE - INTERVAL '5 years')) AS NUMERIC (12,2)) / CAST(COUNT (id) AS NUMERIC (12,2)), 6) AS "Percentage_5_years",
          COUNT (id) FILTER(WHERE last_checkout_gmt IS NOT NULL) AS "have_circed_within_5+_years",
          ROUND(CAST(COUNT(id) FILTER(WHERE last_checkout_gmt IS NOT NULL) AS NUMERIC (12,2)) / CAST(COUNT (id) AS NUMERIC (12,2)), 6) AS "Percentage_5+_years",
          COUNT (id) FILTER(WHERE last_checkout_gmt IS NULL) AS "0_circs",
          ROUND(CAST(COUNT(id) FILTER(WHERE last_checkout_gmt IS NULL) AS NUMERIC (12,2)) / CAST(COUNT (id) AS NUMERIC (12,2)), 6) AS "Percentage_0_circs",
          ROUND((COUNT(id) *(AVG(price) FILTER(WHERE price>'0' AND price <'10000'))/(NULLIF((SUM(checkout_total) + SUM(renewal_total)),0))),2) AS "Cost_Per_Circ_By_AVG_price",
          ROUND(CAST(SUM(checkout_total) + SUM(renewal_total) AS NUMERIC (12,2))/CAST(COUNT (id) AS NUMERIC (12,2)), 2) AS turnover,
          ROUND(CAST(count(id) AS NUMERIC (12,2)) / (SELECT CAST(COUNT (id) AS NUMERIC (12,2))FROM sierra_view.item_record WHERE location_code ~ '^"""
            + libcode.lower()
            + r"""' AND item_status_code NOT IN ('o', 'n', '$', 'w', 'z', 'd')), 6) AS relative_item_total,
          ROUND(CAST(SUM(checkout_total) + SUM(renewal_total) AS NUMERIC (12,2)) / (SELECT CAST(SUM(checkout_total) + SUM(renewal_total) AS NUMERIC (12,2)) FROM sierra_view.item_record WHERE location_code ~ '^"""
            + libcode.lower()
            + r"""' AND item_status_code NOT IN ('o', 'n', '$', 'w', 'z', 'd')), 6) AS relative_circ

        FROM sierra_view.item_record
        WHERE location_code ~ '^"""
            + libcode.lower()
            + r"""'
          AND item_status_code NOT IN ('o', 'n', '$', 'w', 'z', 'd')
        GROUP BY 1
        ORDER BY 1
        """
        )

        query_results = run_query(query)
        # Name of Excel File
        excel_file = (
            "/Scripts/Collection Dev By Scat/Temp Files/"
            + libcode
            + "CollectionDevByScat{}.xlsx".format(date.today())
        )
        excel_writer(query_results, excel_file)
        sftp_file(
            "C:\\Scripts\\Collection Dev By Scat\\Temp Files\\"
            + libcode
            + "CollectionDevByScat{}.xlsx".format(date.today()),
            library,
        )

    except:
        # read config file with recipient list for email
        config_recipient = configparser.ConfigParser()
        config_recipient.read("C:\\Scripts\\Creds\\emails.ini")
        emailto = config_recipient["script_error"]["recipients"].split()

        # craft email subject and message containing error message details from traceback
        email_subject = "collection dev by scat " + library + " script error"
        email_message = (
            "Your script failed with the following error:\n\n" + traceback.format_exc()
        )

        send_email_error(email_subject, email_message, emailto)
        raise


if __name__ == "__main__":
    # run for each library within Minuteman
    main("Acton", "ACT")
    main("Acton", "AC2")
    main("Arlington", "ARL")
    main("Arlington", "AR2")
    main("Ashland", "ASH")
    main("Bedford", "BED")
    main("Belmont", "BLM")
    main("Brookline", "BRK")
    main("Brookline", "BR2")
    main("Brookline", "BR3")
    main("Cambridge", "CAM")
    main("Cambridge", "CA3")
    main("Cambridge", "CA4")
    main("Cambridge", "CA5")
    main("Cambridge", "CA6")
    main("Cambridge", "CA7")
    main("Cambridge", "CA8")
    main("Cambridge", "CA9")
    main("Concord", "CON")
    main("Concord", "CO2")
    main("Dedham", "DDM")
    main("Dedham", "DD2")
    main("Dean", "DEA")
    main("Dover", "DOV")
    main("Framingham Public", "FPL")
    main("Framingham Public", "FP2")
    main("Framingham State", "FST")
    main("Franklin", "FRK")
    main("Holliston", "HOL")
    main("Lasell", "LAS")
    main("Lexington", "LEX")
    main("Lincoln", "LIN")
    main("Maynard", "MAY")
    main("Medfield", "MLD")
    main("Medford", "MED")
    main("Medway", "MWY")
    main("Millis", "MIL")
    main("Natick", "NAT")
    main("Natick", "NA2")
    main("Needham", "NEE")
    main("Newton", "NTN")
    main("Norwood", "NOR")
    main("Olin", "OLN")
    main("Regis", "REG")
    main("Sherborn", "SHR")
    main("Somerville", "SOM")
    main("Somerville", "SO2")
    main("Somerville", "SO3")
    main("Stow", "STO")
    main("Sudbury", "SUD")
    main("Waltham", "WLM")
    main("Waltham", "WL2")
    main("Watertown", "WAT")
    main("Watertown", "WA4")
    main("Wayland", "WYL")
    main("Wellesley", "WEL")
    main("Wellesley", "WE2")
    main("Wellesley", "WE3")
    main("Weston", "WSN")
    main("Westwood", "WWD")
    main("Westwood", "WW2")
    main("Winchester", "WIN")
    main("Woburn", "WOB")
