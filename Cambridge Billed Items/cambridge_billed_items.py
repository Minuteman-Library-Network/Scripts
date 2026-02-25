#!/usr/bin/env python3

"""
Jeremy Goldstein
Minuteman Library Network

Query gathers report of billed childrens items belonging to Cambridge
and sends excel file of results to designated staff at the library
"""

# run in py38

import psycopg2
import xlsxwriter
import smtplib
import os
import configparser
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email.utils import formatdate
from email import encoders
from datetime import date
import traceback


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
    # return variables containing query results and column headers
    return rows


# convert sql query results into formatted excel file
def excel_writer(query_results, excel_file):

    # Creating the Excel file for staff
    workbook = xlsxwriter.Workbook(excel_file)
    worksheet = workbook.add_worksheet()

    # Formatting our Excel worksheet
    worksheet.set_landscape()
    worksheet.hide_gridlines(0)

    # Formatting Cells
    eformat = workbook.add_format({"text_wrap": True, "valign": "top"})
    eformatlabel = workbook.add_format(
        {"text_wrap": True, "valign": "top", "bold": True}
    )
    eformat2 = workbook.add_format({"num_format": "mm/dd/yy"})

    # Setting the column widths
    worksheet.set_column(0, 0, 13.4)
    worksheet.set_column(1, 1, 10.3)
    worksheet.set_column(2, 2, 15.3)
    worksheet.set_column(3, 3, 78.7)
    worksheet.set_column(4, 4, 11.6)
    worksheet.set_column(5, 5, 17)
    worksheet.set_column(6, 6, 29.15)

    # Inserting a header
    worksheet.set_header("Cambridge Monthly Bills")

    # Adding column labels
    worksheet.write(0, 0, "Assessed_date", eformatlabel)
    worksheet.write(0, 1, "Pnumber", eformatlabel)
    worksheet.write(0, 2, "Barcode", eformatlabel)
    worksheet.write(0, 3, "Title", eformatlabel)
    worksheet.write(0, 4, "Charge Amt", eformatlabel)
    worksheet.write(0, 5, "Checkout Loc", eformatlabel)
    worksheet.write(0, 6, "eMail", eformatlabel)

    # Writing the report for staff to the Excel worksheet
    for rownum, row in enumerate(query_results):
        worksheet.write(rownum + 1, 0, row[0], eformat2)
        worksheet.write(rownum + 1, 1, row[1], eformat)
        worksheet.write(rownum + 1, 2, row[2], eformat)
        worksheet.write(rownum + 1, 3, row[3], eformat)
        worksheet.write(rownum + 1, 4, row[4], eformat)
        worksheet.write(rownum + 1, 5, row[5], eformat)
        worksheet.write(rownum + 1, 6, row[6], eformat)

    workbook.close()


# function takes a file as a parameter and attaches that file to an outgoing email
def send_email(subject, message, attachment):
    # read config file with credentials for email account
    config = configparser.ConfigParser()
    config.read("C:\\Scripts\\Creds\\config.ini")
    # read config file with recipient list for email
    config_recipient = configparser.ConfigParser()
    config_recipient.read("C:\\Scripts\\Creds\\emails.ini")

    # These are variables for the email that will be sent, taken from .ini files referenced above
    emailhost = config["email"]["host"]
    emailuser = config["email"]["user"]
    emailpass = config["email"]["pw"]
    emailport = config["email"]["port"]
    emailfrom = config["email"]["sender"]
    emailto = config_recipient["cambridge_billed_items"]["recipients"].split()
    # plain text of email message
    emailmessage = message

    # Creating the email message
    msg = MIMEMultipart()
    msg["From"] = emailfrom
    if type(emailto) is list:
        msg["To"] = ", ".join(emailto)
    else:
        msg["To"] = emailto
    msg["Date"] = formatdate(localtime=True)
    msg["Subject"] = subject
    msg.attach(MIMEText(emailmessage))
    part = MIMEBase("application", "octet-stream")
    part.set_payload(open(attachment, "rb").read())
    encoders.encode_base64(part)
    part.add_header(
        "Content-Disposition", "attachment; filename=%s" % attachment.rsplit("/", 1)[-1]
    )
    msg.attach(part)

    # Sending the email message
    smtp = smtplib.SMTP(emailhost, emailport)
    smtp.ehlo()
    smtp.starttls()
    smtp.login(emailuser, emailpass)
    smtp.sendmail(emailfrom, emailto, msg.as_string())
    smtp.quit()


# function constructs and sends outgoing email given a subject, a recipient and body text in both txt and html forms
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


def main():
    # query to identify bills from the past week for childrens items belonging to Cambridge
    query = """
           SELECT
               f.assessed_gmt::DATE AS assessed_date,
               rm.record_type_code||rm.record_num||'a' AS pnumber,
               ip.barcode,
               f.title,
               f.item_charge_amt::MONEY AS charge_amt,
               CASE 
	               WHEN (f.loanrule_code_num BETWEEN 2 AND 12 OR f.loanrule_code_num BETWEEN 501 AND 509) THEN 'Acton'
	               WHEN (f.loanrule_code_num BETWEEN 13 AND 23 OR f.loanrule_code_num BETWEEN 510 AND 518) THEN 'Arlington'
	               WHEN (f.loanrule_code_num BETWEEN 24 AND 34 OR f.loanrule_code_num BETWEEN 519 AND 527) THEN 'Ashland'
	               WHEN (f.loanrule_code_num BETWEEN 35 AND 45 OR f.loanrule_code_num BETWEEN 528 AND 536) THEN 'Bedford'
	               WHEN (f.loanrule_code_num BETWEEN 46 AND 56 OR f.loanrule_code_num BETWEEN 537 AND 545) THEN 'Belmont'
	               WHEN (f.loanrule_code_num BETWEEN 57 AND 67 OR f.loanrule_code_num BETWEEN 546 AND 554) THEN 'Brookline'
	               WHEN (f.loanrule_code_num BETWEEN 68 AND 78 OR f.loanrule_code_num BETWEEN 555 AND 563) THEN 'Cambridge'
	               WHEN (f.loanrule_code_num BETWEEN 79 AND 89 OR f.loanrule_code_num BETWEEN 564 AND 572) THEN 'Concord'
	               WHEN (f.loanrule_code_num BETWEEN 90 AND 100 OR f.loanrule_code_num BETWEEN 573 AND 581) THEN 'Dedham'
	               WHEN (f.loanrule_code_num BETWEEN 101 AND 111 OR f.loanrule_code_num BETWEEN 582 AND 590) THEN 'Dean'
	               WHEN (f.loanrule_code_num BETWEEN 112 AND 122 OR f.loanrule_code_num BETWEEN 591 AND 599) THEN 'Dover'
	               WHEN (f.loanrule_code_num BETWEEN 123 AND 133 OR f.loanrule_code_num BETWEEN 600 AND 608) THEN 'Framingham'
	               WHEN (f.loanrule_code_num BETWEEN 134 AND 144 OR f.loanrule_code_num BETWEEN 609 AND 617) THEN 'Franklin'
	               WHEN (f.loanrule_code_num BETWEEN 145 AND 155 OR f.loanrule_code_num BETWEEN 618 AND 626) THEN 'Framingham State'
	               WHEN (f.loanrule_code_num BETWEEN 156 AND 166 OR f.loanrule_code_num BETWEEN 627 AND 635) THEN 'Holliston'
	               WHEN (f.loanrule_code_num BETWEEN 167 AND 177 OR f.loanrule_code_num BETWEEN 636 AND 644) THEN 'Lasell'
	               WHEN (f.loanrule_code_num BETWEEN 178 AND 188 OR f.loanrule_code_num BETWEEN 645 AND 653) THEN 'Lexington' 
	               WHEN (f.loanrule_code_num BETWEEN 189 AND 199 OR f.loanrule_code_num BETWEEN 654 AND 662) THEN 'Lincoln'
	               WHEN (f.loanrule_code_num BETWEEN 200 AND 210 OR f.loanrule_code_num BETWEEN 663 AND 671) THEN 'Maynard'
	               WHEN (f.loanrule_code_num BETWEEN 222 AND 232 OR f.loanrule_code_num BETWEEN 681 AND 689) THEN 'Medford'
	               WHEN (f.loanrule_code_num BETWEEN 233 AND 243 OR f.loanrule_code_num BETWEEN 690 AND 698) THEN 'Millis' 
	               WHEN (f.loanrule_code_num BETWEEN 244 AND 254 OR f.loanrule_code_num BETWEEN 699 AND 707) THEN 'Medfield'
	               WHEN (f.loanrule_code_num BETWEEN 266 AND 276 OR f.loanrule_code_num BETWEEN 717 AND 725) THEN 'Medway' 
	               WHEN (f.loanrule_code_num BETWEEN 277 AND 287 OR f.loanrule_code_num BETWEEN 726 AND 734) THEN 'Natick'
	               WHEN (f.loanrule_code_num BETWEEN 299 AND 309 OR f.loanrule_code_num BETWEEN 744 AND 752) THEN 'Needham' 
	               WHEN (f.loanrule_code_num BETWEEN 310 AND 320 OR f.loanrule_code_num BETWEEN 753 AND 761) THEN 'Norwood' 
	               WHEN (f.loanrule_code_num BETWEEN 321 AND 331 OR f.loanrule_code_num BETWEEN 762 AND 770) THEN 'Newton' 
	               WHEN (f.loanrule_code_num BETWEEN 289 AND 298 OR f.loanrule_code_num BETWEEN 734 AND 743) THEN 'Olin'
	               WHEN (f.loanrule_code_num BETWEEN 332 AND 342 OR f.loanrule_code_num BETWEEN 771 AND 779) THEN 'Somerville' 
	               WHEN (f.loanrule_code_num BETWEEN 343 AND 353 OR f.loanrule_code_num BETWEEN 780 AND 788) THEN 'Stow' 
	               WHEN (f.loanrule_code_num BETWEEN 354 AND 364 OR f.loanrule_code_num BETWEEN 789 AND 797) THEN 'Sudbury'
	               WHEN (f.loanrule_code_num BETWEEN 365 AND 375 OR f.loanrule_code_num BETWEEN 798 AND 806) THEN 'Watertown' 
	               WHEN (f.loanrule_code_num BETWEEN 376 AND 386 OR f.loanrule_code_num BETWEEN 807 AND 815) THEN 'Wellesley' 
	               WHEN (f.loanrule_code_num BETWEEN 387 AND 397 OR f.loanrule_code_num BETWEEN 816 AND 824) THEN 'Winchester' 
	               WHEN (f.loanrule_code_num BETWEEN 398 AND 408 OR f.loanrule_code_num BETWEEN 825 AND 833) THEN 'Waltham' 
	               WHEN (f.loanrule_code_num BETWEEN 409 AND 419 OR f.loanrule_code_num BETWEEN 834 AND 842) THEN 'Woburn'
	               WHEN (f.loanrule_code_num BETWEEN 420 AND 430 OR f.loanrule_code_num BETWEEN 843 AND 851) THEN 'Weston' 
	               WHEN (f.loanrule_code_num BETWEEN 431 AND 441 OR f.loanrule_code_num BETWEEN 852 AND 860) THEN 'Westwood' 
	               WHEN (f.loanrule_code_num BETWEEN 442 AND 452 OR f.loanrule_code_num BETWEEN 861 AND 869) THEN 'Wayland' 
	               WHEN (f.loanrule_code_num BETWEEN 453 AND 463 OR f.loanrule_code_num BETWEEN 870 AND 878) THEN 'Pine Manor' 
	               WHEN (f.loanrule_code_num BETWEEN 464 AND 474 OR f.loanrule_code_num BETWEEN 879 AND 887) THEN 'Regis' 
	               WHEN (f.loanrule_code_num BETWEEN 475 AND 485 OR f.loanrule_code_num BETWEEN 888 AND 896) THEN 'Sherborn' 
	               ELSE 'Other'
               END AS checkout_location,
               COALESCE(email.field_content,'') AS email

           FROM sierra_view.fine f
           JOIN sierra_view.item_record i
             ON f.item_record_metadata_id = i.id
           JOIN sierra_view.record_metadata rm
             ON f.patron_record_id = rm.id
           JOIN sierra_view.item_record_property ip
             ON i.id = ip.item_record_id
           LEFT JOIN sierra_view.varfield email
             ON f.patron_record_id = email.record_id
             AND email.varfield_type_code = 'z'

           WHERE f.charge_code IN ('3','5')
             AND f.assessed_gmt::DATE >= CURRENT_DATE - INTERVAL '1 week'
             AND i.location_code ~ '^ca\w{1}(j|y)'
           ORDER BY 2,1

           """
    query_results = run_query(query)

    # generate excel file from those query results
    excel_file = (
        "/Scripts/Cambridge Billed Items/Temp Files/CAMBilledItems{}.xlsx".format(
            date.today()
        )
    )
    excel_writer(query_results, excel_file)

    # send email
    email_subject = "Cambridge Billed Items"
    email_message = """***This is an automated email***


The Cambridge Billed items report has been attached."""
    send_email(email_subject, email_message, excel_file)

    # delete local file
    os.remove(excel_file)


# run main function and send error email to admin of script encounters an error
if __name__ == "__main__":
    try:
        main()
    except Exception:
        # read config file with recipient list for email
        config_recipient = configparser.ConfigParser()
        config_recipient.read("C:\\Scripts\\Creds\\emails.ini")
        emailto = config_recipient["script_error"]["recipients"].split()

        # craft email subject and message containing error message details from traceback
        email_subject = "annual reports: record count script error"
        email_message = (
            "Your script failed with the following error:\n\n" + traceback.format_exc()
        )

        send_email_error(email_subject, email_message, emailto)
        raise
