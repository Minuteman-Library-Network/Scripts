#!/usr/bin/env python3

"""
Jeremy Goldstein
Minuteman Library Network

Generates ARIS holdings count report for annual statewide reporting
"""
# run in py38

import psycopg2
import xlsxwriter
import configparser
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email.utils import formatdate
from email import encoders
from datetime import date


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

    #Creating the Excel file for staff
    workbook = xlsxwriter.Workbook(excel_file)
    worksheet = workbook.add_worksheet()


    #Formatting our Excel worksheet
    worksheet.set_landscape()
    worksheet.hide_gridlines(0)

    #Formatting Cells
    eformat= workbook.add_format({'text_wrap': True, 'valign': 'top'})
    eformatlabel= workbook.add_format({'text_wrap': True, 'valign': 'top', 'bold': True})


    # Setting the column widths
    worksheet.set_column(0,0,16.43)
    worksheet.set_column(1,1,16.43)
    worksheet.set_column(2,2,16.43)
    worksheet.set_column(3,3,16.43)

    #Inserting a header
    worksheet.set_header('ARIS Holdings count')

    # Adding column labels
    worksheet.write(0,0,'ARIS_Category', eformatlabel)
    worksheet.write(0,1,'Age_Lvl', eformatlabel)
    worksheet.write(0,2,'Title_Count', eformatlabel)
    worksheet.write(0,3,'Item_Count', eformatlabel)

    # Writing the report for staff to the Excel worksheet
    for rownum, row in enumerate(query_results):
        worksheet.write(rownum+1,0,row[0], eformat)
        worksheet.write(rownum+1,1,row[1], eformat)
        worksheet.write(rownum+1,2,row[2], eformat)
        worksheet.write(rownum+1,3,row[3], eformat)
    
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
    emailto = config_recipient["annual_reports"]["recipients"].split()
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


def main():
    # query to identify patron records with incorrect owed_amt fields
    query = """
           SELECT
	         CASE 
	           WHEN b.material_code='2' OR b.material_code='9' OR b.material_code='a' OR b.material_code='c' OR b.material_code='t' OR b.material_code='f' THEN 'Books'
	           WHEN b.material_code='3' THEN 'Periodicals'
	           WHEN b.material_code='4' OR b.material_code='7' OR b.material_code='8' OR b.material_code='i' OR b.material_code='j' OR b.material_code='z' THEN 'Audio'
	           WHEN b.material_code='5' OR b.material_code='g' OR b.material_code='u' OR b.material_code='x' THEN 'Video'
	           WHEN b.material_code='h' THEN 'E-book'
	           WHEN b.material_code='s' or b.material_code='w' THEN 'Downloadable audio'
	           WHEN b.material_code='l' THEN 'Downloadable video'
	           WHEN b.material_code='m' OR b.material_code='n' THEN 'Materials in Electronic Format'
	           WHEN b.material_code='6' OR b.material_code='b' OR b.material_code='e' OR b.material_code='k' OR b.material_code='o' OR b.material_code='p' OR b.material_code='q' or b.material_code='r' OR b.material_code='v' THEN 'Miscellaneous'
	           WHEN b.material_code='y' THEN 'Electronic collections' 
	           ELSE 'Unknown'
	         END AS "ARIS CATEGORY",
	         CASE
	           WHEN SUBSTRING(i.location_code,4,1)='j' THEN 'Juv'
	           WHEN SUBSTRING(i.location_code,4,1)='y' THEN 'YA'
	           Else 'Adult'
	         END AS "Age level",
	         COUNT(DISTINCT b.id) AS "title count",
	         COUNT(i.id) AS "item count"
           
           FROM sierra_view.item_record AS i
	       JOIN sierra_view.bib_record_item_record_link	AS bi
	         ON i.record_id=bi.item_record_id
             AND i.itype_code_num != '80'
	       JOIN sierra_view.bib_record_property AS b
	         ON bi.bib_record_id=b.bib_record_id
 
           GROUP BY 1,2
           ORDER BY 1
           """
    query_results = run_query(query)

    # generate excel file from those query results
    excel_file = "/Scripts/Annual Reports/Archive/ArisHoldingsCount{}.xlsx".format(
        date.today()
    )
    excel_writer(query_results, excel_file)

    # send email with attached file
    email_subject = "ARIS Holdings Count"
    email_message = """***This is an automated email***


The ARIS Holdings Count report has been attached."""
    send_email(email_subject, email_message, excel_file)


main()
