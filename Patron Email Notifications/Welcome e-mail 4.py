#!/usr/bin/env python3

# Run in py38

"""
Jeremy Goldstein
Minuteman Library Network
Generate and send email notification to patrons with a new library card
"""

import psycopg2
import smtplib
import configparser
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email.utils import formatdate
from email import encoders
from datetime import date


def run_query(query):
    # read config file with Sierra login credentials
    config = configparser.ConfigParser()
    config.read("C:\\Scripts\\Creds\\config.ini")

    # Connecting to Sierra PostgreSQL database
    try:
        conn = psycopg2.connect(config["sql"]["connection_string"])
    except psycopg2.Error as e:
        print("Unable to connect to database: " + str(e))

    # Opening a session and querying the database
    cursor = conn.cursor()
    cursor.execute(query)
    # For now, just storing the data in a variable. We'll use it later.
    rows = cursor.fetchall()
    conn.close()
    return rows


# function constructs and sends outgoing email given a subject, a recipient and body text in both txt and html forms
def send_email(subject, message_text, message_html, recipient):
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

    # Creating the email message with html and plaintxt options
    msg = MIMEMultipart("alternative")
    part1 = MIMEText(message_text, "plain")
    part2 = MIMEText(message_html, "html")
    msg["From"] = emailfrom
    if type(recipient) is list:
        msg["To"] = ", ".join(recipient)
    else:
        msg["To"] = recipient
    msg["Date"] = formatdate(localtime=True)
    msg["Subject"] = subject
    msg.attach(part1)
    msg.attach(part2)

    # Sending the email message
    smtp = smtplib.SMTP(emailhost, emailport)
    # for Gmail connection used within Minuteman
    smtp.ehlo()
    smtp.starttls()
    smtp.login(emailuser, emailpass)
    smtp.sendmail(emailfrom, recipient, msg.as_string())
    smtp.quit()


def main():
    query = """
      --Find patrons in the fourth quarter of MLN libraries who got a library card yesterday
      SELECT
        MIN(n.first_name),
        MIN(n.last_name),
        MIN(v.field_content) AS email,
        p.barcode,
        CASE
          WHEN p.ptype_code = 5 THEN 'the Belmont Public Library'
          WHEN p.ptype_code = 6 THEN 'the Public Library of Brookline'
          WHEN p.ptype_code = 7 THEN 'the Cambridge Public Library'
          WHEN p.ptype_code = 8 THEN 'the Concord Free Public Library'
          WHEN p.ptype_code IN ('10','110') THEN 'the Dedham Public Library'
          WHEN p.ptype_code IN ('17', '117') THEN 'the Cary Memorial Library'
          WHEN p.ptype_code IN ('29', '129') THEN 'the Newton Free Library'
          WHEN p.ptype_code = 31 THEN 'the Somerville Public Library'
          WHEN p.ptype_code = 35 THEN 'the Watertown Free Public Library'
          WHEN p.ptype_code IN ('37', '137') THEN 'the Wellesley Free Library'
          WHEN p.ptype_code = 1 THEN 'the Acton Public Library'
          WHEN p.ptype_code = 2 THEN 'the Robbins Library'
          WHEN p.ptype_code = 3 THEN 'the Ashland Public Library'
          WHEN p.ptype_code = 4 THEN 'the Bedford Free Public Library'
          WHEN p.ptype_code = 11 THEN 'the Dover Town Library'
          WHEN p.ptype_code = 12 THEN 'the Framingham Public Library'
          WHEN p.ptype_code = 14 THEN 'the Franklin Public Library'
          WHEN p.ptype_code IN ('15', '115') THEN 'the Holliston Public Library'
          WHEN p.ptype_code = 18 THEN 'the Lincoln Public Library'
          WHEN p.ptype_code IN ('20', '120') THEN 'the Maynard Public Library'
          WHEN p.ptype_code IN ('21', '121') THEN 'the Medfield Public Library'
          WHEN p.ptype_code IN ('22', '122') THEN 'the Medford Public Library'
          WHEN p.ptype_code = 23 THEN 'the Medway Public Library'
          WHEN p.ptype_code = 24 THEN 'the Millis Public Library'
          WHEN (p.ptype_code = 26 AND p.home_library_code != 'na2z') THEN 'the Morse Institute Library'
          WHEN (p.ptype_code = 26 AND p.home_library_code = 'na2z') THEN 'the Bacon Free Library'
          WHEN p.ptype_code = 27 THEN 'the Needham Free Public Library'
          WHEN p.ptype_code IN ('30', '130') THEN 'the Morrill Memorial Library'
          WHEN p.ptype_code = 32 THEN 'the Randall Library'
          WHEN p.ptype_code IN ('33', '133') THEN 'the Goodnow Library'
          WHEN p.ptype_code = 34 THEN 'the Waltham Public Library'
          WHEN p.ptype_code = 36 THEN 'the Wayland Free Library'
          WHEN p.ptype_code = 38 THEN 'the Weston Public Library'
          WHEN p.ptype_code = 39 THEN 'the Westwood Public Library'
          WHEN p.ptype_code = 40 THEN 'the Winchester Public Library'
          WHEN p.ptype_code = 41 THEN 'the Woburn Public Library'
          WHEN p.ptype_code = 46 THEN 'the Sherborn Public Library'
          ELSE 'the Minuteman Library Network'
        END AS library,
        CASE
          WHEN p.ptype_code = 5 THEN 'the Belmont Public Library</a>, a member of the Minuteman Library Network'
          WHEN p.ptype_code = 6 THEN 'the Public Library of Brookline</a>, a member of the Minuteman Library Network'
          WHEN p.ptype_code = 7 THEN 'the Cambridge Public Library</a>, a member of the Minuteman Library Network'
          WHEN p.ptype_code = 8 THEN 'the Concord Free Public Library</a>, a member of the Minuteman Library Network'
          WHEN p.ptype_code IN ('10','110') THEN 'the Dedham Public Library</a>, a member of the Minuteman Library Network'
          WHEN p.ptype_code IN ('17', '117') THEN 'the Cary Memorial Library</a>, a member of the Minuteman Library Network'
          WHEN p.ptype_code IN ('29', '129') THEN 'the Newton Free Library</a>, a member of the Minuteman Library Network'
          WHEN p.ptype_code = 31 THEN 'the Somerville Public Library</a>, a member of the Minuteman Library Network'
          WHEN p.ptype_code = 35 THEN 'the Watertown Free Public Library</a>, a member of the Minuteman Library Network'
          WHEN p.ptype_code IN ('37', '137') THEN 'the Wellesley Free Library</a>, a member of the Minuteman Library Network'
          WHEN p.ptype_code = 1 THEN 'the Acton Public Library</a>, a member of the Minuteman Library Network'
          WHEN p.ptype_code = 2 THEN 'the Robbins Library</a>, a member of the Minuteman Library Network'
          WHEN p.ptype_code = 3 THEN 'the Ashland Public Library</a>, a member of the Minuteman Library Network'
          WHEN p.ptype_code = 4 THEN 'the Bedford Free Public Library</a>, a member of the Minuteman Library Network'
          WHEN p.ptype_code = 11 THEN 'the Dover Town Library</a>, a member of the Minuteman Library Network'
          WHEN p.ptype_code = 12 THEN 'the Framingham Public Library</a>, a member of the Minuteman Library Network'
          WHEN p.ptype_code = 14 THEN 'the Franklin Public Library</a>, a member of the Minuteman Library Network'
          WHEN p.ptype_code IN ('15', '115') THEN 'the Holliston Public Library</a>, a member of the Minuteman Library Network'
          WHEN p.ptype_code = 18 THEN 'the Lincoln Public Library</a>, a member of the Minuteman Library Network'
          WHEN p.ptype_code IN ('20', '120') THEN 'the Maynard Public Library</a>, a member of the Minuteman Library Network'
          WHEN p.ptype_code IN ('21', '121') THEN 'the Medfield Public Library</a>, a member of the Minuteman Library Network'
          WHEN p.ptype_code IN ('22', '122') THEN 'the Medford Public Library</a>, a member of the Minuteman Library Network'
          WHEN p.ptype_code = 23 THEN 'the Medway Public Library</a>, a member of the Minuteman Library Network'
          WHEN p.ptype_code = 24 THEN 'the Millis Public Library</a>, a member of the Minuteman Library Network'
          WHEN (p.ptype_code = 26 AND p.home_library_code != 'na2z') THEN 'the Morse Institute Library</a>, a member of the Minuteman Library Network'
          WHEN (p.ptype_code = 26 AND p.home_library_code = 'na2z') THEN 'the Bacon Free Library</a>, a member of the Minuteman Library Network'
          WHEN p.ptype_code = 27 THEN 'the Needham Free Public Library</a>, a member of the Minuteman Library Network'
          WHEN p.ptype_code IN ('30', '130') THEN 'the Morrill Memorial Library</a>, a member of the Minuteman Library Network'
          WHEN p.ptype_code = 32 THEN 'the Randall Library</a>, a member of the Minuteman Library Network'
          WHEN p.ptype_code IN ('33', '133') THEN 'the Goodnow Library</a>, a member of the Minuteman Library Network'
          WHEN p.ptype_code = 34 THEN 'the Waltham Public Library</a>, a member of the Minuteman Library Network'
          WHEN p.ptype_code = 36 THEN 'the Wayland Free Library</a>, a member of the Minuteman Library Network'
          WHEN p.ptype_code = 38 THEN 'the Weston Public Library</a>, a member of the Minuteman Library Network'
          WHEN p.ptype_code = 39 THEN 'the Westwood Public Library</a>, a member of the Minuteman Library Network'
          WHEN p.ptype_code = 40 THEN 'the Winchester Public Library</a>, a member of the Minuteman Library Network'
          WHEN p.ptype_code = 41 THEN 'the Woburn Public Library</a>, a member of the Minuteman Library Network'
          WHEN p.ptype_code = 46 THEN 'the Sherborn Public Library</a>, a member of the Minuteman Library Network'
          ELSE 'the Minuteman Library Network</a>'
        END AS library_2,
        CASE
          WHEN p.ptype_code = 5 THEN 'https://belmontpubliclibrary.net/'
          WHEN p.ptype_code = 6 THEN 'https://www.brooklinelibrary.org/'
          WHEN p.ptype_code = 7 THEN 'https://www.cambridgema.gov/cpl'
          WHEN p.ptype_code = 8 THEN 'https://concordlibrary.org/'
          WHEN p.ptype_code IN ('10','110') THEN 'https://www.dedhamlibrary.com/'
          WHEN p.ptype_code IN ('17', '117') THEN 'https://www.carylibrary.org/'
          WHEN p.ptype_code IN ('29', '129') THEN 'https://newtonfreelibrary.net/'
          WHEN p.ptype_code = 31 THEN 'https://www.somervillepubliclibrary.org/'
          WHEN p.ptype_code = 35 THEN 'https://watertownlib.org/welcome'
          WHEN p.ptype_code IN ('37', '137') THEN 'https://www.wellesleyfreelibrary.org/'
          WHEN p.ptype_code = 1 THEN 'http://www.actonmemoriallibrary.org/'
          WHEN p.ptype_code = 2 THEN 'https://www.robbinslibrary.org/'
          WHEN p.ptype_code = 3 THEN 'https://www.ashlandmass.com/184/Ashland-Public-Library'
          WHEN p.ptype_code = 4 THEN 'https://www.bedfordlibrary.net/'
          WHEN p.ptype_code = 11 THEN 'https://dovertownlibrary.org/'
          WHEN p.ptype_code = 12 THEN 'https://framinghamlibrary.org/'
          WHEN p.ptype_code = 14 THEN 'https://www.franklinma.gov/franklin-public-library'
          WHEN p.ptype_code IN ('15', '115') THEN 'https://www.hollistonlibrary.org/'
          WHEN p.ptype_code = 18 THEN 'https://www.lincolnpl.org/'
          WHEN p.ptype_code IN ('20', '120') THEN 'https://www.maynardpubliclibrary.org/'
          WHEN p.ptype_code IN ('21', '121') THEN 'https://www.medfieldpubliclibrary.org/'
          WHEN p.ptype_code IN ('22', '122') THEN 'https://www.medfordlibrary.org/'
          WHEN p.ptype_code = 23 THEN 'https://medwaylib.org/'
          WHEN p.ptype_code = 24 THEN 'https://www.millislibrary.org/'
          WHEN (p.ptype_code = 26 AND p.home_library_code != 'na2z') THEN 'https://morseinstitute.org/'
          WHEN (p.ptype_code = 26 AND p.home_library_code = 'na2z') THEN 'https://baconfreelibrary.org/'
          WHEN p.ptype_code = 27 THEN 'https://needhamlibrary.org/'
          WHEN p.ptype_code IN ('30', '130') THEN 'http://www.norwoodlibrary.org/'
          WHEN p.ptype_code = 32 THEN 'https://www.stow-ma.gov/randall-library/'
          WHEN p.ptype_code IN ('33', '133') THEN 'https://goodnowlibrary.org/'
          WHEN p.ptype_code = 34 THEN 'https://waltham.lib.ma.us/'
          WHEN p.ptype_code = 36 THEN 'https://waylandlibrary.org/'
          WHEN p.ptype_code = 38 THEN 'https://www.westonlibrary.org/'
          WHEN p.ptype_code = 39 THEN 'https://www.westwoodlibrary.org/'
          WHEN p.ptype_code = 40 THEN 'https://www.winpublib.org/'
          WHEN p.ptype_code = 41 THEN 'https://woburnpubliclibrary.org/'
          WHEN p.ptype_code = 46 THEN 'https://sherbornlibrary.org/'
          ELSE 'https://www.minlib.net/our-libraries'
        END AS url
      FROM sierra_view.patron_view as p
      JOIN sierra_view.varfield v		
        ON p.id = v.record_id
        AND v.varfield_type_code = 'z'
      JOIN sierra_view.patron_record_fullname n
        ON p.id = n.patron_record_id
      JOIN sierra_view.record_metadata m
        ON p.record_num = m.record_num
        AND m.record_type_code = 'p'
      WHERE m.creation_date_gmt::date = (CURRENT_DATE - interval '1 day')
        AND p.patron_agency_code_num != '47'
        --Opt out list, CAM, NTN
        AND p.ptype_code IN('35', '36', '137', '37', '38', '39', '40', '41', '43')  
        AND p.ptype_code NOT IN('7', '9', '13', '16', '19', '25', '28', '29', '44', '45', '116', '129', '159', '163', '166', '169', '175', '178', '194', '195', '199', '200', '201', '202', '203', '204', '205', '206', '207', '254', '255')
      GROUP BY 4, 5, 6, 7
      """

    query_results = run_query(query)

    for rownum, row in enumerate(query_results):

        # emailto can send to multiple addresses by separating emails with commas
        emailto = [str(row[2])]
        emailsubject = "Enjoy your new library card!"
        # Creating the email message
        email_text = """Dear {} {},   
Welcome! You have been registered for a library card from {}.
    
With your new library card you can:
    
    -Borrow and request items from all Minuteman member libraries.  Most materials can be returned at any Minuteman location.
    -Visit the shared catalog at find.minlib.net  to discover books, movies, music, audiobooks, and much more.
    -Access digital ebooks, audiobooks, magazines and streaming video from OverDrive.
    -Renew items, manage your account, and track your reading history with MyAccount. To get started, go to find.minlib.net/iii/encore/myaccount and follow the instructions to set-up your login.
    -You can also save time by adding the Minuteman App and Text Message Notifications to your mobile phone or tablet.
    -Explore Minuteman libraries’ ever-expanding collections of toys and games, technology, household tools, musical instruments, and more. 
    
""".format(
            str(row[0]), str(row[1]), str(row[5])
        )
        if str(row[4]) != "the Minuteman Library Network":
            email_text += """ 
Visit your home library's website ({}) or talk to your librarian to find out about even more online collections, services, and events offered by your local library.<br><br>
        
""".format(
                str(row[6])
            )
        email_text += """Enjoy your new library card!
    
***This is an automated email.  Do not reply.***"""

        email_html = """
    <html>
    <head></head>
    <body style="background-color:#FFFFFF;">
    <table style="width: 70%; margin-left: 15%; margin-right: 15%; border: 0; cellspacing: 0; cellpadding: 0; background-color: #FFFFFF;">
    <tr>
    <font face="Scala Sans, Calibri, Arial"; size="3">
    <p>Dear {} {},<br><br>    
    Welcome! You have been registered for a library card from <a href="{}">{}.<br><br>
        
    With your new library card you can:<br><br>
    <ul>
    <li>Borrow and request items from all <a href="https://www.minlib.net/our-libraries">Minuteman member libraries</a>.  Most materials can be returned at any Minuteman location.</li>
    <li>Visit the shared catalog at <a href="https://catalog.minlib.net">catalog.minlib.net</a> to discover books, movies, music, audiobooks, and much more.</li>
    <li>Access digital ebooks, audiobooks, magazines and streaming video from <a href="minuteman.overdrive.com">OverDrive.</a></li>
    <li>Renew items, manage your account, and track your reading history with MyAccount.  To get started, go to <a href="https://catalog.minlib.net/MyAccount">catalog.minlib.net/MyAccount</a> and follow the instructions to set-up your login.</li>
    <li>You can also save time by adding the <a href="https://www.minlib.net/services#mln-app">Minuteman App</a> and <a href="https://www.minlib.net/services#text-messages">Text Message Notifications</a> to your mobile phone or tablet.</li>
    <li>Explore Minuteman libraries’ ever-expanding collections of toys and games, technology, household tools, musical instruments, and more.</li>
    </ul><br>
    """.format(
            str(row[0]), str(row[1]), str(row[6]), str(row[5])
        )
        if str(row[4]) != "the Minuteman Library Network":
            email_html += """ 
	    Visit <a href="{}">your home library’s website</a> or talk to your librarian to find out about even more online collections, services, and events offered by your local library.<br><br>
        """.format(
                str(row[6])
            )
        email_html += """Enjoy your new library card!<br><br>
    ***This is an automated email.  Do not reply.***    
    </font>
    </p>
    <img src="https://www.minlib.net/sites/default/files/glazed_builder_images/logo-print-small.jpg" alt="Minuteman logo" height="32" width="188">
    </tr>
    </table>
    </body>  
    </html>
    """
        send_email(emailsubject, email_text, email_html, emailto)


main()
