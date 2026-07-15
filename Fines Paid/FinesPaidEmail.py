'''
Jeremy Goldstein
Minuteman Library Network
Sends email alert to designated staff to inform them of a newly available report located
on our staff site.  Script schedule to run after the one producing that report.
'''

import smtplib
import configparser
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email.utils import formatdate
from email import encoders
from datetime import date, timedelta
import traceback

# function constructs and sends outgoing email given a subject, a recipient and body text in both txt and html forms
def send_email(subject, message_text, message_html, recipient, replyto):
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
    msg.add_header('reply-to', replyto)
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

    # read config file with Sierra login credentials
    config = configparser.ConfigParser()
    config.read("C:\\Scripts\\Creds\\emails.ini")

    emailto = config["fines_paid"]["recipients"].split()
    replyto = config["fines_paid"]["reply_to"]
    emailsubject = "Monthly ECommerce Reports Now Available"
    last_month = date.today().replace(day=1) - timedelta(1)
    email_text = """***This is an automated email.  Do not reply.***
    
Dear Directors,
  
The monthly ECommerce reports for {} are now available here: https://staff.minlib.net/reports/static?path=%2Freports%2FLibrary-Specific%20Reports">Library-Specific Reports section</a> of the Staff Information Center. <br><br>
The Excel spreadsheet includes fines (including lost or replacement fines) that pointed to an item owned by your library.
More information on this report can be found here: https://docs.google.com/document/d/1eFifxBsQdGnwrZoDaoA-egbykHTqb2kG_bIN789ubO4/edit#heading=h.xbpej6bnbbsf

Thank you,
Sharon

""".format(last_month.strftime("%B %Y"))

    email_html = html = """
<html>
<head></head>
<body>
<p>***This is an automated email.***<br><br>
Dear Directors,<br><br>
The monthly monthly ECommerce reports for {} are now available in the <a href="https://staff.minlib.net/reports/static?path=%2Freports%2FLibrary-Specific%20Reports">Library-Specific Reports section</a> of the Staff Information Center. <br><br>
The Excel spreadsheet includes fines (including lost or replacement fines) that pointed to an item owned by your library.<br><br>
More information on this report can be found <a href="https://docs.google.com/document/d/1eFifxBsQdGnwrZoDaoA-egbykHTqb2kG_bIN789ubO4/edit#heading=h.xbpej6bnbbsf">here</a>.<br><br>

Thank you,<br>
Sharon
</p>
</body>  
</html>
""".format(last_month.strftime("%B %Y"))
    
    send_email(emailsubject, email_text, email_html, emailto, replyto)

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
        email_subject = "Ecommerce fines paid email script error"
        email_message = (
            "Your script failed with the following error:\n\n" + traceback.format_exc()
        )

        send_email_error(email_subject, email_message, emailto)
        raise