#!/usr/bin/env python3

#Run in py313
"""
Jeremy Goldstein
Minuteman Library Network

Generates report on the monthly
circ totals for Woburn items by custom scat table
"""

import psycopg2
import xlsxwriter
import os
import configparser
import smtplib
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
    #Creating the Excel file for staff
    workbook = xlsxwriter.Workbook(excel_file, {'remove_timezone': True})
    worksheet = workbook.add_worksheet()


    #Formatting our Excel worksheet
    worksheet.set_landscape()
    worksheet.hide_gridlines(0)

    #Formatting Cells
    eformat= workbook.add_format({'text_wrap': True, 'valign': 'top'})
    eformatlabel= workbook.add_format({'text_wrap': True, 'valign': 'top', 'bold': True})


    # Setting the column widths
    worksheet.set_column(0,0,10.29)
    worksheet.set_column(1,1,38.71)
    worksheet.set_column(2,2,14.14)
    worksheet.set_column(3,3,23.86)
    worksheet.set_column(4,4,19.29)

    #Inserting a header
    worksheet.set_header('Woburn Monthly Circ By Scat')

    # Adding column labels
    worksheet.write(0,0,'Scat Code', eformatlabel)
    worksheet.write(0,1,'Scat', eformatlabel)
    worksheet.write(0,2,'Total Circ', eformatlabel)
    worksheet.write(0,3,'Total Circ Local', eformatlabel)
    worksheet.write(0,4,'Total Circ Network', eformatlabel)


    # Writing the report for staff to the Excel worksheet
    for rownum, row in enumerate(query_results):
        worksheet.write(rownum+1,0,row[0], eformat)
        worksheet.write(rownum+1,1,row[1], eformat)
        worksheet.write(rownum+1,2,row[2], eformat)
        worksheet.write(rownum+1,3,row[3], eformat)
        worksheet.write(rownum+1,4,row[4], eformat)
    
    workbook.close()
    

# function takes a file as a parameter and attaches that file to an outgoing email
def send_email(subject, message, attachment, recipient):
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
    # plain text of email message
    emailmessage = message

    # Creating the email message
    msg = MIMEMultipart()
    msg["From"] = emailfrom
    if type(recipient) is list:
        msg["To"] = ", ".join(recipient)
    else:
        msg["To"] = recipient
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

    query = r"""
SELECT
  i.icode1 AS scat_code,
  CASE
    WHEN i.icode1='0' THEN 'BONNIE NO SCAT'
    WHEN i.icode1='1' THEN 'FICTION'
    WHEN i.icode1='2' THEN 'MYSTERY'
    WHEN i.icode1='3' THEN 'SCIENCE FICTION'
    WHEN i.icode1='4' THEN 'WESTERN'
    WHEN i.icode1='5' THEN 'PAPERBACK FICTION RACK'
    WHEN i.icode1='6' THEN 'LARGE PRINT FICTION'
    WHEN i.icode1='7' THEN 'ON ORDER'
    WHEN i.icode1='8' THEN 'circulating CD-ROM [software]'
    WHEN i.icode1='9' THEN 'FICTION CLASSIC'
    WHEN i.icode1='10' THEN '100   Philosophy'
    WHEN i.icode1='11' THEN '110   Metaphysics'
    WHEN i.icode1='12' THEN '120   Knowledge'
    WHEN i.icode1='13' THEN '130   Parapsychology'
    WHEN i.icode1='14' THEN '140   Philosophy'
    WHEN i.icode1='15' THEN '150   Psychology'
    WHEN i.icode1='16' THEN '160   Logic'
    WHEN i.icode1='17' THEN '170   Ethics'
    WHEN i.icode1='18' THEN '180   Ancient Philosophy'
    WHEN i.icode1='19' THEN '190   Modern Philosophy'
    WHEN i.icode1='20' THEN '200   Religion'
    WHEN i.icode1='21' THEN '210   Natural Religion'
    WHEN i.icode1='22' THEN '220   Bible'
    WHEN i.icode1='23' THEN '230   Christian Theology'
    WHEN i.icode1='24' THEN '240   Christian Devotions'
    WHEN i.icode1='25' THEN '250   Local Churches'
    WHEN i.icode1='26' THEN '260   Ecclesiastical Theology'
    WHEN i.icode1='27' THEN '270   Church History'
    WHEN i.icode1='28' THEN '280   Christian Denominations'
    WHEN i.icode1='29' THEN '290   Other Religions'
    WHEN i.icode1='30' THEN '300   Social Sciences'
    WHEN i.icode1='31' THEN '310   Statistics'
    WHEN i.icode1='32' THEN '320   Political Science'
    WHEN i.icode1='33' THEN '330   Economics'
    WHEN i.icode1='34' THEN '340   Law'
    WHEN i.icode1='35' THEN '350   Public Institutions'
    WHEN i.icode1='36' THEN '360   Social Path.'
    WHEN i.icode1='37' THEN '370   Education'
    WHEN i.icode1='38' THEN '380   Commerce'
    WHEN i.icode1='39' THEN '390   Customs & Folklore'
    WHEN i.icode1='40' THEN '400   Language'
    WHEN i.icode1='41' THEN '410   Linguistic'
    WHEN i.icode1='42' THEN '420   Eng. & A - S'
    WHEN i.icode1='43' THEN '430   German'
    WHEN i.icode1='44' THEN '440   French'
    WHEN i.icode1='45' THEN '450   Italian'
    WHEN i.icode1='46' THEN '460   Span. & Port.'
    WHEN i.icode1='47' THEN '470   Latin'
    WHEN i.icode1='48' THEN '480   Greek'
    WHEN i.icode1='49' THEN '490   Other'
    WHEN i.icode1='50' THEN '500   Pure Sci.'
    WHEN i.icode1='51' THEN '510   Math'
    WHEN i.icode1='52' THEN '520   Astronomy'
    WHEN i.icode1='53' THEN '530   Physics'
    WHEN i.icode1='54' THEN '540   Chemistry'
    WHEN i.icode1='55' THEN '550   Earth Sciences'
    WHEN i.icode1='56' THEN '560   Paleontology'
    WHEN i.icode1='57' THEN '570   Life Sciences'
    WHEN i.icode1='58' THEN '580   Botany'
    WHEN i.icode1='59' THEN '590   Zoology'
    WHEN i.icode1='60' THEN '600   Technology'
    WHEN i.icode1='61' THEN '610   Medicine'
    WHEN i.icode1='62' THEN '620   Engineering'
    WHEN i.icode1='63' THEN '630   Agriculture'
    WHEN i.icode1='64' THEN '640   Domestic Science'
    WHEN i.icode1='65' THEN '650   Management'
    WHEN i.icode1='66' THEN '660   Chem. Tech.'
    WHEN i.icode1='67' THEN '670   Manufacturing'
    WHEN i.icode1='68' THEN '680   Misc. Man.'
    WHEN i.icode1='69' THEN '690   Building'
    WHEN i.icode1='70' THEN '700   Arts'
    WHEN i.icode1='71' THEN '710   Landscaping'
    WHEN i.icode1='72' THEN '720   Architecture'
    WHEN i.icode1='73' THEN '730   Sculpture'
    WHEN i.icode1='74' THEN '740   Drawing'
    WHEN i.icode1='75' THEN '750   Painting'
    WHEN i.icode1='76' THEN '760   Graphics'
    WHEN i.icode1='77' THEN '770   Photography'
    WHEN i.icode1='78' THEN '780   Music'
    WHEN i.icode1='79' THEN '790   Sports-Rec.'
    WHEN i.icode1='80' THEN '800   Literature'
    WHEN i.icode1='81' THEN '810   American Literature'
    WHEN i.icode1='82' THEN '820   English Literature'
    WHEN i.icode1='83' THEN '830   German Literature'
    WHEN i.icode1='84' THEN '840   French Literature'
    WHEN i.icode1='85' THEN '850   Italian Literature'
    WHEN i.icode1='86' THEN '860   Spanish & Portuguese Lit'
    WHEN i.icode1='87' THEN '870   Latin Literature'
    WHEN i.icode1='88' THEN '880   Greek Literature'
    WHEN i.icode1='89' THEN '890   Other Literature'
    WHEN i.icode1='90' THEN '900   Geography & History'
    WHEN i.icode1='91' THEN '910   Travel'
    WHEN i.icode1='92' THEN '920   920-929 & Biography'
    WHEN i.icode1='93' THEN '930   Ancient History'
    WHEN i.icode1='94' THEN '940   European History'
    WHEN i.icode1='95' THEN '950   Asian History'
    WHEN i.icode1='96' THEN '960   African History'
    WHEN i.icode1='97' THEN '970   North American History'
    WHEN i.icode1='98' THEN '980   South American History'
    WHEN i.icode1='99' THEN '990   Other History'
    WHEN i.icode1='100' THEN '000   General'
    WHEN i.icode1='101' THEN 'PERIODICALS (ADULT)'
    WHEN i.icode1='102' THEN '001-006 Computers'
    WHEN i.icode1='103' THEN 'LARGE PRINT NONFICTION'
    WHEN i.icode1='104' THEN 'GOVERNMENT DOCS.'
    WHEN i.icode1='105' THEN 'PAMPHLETS / VERTICAL FILE'
    WHEN i.icode1='106' THEN 'COLLEGE CATALOGS'
    WHEN i.icode1='107' THEN 'WOBURN COLLECTION'
    WHEN i.icode1='108' THEN 'PAPERBACK MYSTERY'
    WHEN i.icode1='109' THEN 'PAPERBACK ROMANCE'
    WHEN i.icode1='110' THEN 'PAPERBACK SCIENCE FICTION'
    WHEN i.icode1='111' THEN 'WOBURN READS'
    WHEN i.icode1='112' THEN 'MUSEUM PASSES'
    WHEN i.icode1='113' THEN 'ANNUAL REPORTS'
    WHEN i.icode1='114' THEN 'OFFICE & PROFESSIONAL'
    WHEN i.icode1='115' THEN 'World Language Collection'
    WHEN i.icode1='116' THEN 'LARGE PRINT MYSTERY'
    WHEN i.icode1='117' THEN 'LARGE PRINT WESTERN'
    WHEN i.icode1='118' THEN 'TALKING BOOKS (via PERKINS)'
    WHEN i.icode1='119' THEN 'z not in use'
    WHEN i.icode1='120' THEN 'AUDIO-CASSETTES (MUSIC)'
    WHEN i.icode1='121' THEN 'z not in use'
    WHEN i.icode1='122' THEN 'MYSTERY SHORT STORIES'
    WHEN i.icode1='123' THEN 'SPEED READ FICTION'
    WHEN i.icode1='124' THEN 'SPEED READ NONFICTION'
    WHEN i.icode1='125' THEN 'Childrens Library of Things, Objects.'
    WHEN i.icode1='126' THEN 'Adult Library of Things, Objects.'
    WHEN i.icode1='127' THEN 'PLAYAWAY'
    WHEN i.icode1='128' THEN 'ILL NON-MLN'
    WHEN i.icode1='129' THEN 'AUDIO CD [spoken audio cd]'
    WHEN i.icode1='130' THEN 'z not in use'
    WHEN i.icode1='131' THEN 'AUDIO CD language'
    WHEN i.icode1='132' THEN 'z not in use'
    WHEN i.icode1='133' THEN 'FANTASY FICTION'
    WHEN i.icode1='134' THEN 'HORROR FICTION'
    WHEN i.icode1='135' THEN 'VIDEOCASSETTES (NON-FIC.)'
    WHEN i.icode1='136' THEN 'VIDEOCASSETTES (FEATURE)'
    WHEN i.icode1='137' THEN 'STAFF COLL (noncirc)'
    WHEN i.icode1='138' THEN 'EQUIPMENT (AV, CAMERA) Public Laptops & I-Pads'
    WHEN i.icode1='139' THEN 'CAREER/EDUCATION CENTER '
    WHEN i.icode1='140' THEN 'ART PRINTS'
    WHEN i.icode1='141' THEN 'CD ?'
    WHEN i.icode1='142' THEN 'SCULPTURE'
    WHEN i.icode1='143' THEN 'Equipment -  Hotspot'
    WHEN i.icode1='144' THEN 'CD '
    WHEN i.icode1='145' THEN 'Adult Videogames'
    WHEN i.icode1='146' THEN 'LARGE PRINT ROMANCE'
    WHEN i.icode1='147' THEN 'DVD (NONFICTION)'
    WHEN i.icode1='148' THEN 'DVD (FEATURE)'
    WHEN i.icode1='149' THEN 'DVD TV Series Itype 19,20 [Juv & Adult]'
    WHEN i.icode1='150' THEN 'PLAYAWAY, J'
    WHEN i.icode1='151' THEN 'CD ?'
    WHEN i.icode1='152' THEN 'CD ?'
    WHEN i.icode1='153' THEN 'TEEN MANGA'
    WHEN i.icode1='154' THEN 'MUSIC CD Pop'
    WHEN i.icode1='155' THEN 'MUSIC CD Country'
    WHEN i.icode1='156' THEN 'VIDEOCASSETTES ?'
    WHEN i.icode1='157' THEN 'VIDEOCASSETTES ?'
    WHEN i.icode1='158' THEN 'Boardgames'
    WHEN i.icode1='159' THEN 'TEEN VIDEOGAMES'
    WHEN i.icode1='160' THEN 'TEEN LARGE PRINT'
    WHEN i.icode1='161' THEN 'TEEN FICTION'
    WHEN i.icode1='162' THEN 'TEEN PERIODICALS'
    WHEN i.icode1='163' THEN 'TEEN SF'
    WHEN i.icode1='164' THEN 'TEEN SHORT STORIES, COLLECTIONS'
    WHEN i.icode1='165' THEN 'TEEN PAPERBACK'
    WHEN i.icode1='166' THEN 'TEEN NONFICTION'
    WHEN i.icode1='167' THEN 'TEEN GRAPHIC NOVEL'
    WHEN i.icode1='168' THEN 'TEEN AUDIO CD'
    WHEN i.icode1='169' THEN 'GRAPHIC NOVEL'
    WHEN i.icode1='170' THEN 'SCIENCE FAIR/PROJECTS'
    WHEN i.icode1='171' THEN 'G1'
    WHEN i.icode1='172' THEN 'G2'
    WHEN i.icode1='173' THEN 'G3'
    WHEN i.icode1='174' THEN 'local history upstairs'
    WHEN i.icode1='175' THEN 'local history Industry'
    WHEN i.icode1='176' THEN 'Mass'
    WHEN i.icode1='177' THEN 'Mass A'
    WHEN i.icode1='178' THEN 'Mass C'
    WHEN i.icode1='179' THEN 'Mass M'
    WHEN i.icode1='180' THEN 'Mass M2'
    WHEN i.icode1='181' THEN '906 Mass S'
    WHEN i.icode1='182' THEN 'Mass Towns'
    WHEN i.icode1='183' THEN 'Glennon Archives Coll'
    WHEN i.icode1='184' THEN 'Winn Collection'
    WHEN i.icode1='185' THEN '905'
    WHEN i.icode1='186' THEN 'local history Industry?'
    WHEN i.icode1='187' THEN 'Woburn authors'
    WHEN i.icode1='188' THEN '906'
    WHEN i.icode1='189' THEN ''
    WHEN i.icode1='190' THEN 'JUV GRAPHIC NOVELS'
    WHEN i.icode1='191' THEN 'CALDECOTT winners'
    WHEN i.icode1='192' THEN 'JUV early chapter books sci fi'
    WHEN i.icode1='193' THEN 'JUV early chapter books fantasy'
    WHEN i.icode1='194' THEN 'JUV EASY-read BIOGRAPHIES'
    WHEN i.icode1='195' THEN 'JUV early chapter books'
    WHEN i.icode1='196' THEN 'JUV ALPHABET/COUNTING BOOKS AND ACTION'
    WHEN i.icode1='197' THEN 'JUV early chapter books mystery'
    WHEN i.icode1='198' THEN 'JUV VIDEO GAMES'
    WHEN i.icode1='199' THEN 'JUV spoken word CD-ROM and book'
    WHEN i.icode1='200' THEN 'JUV BOARD BOOKS'
    WHEN i.icode1='201' THEN 'JUV FICTION'
    WHEN i.icode1='202' THEN 'JUV MYSTERY'
    WHEN i.icode1='203' THEN 'picture book for older readers'
    WHEN i.icode1='204' THEN 'JUV SCIENCE FICTION'
    WHEN i.icode1='205' THEN 'JUV FANTASY'
    WHEN i.icode1='206' THEN 'JUV PICTURE BOOKS, mini picture books'
    WHEN i.icode1='207' THEN 'JUV HOLIDAY picture books'
    WHEN i.icode1='208' THEN 'JUV PARENT-Teacher Resources Collection'
    WHEN i.icode1='209' THEN 'JUV EASY READERS'
    WHEN i.icode1='210' THEN 'JUV 000   General'
    WHEN i.icode1='211' THEN 'JUV 100   Philosophy & Psych'
    WHEN i.icode1='212' THEN 'JUV 200   Religion'
    WHEN i.icode1='213' THEN 'JUV 300   Social Sciences'
    WHEN i.icode1='214' THEN 'JUV 400   Language'
    WHEN i.icode1='215' THEN 'JUV 500   Science'
    WHEN i.icode1='216' THEN 'JUV 600   Technology'
    WHEN i.icode1='217' THEN 'JUV 700   Art & Recreation'
    WHEN i.icode1='218' THEN 'JUV 800   Literature'
    WHEN i.icode1='219' THEN 'JUV 900   History & Travel'
    WHEN i.icode1='220' THEN 'JUV BIOGRAPHIES / AUTOBIOG.'
    WHEN i.icode1='221' THEN 'JUV PERIODICALS'
    WHEN i.icode1='222' THEN 'JUV REFERENCE'
    WHEN i.icode1='225' THEN 'JUV  spoken-word CD '
    WHEN i.icode1='226' THEN 'J series'
    WHEN i.icode1='228' THEN 'JUV KITS'
    WHEN i.icode1='229' THEN 'J Book with CD-ROM ("KIT")'
    WHEN i.icode1='230' THEN 'J CD-ROM game (program)'
    WHEN i.icode1='231' THEN 'J Music Cassette'
    WHEN i.icode1='232' THEN 'J Big Book'
    WHEN i.icode1='233' THEN 'J StoryHour'
    WHEN i.icode1='234' THEN 'J music CD'
    WHEN i.icode1='235' THEN 'J DVD NF'
    WHEN i.icode1='236' THEN 'J DVD feature'
    WHEN i.icode1='237' THEN 'World Languages (about or in language)'
    WHEN i.icode1='238' THEN 'World Languages - bilingual books'
    WHEN i.icode1='239' THEN 'J Braille books'
    WHEN i.icode1='241' THEN 'SPOKEN CASSETTES'
    WHEN i.icode1='243' THEN 'LANGUAGE CASSETTES'
    WHEN i.icode1='244' THEN 'ELECTRONIC RESOURCE (in library use)'
    WHEN i.icode1='245' THEN 'ELECTRONIC RESOURCE (remote and in library use)'
    WHEN i.icode1='248' THEN 'SHORT STORIES COLLECTIONS'
    WHEN i.icode1='249' THEN 'REFERENCE'
    ELSE 'N/A'
  END AS scat,
  COUNT(t.id) FILTER(WHERE t.op_code IN ('o','r')) AS total_circ,
  COUNT(t.id) FILTER(WHERE t.op_code IN ('o','r') AND t.stat_group_code_num BETWEEN '790' AND '799') AS total_circ_local,
  COUNT(t.id) FILTER(WHERE t.op_code IN ('o','r') AND t.stat_group_code_num NOT BETWEEN '790' AND '799') AS total_circ_network

FROM sierra_view.item_record i
JOIN sierra_view.circ_trans t
  ON i.id = t.item_record_id
  AND t.transaction_gmt::DATE >= CURRENT_DATE - INTERVAL '1 MONTH'

WHERE i.location_code ~ '^wob'

GROUP BY 1,2
ORDER BY 1
            """
    
    query_results = run_query(query)

    excel_file =  "/Scripts/Woburn Monthly Circ By Scat/Temp Files/wob monthly circ by scat{}.xlsx".format(date.today())
    excel_writer(query_results, excel_file)

    # send email
    email_subject = "WOB monthly circ by scat"
    email_message = """***This is an automated email***


The Woburn Monthly Circ By Scat Code Report has been attached."""
    # read config file with recipient list for email
    config_recipient = configparser.ConfigParser()
    config_recipient.read("C:\\Scripts\\Creds\\emails.ini")
    recipient = config_recipient["woburn_circ_by_scat"]["recipients"].split()
    send_email(email_subject, email_message, excel_file, recipient)

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
        emailto = config_recipient["script_error_extended"]["recipients"].split()

        # craft email subject and message containing error message details from traceback
        email_subject = "Woburn Circ By Scat script error"
        email_message = (
            "Your script failed with the following error:\n\n" + traceback.format_exc()
        )

        send_email_error(email_subject, email_message, emailto)
        raise


