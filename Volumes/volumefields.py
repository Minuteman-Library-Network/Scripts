#!/usr/bin/env python3

# Run in py313

"""
Create and email a list of items with volume fields that should not contain them
to a designated staff member to assist with record cleanup

Author: Jeremy Goldstein
Contact Info: jgoldstein@minlib.net
"""

import psycopg2
import configparser
import csv
import smtplib
import os
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
    # Gather column headers, which are not included in cursor.fetchall() and store in another variable
    columns = [i[0] for i in cursor.description]
    # close database connection
    conn.close()
    # return variables containing query results and column headers
    return rows, columns

# function takes the results of a query and converts them to a csv file
def write_csv(query_results, headers, csv_file):
    # open csvfile in write mode and add a row to it for the headers and each line of query_results
    with open(csv_file, "w", encoding="utf-8", newline="") as tempFile:
        myFile = csv.writer(tempFile, delimiter=",")
        myFile.writerow(headers)
        myFile.writerows(query_results)
    tempFile.close()
    # return variable containing the newly created csv file
    return csv_file



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
    emailto = config_recipient["volumes"]["recipients"].split()
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
    # query to identify patron records with incorrect owed_amt fields
    query = r"""
            SELECT *
            FROM (
              SELECT
                m.record_type_code||m.record_num||'a' AS record_num,
                bp.best_title AS title,
                COUNT(CASE WHEN v.field_content IS NOT NULL THEN 1 ELSE NULL END) AS volume_count,
                COUNT(i.id) AS item_count

              FROM sierra_view.bib_record b
              JOIN sierra_view.bib_record_item_record_link l
                ON b.id = l.bib_record_id
              JOIN sierra_view.item_record i
                ON i.id = l.item_record_id
                AND i.item_message_code != 'f'
              JOIN sierra_view.record_metadata rmi
                ON i.id = rmi.id
              JOIN sierra_view.bib_record_property bp
                ON b.id = bp.bib_record_id
              JOIN sierra_view.record_metadata m
                ON b.id = m.id
              LEFT JOIN sierra_view.varfield v
                ON i.id = v.record_id
                AND v.varfield_type_code = 'v'
              
              WHERE b.bcode3 NOT IN ('g','a')
                AND m.record_num NOT IN (
'2998103',
'1989394',
'3058721',
'2199988',
'3507262',
'2456816',
'3069146',
'2464205',
'3713898',
'2499131',
'3628444',
'2847132',
'2347604',
'2275155',
'2992970',
'2936630',
'2985430',
'3639087',
'3771271',
'2325993',
'2994756',
'3149139',
'3539372',
'3455520',
'3103773',
'2891865',
'2313066',
'3236669',
'3278466',
'2925558',
'2279043',
'3725861',
'2223803',
'2658487',
'2923153',
'2232491',
'3412316',
'2337195',
'2343194',
'2281084',
'3013672',
'2464945',
'2447863',
'3038185',
'2107099',
'2623458',
'1609345',
'2540712',
'2105484',
'2281292',
'3742462',
'2952223',
'2770836',
'3028235',
'3622227',
'2665142',
'2397564',
'3761377',
'3538506',
'2837577',
'3174940',
'2615127',
'3733122',
'1383967',
'2596908',
'2206037',
'1775182',
'3774718',
'1251240',
'3757111',
'3652642',
'3821950',
'3160723',
'1005102',
'1005102',
'3742780',
'2621951',
'2551172',
'2550953',
'2981843',
'2390675',
'2844692',
'2164369',
'2606319',
'1630820',
'3711941',
'2314059',
'2615091',
'3732912',
'3808520',
'2369556',
'2526007',
'3766942',
'3276042',
'2865054',
'3652524',
'1962622',
'3276040',
'2359936',
'3534010',
'3223486',
'1652267',
'2241200',
'3772703',
'2487721',
'2209740',
'2382472',
'3651221',
'1929527',
'3661196',
'2735244',
'2612563',
'3214622',
'3790350',
'2267999',
'2637581',
'3727928',
'2634648',
'3640161',
'3478260',
'3735309',
'2538240',
'3182363',
'2183339',
'3712505',
'3153166',
'3198334',
'3789927',
'3856477',
'2752559',
'2768468',
'3192210',
'2839214',
'2780895',
'2151584',
'2111759',
'2748384',
'3867795',
'3225069',
'2398356',
'2751679',
'3719252',
'2520158',
'2524418',
'2503192',
'2615916',
'3175652',
'3713286',
'3778375',
'2979248',
'1480704',
'3536915',
'2111277',
'2475665',
'3101496',
'3157059',
'3192249',
'3733767',
'3863083',
'2136758',
'2276391',
'3854249',
'2764221',
'2848948',
'2174309',
'2703531',
'3639712',
'2618076',
'2162950',
'2938725',
'2912995',
'3917018',
'3815410',
'3810307',
'2164723',
'1728015',
'2837580',
'3197431',
'2115157',
'2151582',
'3851619',
'3466199',
'2229898',
'2483066',
'2366875',
'3943241',
'2374014',
'2236243',
'2282433',
'2326029',
'3176085',
'3912980',
'1052278',
'1445469',
'2462748',
'3913215',
'3906423',
'2391933',
'3785391',
'2844353',
'3680902',
'3938471',
'2748335',
'2664749',
'3985342',
'3004825',
'2401782',
'2937967',
'3095656',
'2320374',
'2992889',
'2189617',
'2050789',
'1027834',
'2658760',
'3919249',
'2118346',
'4004838',
'3912690',
'4004839',
'3983701',
'3911872',
'2893643',
'3646350',
'3713290',
'3139536',
'4013469',
'3733703',
'2882334',
'3109561',
'2631258',
'2374076',
'3990600',
'2839085',
'2181593',
'2280124',
'2374626',
'3715913',
'2247617',
'3981635',
'3163285',
'3745360',
'4020318',
'2531863',
'3875759',
'2619343',
'2265273',
'3836058',
'3432828',
'3567533',
'4047800',
'2390637',
'2545738',
'2616080',
'2837585',
'2998686',
'9184720',
'3753374',
'2317136',
'2346815',
'4057204',
'2210342',
'2618761',
'2770855',
'3670953',
'4056851',
'3114990',
'3427635',
'3895811',
'4056932',
'3646336',
'4023790',
'4070631',
'3006103',
'4021125',
'4084267',
'1573573',
'2466175',
'2264691',
'2408627',
'2609872',
'2870154',
'3052674',
'3058673',
'3068354',
'3186163',
'3507260',
'2837572',
'2903374',
'3036917',
'3225322',
'3236665',
'1618123',
'2269168',
'3912695',
'4106669',
'2303203',
'1016988',
'3743625',
'2250405',
'2150195',
'2367716',
'2510660',
'2946443',
'2184067',
'1020894',
'3537118',
'2200266',
'3741266',
'4022387',
'3662531',
'3866327',
'3212783',
'3437641',
'3085937',
'2475668',
'1442324',
'2500315',
'3132977',
'3945374',
'3000232',
'3912697',
'4005077',
'2670678',
'2333541',
'3889278',
'4086150',
'3977970',
'1405819',
'3809904',
'3972536',
'3972538',
'3237289',
'3941763',
'1531799',
'3060644',
'2921930',
'2926277',
'2926291',
'4147002',
'2451290',
'2462081',
'2559080',
'2699735',
'3940089',
'2920802',
'3779787',
'1382875',
'2093759',
'2151570',
'2218054',
'2411069',
'2136122',
'3655782',
'3894683',
'1205892',
'1779620',
'4167886',
'2389908',
'2396019',
'3468194',
'2926295',
'4038527',
'4195116',
'2343033',
'2212264',
'2511670',
'3524676',
'3947102',
'2167910',
'2220498',
'4150213',
'3489024',
'2928439',
'4022507',
'4157335',
'2347689',
'3919471',
'3972153',
'3966645',
'3192200',
'3613321',
'3234104',
'2620100',
'2615124',
'3595479',
'4226492',
'4071955',
'3919464',
'3919465',
'3234863',
'2843867',
'2142830',
'2250859',
'2453927',
'2987273',
'2774030',
'3873719',
'3103027',
'4191624',
'4043260',
'4132153',
'4132153',
'2702085',
'2725928',
'3237207',
'3585220',
'4258228',
'4215217',
'4264309',
'3214538',
'3628478',
'3160011',
'4157341',
'2469097',
'3984958',
'2182586',
'4178786',
'2479432',
'4285130',
'4285177',
'3905599',
'4281066',
'4281068',
'4281070',
'4281072',
'4281073',
'4281076',
'4281081',
'3628339',
'2246470',
'4068502',
'4209380',
'3947513',
'1478589',
'4254960',
'4295108',
'2612148',
'4309209',
'4148617',
'2964767',
'2750271',
'4248854',
'3961169',
'2596616',
'4301864',
'2302919',
'2333884',
'4309368',
'2311974',
'3426627',
'4280237',
'4157339',
'4320662',
'3052294',
'2186492',
'2268122',
'4296856',
'2209822',
'2958990',
'3831971',
'2082805',
'3830868',
'4187908',
'4318235',
'1015773',
'2380861',
'4343395',
'3866331',
'2293014',
'3984585',
'4355957',
'3854841',
'2402659',
'2669067',
'3891518',
'4380162',
'1815116',
'4043930',
'1496306',
'2505796',
'2343271',
'2195409',
'2247780',
'2259190',
'2195409',
'2247780'
              )
            GROUP BY 1, 2
            HAVING MAX(rmi.creation_date_gmt::DATE) >= CURRENT_DATE - INTERVAL '1 year'
          ) inner_query

          WHERE volume_count != 0
            AND item_count != volume_count
            /*
            Set items to volume fields ratio in order to narrow results to records worth your time to review
            Originally set to 6 (1 volume per six items) and gradually reduced as we worked through backlog of record errors
            */
             AND item_count / volume_count >= 2
           ORDER BY 2
           """
    query_results, headers = run_query(query)

    # generate csv file from those query results
    csv_file = "/Scripts/Volumes/Temp Files/VolumeFields{}.csv".format(date.today())
    local_file = write_csv(query_results, headers, csv_file)

    # send email with attached file
    email_subject = "Volume Fields"
    email_message = """***This is an automated email***
    
    
    The Volume Field Problem report has been attached."""
    send_email(email_subject, email_message, local_file)

    # delete csv file once email has been sent
    os.remove(local_file)


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
        email_subject = "volumes script error"
        email_message = (
            "Your script failed with the following error:\n\n" + traceback.format_exc()
        )

        send_email_error(email_subject, email_message, emailto)
        raise
