"""
Jeremy Goldstein
Minuteman Library Network

Script gathers monthly statistics on checkouts by checkout location and town of residence
Data is appended to a collection of Google Sheets used as data sources by Looker Studio
"""

# run in py313

import configparser
import psycopg2
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email.utils import formatdate
from email import encoders
import traceback
import datetime
import pygsheets


# function takes a sql query as a parameter, connects to a database and returns the results
def runquery(query):
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


# log items that were corrected to an existing Google Sheet
def appendToSheet(spreadSheetId, data):
    if not data:
        raise ValueError("Data must not be empty.")

    gc = pygsheets.authorize(
        service_file="C:\\Scripts\\Creds\\GSheet updater creds.json"
    )

    sh = gc.open_by_key(spreadSheetId)
    wks = sh.sheet1  # or sh.worksheet_by_title("My Sheet")

    # Find the first empty row and insert data there
    first_empty_row = len(wks.get_all_values(include_tailing_empty_rows=False)) + 1
    rows_needed = first_empty_row + len(data) - 1

    # Expand the sheet if the data would exceed the current grid size
    if rows_needed > wks.rows:
        wks.add_rows(rows_needed - wks.rows)
    wks.update_values(f"A{first_empty_row}", data)


# converts psycopg2 fetchall() output to matrix required by pygsheets
def parse_pg_data(rows):

    def convert(val):
        if val is None:
            return ""
        if isinstance(val, (datetime.date, datetime.datetime)):
            return val.isoformat()  # e.g. "2026-03-04"
        return val  # int, float, str pass through as-is

    return [list(convert(val) for val in row) for row in rows]


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
    # enable creds file for referencing GSheets IDs
    config = configparser.ConfigParser()
    config.read("C:\\Scripts\\Creds\\config.ini")

    # query to gather total fines assessed for each checkout location the prior day
    query = """
    SELECT
      *,
      CASE
	    WHEN a.checkout_location = 'ACTON' THEN 'Acton Memorial Library, 486 Main St, Acton, MA 01720'
	    WHEN a.checkout_location = 'ACTON/WEST' THEN 'West Acton Citizens Library, 21 Windsor Ave, Acton, MA 01720'
	    WHEN a.checkout_location = 'ARLINGTON' THEN 'Robbins Library, 700 Massachusetts Ave, Arlington, MA 02476'
	    WHEN a.checkout_location = 'ARLINGTON/FOX' THEN 'Edith M. Fox Branch Library, 175 Massachusetts Ave, Arlington, MA 02474'
	    WHEN a.checkout_location = 'ASHLAND' THEN 'Ashland Public Library, 66 Front St, Ashland, MA 01721'
	    WHEN a.checkout_location = 'BEDFORD' THEN 'Bedford Free Public Library, 7 Mudge Way, Bedford, MA 01730'
	    WHEN a.checkout_location = 'BELMONT' THEN 'Belmont Public Library, 336 Concord Ave, Belmont, MA 02478'
	    WHEN a.checkout_location = 'BROOKLINE' THEN 'Public Library of Brookline, 361 Washington St, Brookline, MA 02445'
	    WHEN a.checkout_location = 'BROOKLINE/COOLIDGE CORNER' THEN 'Public Library of Brookline: Coolidge Corner Branch, 31 Pleasant St, Brookline, MA 02446'
	    WHEN a.checkout_location = 'BROOKLINE/PUTTERHAM' THEN 'Public Library of Brookline: Putterham Branch, 959 W Roxbury Pkwy, Chestnut Hill, MA 02467'
	    WHEN a.checkout_location = 'CAMBRIDGE' THEN 'Cambridge Public Library, 449 Broadway, Cambridge, MA 02138'
	    WHEN a.checkout_location = 'CAMBRIDGE/BOUDREAU' THEN 'Cambridge Public Library: Boudreau Branch, 245 Concord Ave, Cambridge, MA 02138'
	    WHEN a.checkout_location = 'CAMBRIDGE/CENT SQ' THEN 'Cambridge Public Library: Central Square Branch, 45 Pearl St, Cambridge, MA 02139'
	    WHEN a.checkout_location = 'CAMBRIDGE/COLLINS' THEN 'Cambridge Public Library: Collins Branch, 64 Aberdeen Ave, Cambridge, MA 02138'
	    WHEN a.checkout_location = 'CAMBRIDGE/OCONNELL' THEN 'Cambridge Public Library: O''Connell Branch, 48 6th St, Cambridge, MA 02141'
	    WHEN a.checkout_location = 'CAMBRIDGE/ONEILL' THEN 'Cambridge Public Library: O''Neill Branch, 70 Rindge Ave, Cambridge, MA 02140'
	    WHEN a.checkout_location = 'CAMBRIDGE/OUTREACH' THEN 'Cambridge Public Library, 449 Broadway, Cambridge, MA 02138'
	    WHEN a.checkout_location = 'CAMBRIDGE/VALENTE' THEN 'Cambridge Public Library: Valente Branch, 826 Cambridge St, Cambridge, MA 02141'
	    WHEN a.checkout_location = 'CONCORD' THEN 'Concord Free Public Library, 129 Main St, Concord, MA 01742'
	    WHEN a.checkout_location = 'CONCORD/FOWLER' THEN 'Concord Free Public Library: Fowler Branch, 1322 Main St, Concord, MA 01742'
	    WHEN a.checkout_location = 'DEAN COLLEGE' THEN 'E. Ross Anderson Library, 119 Main St, Franklin, MA 02038'
	    WHEN a.checkout_location = 'DEDHAM' THEN 'Dedham Public Library, 43 Church St, Dedham, MA 02026'
	    WHEN a.checkout_location = 'DEDHAM/ENDICOTT' THEN 'Dedham Public Library: Endicott Branch, 257 Mt Vernon St, Dedham, MA 02026'
	    WHEN a.checkout_location = 'DOVER' THEN 'Dover Town Library, 56 Dedham St, Dover, MA 02030'
	    WHEN a.checkout_location = 'FRAMINGHAM' THEN 'Framingham Public Library, 49 Lexington St, Framingham, MA 01702'
	    WHEN a.checkout_location = 'FRAMINGHAM STATE' THEN 'Henry Whittemore Library, 100 State St, Framingham, MA 01702'
	    WHEN a.checkout_location = 'FRAMINGHAM/BKM' THEN 'Framingham Public Library, 49 Lexington St, Framingham, MA 01702'
	    WHEN a.checkout_location = 'FRAMINGHAM/MCAULIFFE' THEN 'Framingham Public Library: Christa McAuliffe Branch, 746 Water St, Framingham, MA 01701'
	    WHEN a.checkout_location = 'FRANKLIN' THEN 'Franklin Public Library, 118 Main St, Franklin, MA 02038'
	    WHEN a.checkout_location = 'HOLLISTON' THEN 'Holliston Public Library, 752 Washington St, Holliston, MA 01746'
	    WHEN a.checkout_location = 'LASELL UNIVERSITY' THEN 'Brennan Library, 80A Maple St, Auburndale, MA 02466'
	    WHEN a.checkout_location = 'LEXINGTON' THEN 'Cary Memorial Library, 1874 Massachusetts Ave, Lexington, MA 02420'
	    WHEN a.checkout_location = 'LINCOLN' THEN 'Lincoln Public Library, 3 Bedford Rd, Lincoln, MA 01773'
	    WHEN a.checkout_location = 'MAYNARD' THEN 'Maynard Public Library, 77 Nason St, Maynard, MA 01754'
	    WHEN a.checkout_location = 'MEDFIELD' THEN 'Medfield Public Library, 468 Main St, Medfield, MA 02052'
	    WHEN a.checkout_location = 'MEDFORD' THEN 'Medford Public Library, 200 Boston Ave suite G-350, Medford, MA 02155'
	    WHEN a.checkout_location = 'MEDWAY' THEN 'Medway Public Library, 26 High St, Medway, MA 02053'
	    WHEN a.checkout_location = 'MILLIS' THEN 'Millis Public Library, 961 Main St, Millis, MA 02054'
	    WHEN a.checkout_location = 'NATICK' THEN 'Morse Institute Library, 14 E Central St, Natick, MA 01760'
	    WHEN a.checkout_location = 'NATICK/BACON' THEN 'Bacon Free Library, 58 Eliot St, Natick, MA 01760'
	    WHEN a.checkout_location = 'NATICK/BKM' THEN 'Morse Institute Library, 14 E Central St, Natick, MA 01760'
	    WHEN a.checkout_location = 'NEEDHAM' THEN 'Needham Free Public Library, 1139 Highland Ave, Needham Heights, MA 02494'
	    WHEN a.checkout_location = 'NEWTON' THEN 'Newton Free Public Library, 330 Homer St, Newton, MA 02459'
	    WHEN a.checkout_location = 'NORWOOD' THEN 'Morrill Memorial Library, 33 Walpole St, Norwood, MA 02062'
	    WHEN a.checkout_location = 'OLIN COLLEGE' THEN 'Olin College Library, 1000 Olin Way, Needham, MA 02492'
	    WHEN a.checkout_location = 'REGIS' THEN 'Regis College Library, 235 Wellesley St, Weston, MA 02493'
	    WHEN a.checkout_location = 'SHERBORN' THEN 'Sherborn Library, 3 Sanger St, Sherborn, MA 01770'
	    WHEN a.checkout_location = 'SOMERVILLE' THEN 'Somerville Public Library, 79 Highland Ave, Somerville, MA 02143'
	    WHEN a.checkout_location = 'SOMERVILLE/EAST' THEN 'Somerville Public Library: East Branch, 115 Broadway, Somerville, MA 02145'
	    WHEN a.checkout_location = 'SOMERVILLE/WEST' THEN 'Somerville Public Library: West Branch, 167 Holland St, Somerville, MA 02144'
	    WHEN a.checkout_location = 'STOW' THEN 'Randall Library, 19 Crescent St, Stow, MA 01775'
	    WHEN a.checkout_location = 'SUDBURY' THEN 'Goodnow Library, 21 Concord Rd, Sudbury, MA 01776'
	    WHEN a.checkout_location = 'WALTHAM' THEN 'Waltham Public Library, 735 Main St, Waltham, MA 02451'
	    WHEN a.checkout_location = 'WATERTOWN' THEN 'Watertown Free Public Library, 123 Main St, Watertown, MA 02472'
	    WHEN a.checkout_location = 'WAYLAND' THEN 'Wayland Free Public Library, 5 Concord Rd, Wayland, MA 01778'
	    WHEN a.checkout_location = 'WELLESLEY' THEN 'Wellesley Free Library, 530 Washington St, Wellesley, MA 02482'
	    WHEN a.checkout_location = 'WELLESLEY/FELLS' THEN 'Wellesley Free Library: Fells Branch, 308 Weston Rd, Wellesley, MA 02482'
	    WHEN a.checkout_location = 'WELLESLEY/HILLS' THEN 'Wellesley Free Library: Hills Branch, 210 Washington St, Wellesley, MA 02482'
	    WHEN a.checkout_location = 'WESTON' THEN 'Weston Public Library, 87 School St, Weston, MA 02493'
	    WHEN a.checkout_location = 'WESTWOOD' THEN 'Westwood Public Library, 660 High St, Westwood, MA 02090'
	    WHEN a.checkout_location = 'WESTWOOD/ISLINGTON' THEN 'Westwood Public Library: Islington Branch, 288 Washington St, Westwood, MA 02090'
	    WHEN a.checkout_location = 'WINCHESTER' THEN 'Winchester Public Library, 80 Washington St, Winchester, MA 01890'
	    WHEN a.checkout_location = 'WOBURN' THEN 'Woburn Public Library, 45 Pleasant St, Woburn, MA 01801'
      END AS address,
      to_char(CURRENT_DATE - 1,'YYYY-MM-DD') AS "date"

    FROM (
      SELECT
        CASE
	      WHEN p3.name IN ('-', 'Out of state', 'ComCat', 'ILL', 'Other MA') THEN ''
	      WHEN p3.name = 'Fram. State' THEN 'Framingham MA'
	      WHEN p3.name = 'Dean College' THEN 'Franklin MA'
	      WHEN p3.name = 'Lasell University' THEN 'Newton MA'
	      WHEN p3.name = 'Olin College' THEN 'Needham MA'
	      WHEN p3.name = 'Regis College' THEN 'Weston MA'	
	      WHEN p3.name = 'Pine Manor College' THEN 'Newton MA'
	      ELSE p3.name||' MA'
        END AS ma_town,
        l.name AS checkout_location,
        COUNT(t.id) AS checkout_total
	
      FROM sierra_view.circ_trans t
      JOIN sierra_view.patron_record p
        ON t.patron_record_id = p.id
        AND t.op_code = 'o'
      JOIN sierra_view.user_defined_pcode3_myuser p3
        ON p.pcode3::varchar = p3.code
      JOIN sierra_view.statistic_group_myuser S
        ON t.stat_group_code_num = S.code
      JOIN sierra_view.location_myuser l
        ON S.location_code = l.code
      JOIN sierra_view.patron_record_address a
        ON p.id = a.patron_record_id

      WHERE t.transaction_gmt::DATE < CURRENT_DATE AND t.transaction_gmt::DATE >= (CURRENT_DATE - INTERVAL '1 month')

      GROUP BY 1,2
    )a
    WHERE a.ma_town != ''
    ORDER BY 1,2
    """

    # run query and append results to specified GSheet
    results = runquery(query)
    parsed_results = parse_pg_data(results)
    appendToSheet(config["gsheet"]["checkouts_by_residence"], parsed_results)


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
        email_subject = "checkouts by residence script error"
        email_message = (
            "Your script failed with the following error:\n\n" + traceback.format_exc()
        )

        send_email_error(email_subject, email_message, emailto)
        raise
