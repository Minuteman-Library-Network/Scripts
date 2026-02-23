"""
Jeremy Goldstein
Minuteman Library Network

Script gathers daily statistics and transaction counts used by Minuteman's circulation data dashboards
Data is appended to a collection of Google Sheets used as data sources by Looker Studio
"""

# run in py38

import configparser
import psycopg2
import os
from oauth2client.service_account import ServiceAccountCredentials
from googleapiclient.discovery import build
import gspread
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email.utils import formatdate
from email import encoders
import traceback


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
    scopes = [
        "https://spreadsheets.google.com/feeds",
        "https://www.googleapis.com/auth/drive",
    ]
    creds = ServiceAccountCredentials.from_json_keyfile_name(
        "C:\\Scripts\\Creds\\GSheet updater creds.json", scopes
    )
    service = build("sheets", "v4", credentials=creds)
    sheet = service.spreadsheets()
    request = (
        service.spreadsheets()
        .values()
        .append(
            spreadsheetId=spreadSheetId,
            range="A1:Z1",
            valueInputOption="USER_ENTERED",
            body={"values": data},
        )
    )
    result = request.execute()


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
    fines_assessed_query = """
      SELECT
        TO_CHAR(CURRENT_DATE - INTERVAL '1 day','YYYY-MM-DD') AS DATE,
        CASE
          WHEN f.loanrule_code_num BETWEEN 2 AND 12 OR f.loanrule_code_num BETWEEN 501 AND 509 THEN 'Acton'
          WHEN f.loanrule_code_num BETWEEN 13 AND 23 OR f.loanrule_code_num BETWEEN 510 AND 518 THEN 'Arlington'
          WHEN f.loanrule_code_num BETWEEN 24 AND 34 OR f.loanrule_code_num BETWEEN 519 AND 527 THEN 'Ashland'
          WHEN f.loanrule_code_num BETWEEN 35 AND 45 OR f.loanrule_code_num BETWEEN 528 AND 536 THEN 'Bedford'
          WHEN f.loanrule_code_num BETWEEN 46 AND 56 OR f.loanrule_code_num BETWEEN 537 AND 545 THEN 'Belmont'
          WHEN f.loanrule_code_num BETWEEN 57 AND 67 OR f.loanrule_code_num BETWEEN 546 AND 554 THEN 'Brookline'
          WHEN f.loanrule_code_num BETWEEN 68 AND 78 OR f.loanrule_code_num BETWEEN 555 AND 563 THEN 'Cambridge'
          WHEN f.loanrule_code_num BETWEEN 79 AND 89 OR f.loanrule_code_num BETWEEN 564 AND 572 THEN 'Concord'
          WHEN f.loanrule_code_num BETWEEN 90 AND 100 OR f.loanrule_code_num BETWEEN 573 AND 581 THEN 'Dedham'
          WHEN f.loanrule_code_num BETWEEN 101 AND 111 OR f.loanrule_code_num BETWEEN 582 AND 590 THEN 'Dean'
          WHEN f.loanrule_code_num BETWEEN 112 AND 122 OR f.loanrule_code_num BETWEEN 591 AND 599 THEN 'Dover'
          WHEN f.loanrule_code_num BETWEEN 123 AND 133 OR f.loanrule_code_num BETWEEN 600 AND 608 THEN 'Framingham'
          WHEN f.loanrule_code_num BETWEEN 134 AND 144 OR f.loanrule_code_num BETWEEN 609 AND 617 THEN 'Franklin'
          WHEN f.loanrule_code_num BETWEEN 145 AND 155 OR f.loanrule_code_num BETWEEN 618 AND 626 THEN 'Framingham State'
          WHEN f.loanrule_code_num BETWEEN 156 AND 166 OR f.loanrule_code_num BETWEEN 627 AND 635 THEN 'Holliston'
          WHEN f.loanrule_code_num BETWEEN 167 AND 177 OR f.loanrule_code_num BETWEEN 636 AND 644 THEN 'Lasell'
          WHEN f.loanrule_code_num BETWEEN 178 AND 188 OR f.loanrule_code_num BETWEEN 645 AND 653 THEN 'Lexington'
          WHEN f.loanrule_code_num BETWEEN 189 AND 199 OR f.loanrule_code_num BETWEEN 654 AND 662 THEN 'Lincoln'
          WHEN f.loanrule_code_num BETWEEN 200 AND 210 OR f.loanrule_code_num BETWEEN 663 AND 671 THEN 'Maynard'
          WHEN f.loanrule_code_num BETWEEN 222 AND 232 OR f.loanrule_code_num BETWEEN 681 AND 689 THEN 'Medford'
          WHEN f.loanrule_code_num BETWEEN 233 AND 243 OR f.loanrule_code_num BETWEEN 690 AND 698 THEN 'Millis'
          WHEN f.loanrule_code_num BETWEEN 244 AND 254 OR f.loanrule_code_num BETWEEN 699 AND 707 THEN 'Medfield'
          WHEN f.loanrule_code_num BETWEEN 255 AND 265 OR f.loanrule_code_num BETWEEN 708 AND 716 THEN 'Mount Ida'
          WHEN f.loanrule_code_num BETWEEN 266 AND 276 OR f.loanrule_code_num BETWEEN 717 AND 725 THEN 'Medway'
          WHEN f.loanrule_code_num BETWEEN 277 AND 287 OR f.loanrule_code_num BETWEEN 726 AND 733 THEN 'Natick'
          WHEN f.loanrule_code_num BETWEEN 289 AND 298 OR f.loanrule_code_num BETWEEN 734 AND 743 THEN 'Olin'
          WHEN f.loanrule_code_num BETWEEN 299 AND 309 OR f.loanrule_code_num BETWEEN 744 AND 752 THEN 'Needham'
          WHEN f.loanrule_code_num BETWEEN 310 AND 320 OR f.loanrule_code_num BETWEEN 753 AND 761 THEN 'Norwood'
          WHEN f.loanrule_code_num BETWEEN 321 AND 331 OR f.loanrule_code_num BETWEEN 762 AND 770 THEN 'Newton'
          WHEN f.loanrule_code_num BETWEEN 332 AND 342 OR f.loanrule_code_num BETWEEN 771 AND 779 THEN 'Somerville'
          WHEN f.loanrule_code_num BETWEEN 343 AND 353 OR f.loanrule_code_num BETWEEN 780 AND 788 THEN 'Stow'
          WHEN f.loanrule_code_num BETWEEN 354 AND 364 OR f.loanrule_code_num BETWEEN 789 AND 797 THEN 'Sudbury'
          WHEN f.loanrule_code_num BETWEEN 365 AND 375 OR f.loanrule_code_num BETWEEN 798 AND 806 THEN 'Watertown'
          WHEN f.loanrule_code_num BETWEEN 376 AND 386 OR f.loanrule_code_num BETWEEN 807 AND 815 THEN 'Wellesley'
          WHEN f.loanrule_code_num BETWEEN 387 AND 397 OR f.loanrule_code_num BETWEEN 816 AND 824 THEN 'Winchester'
          WHEN f.loanrule_code_num BETWEEN 398 AND 408 OR f.loanrule_code_num BETWEEN 825 AND 833 THEN 'Waltham'
          WHEN f.loanrule_code_num BETWEEN 409 AND 419 OR f.loanrule_code_num BETWEEN 834 AND 842 THEN 'Woburn'
          WHEN f.loanrule_code_num BETWEEN 420 AND 430 OR f.loanrule_code_num BETWEEN 843 AND 851 THEN 'Weston'
          WHEN f.loanrule_code_num BETWEEN 431 AND 441 OR f.loanrule_code_num BETWEEN 852 AND 860 THEN 'Westwood'
          WHEN f.loanrule_code_num BETWEEN 442 AND 452 OR f.loanrule_code_num BETWEEN 861 AND 869 THEN 'Wayland'
          WHEN f.loanrule_code_num BETWEEN 453 AND 463 OR f.loanrule_code_num BETWEEN 870 AND 878 THEN 'Pine Manor'
          WHEN f.loanrule_code_num BETWEEN 464 AND 474 OR f.loanrule_code_num BETWEEN 879 AND 887 THEN 'Regis'
          WHEN f.loanrule_code_num BETWEEN 475 AND 485 OR f.loanrule_code_num BETWEEN 888 AND 896 THEN 'Sherborn'
          ELSE 'Other'
        END AS checkout_location,
        COUNT(DISTINCT f.id) AS count_fines_assessed,
        SUM(f.item_charge_amt + f.processing_fee_amt + f.billing_fee_amt) FILTER (WHERE f.charge_code IN ('1','2','3','4','5','6'))::MONEY AS total_fines_assessed

      FROM sierra_view.fine f

      WHERE f.assessed_gmt::DATE = CURRENT_DATE - INTERVAL '1 day'

      GROUP BY 1,2
      ORDER BY 1,2
      """

    # run query and append results to specified GSheet
    fines_assessed = runquery(fines_assessed_query)
    appendToSheet(config["gsheet"]["fines_assessed"], fines_assessed)

    # repeat steps for each dataset

    # query to gather totals for fines paid the previous day
    fines_paid_query = """
      SELECT
        TO_CHAR(CURRENT_DATE - INTERVAL '1 day','YYYY-MM-DD') AS paid_date,
        CASE
          WHEN fp.loan_rule_code_num BETWEEN 2 AND 12 OR fp.loan_rule_code_num BETWEEN 501 AND 509 THEN 'Acton'
          WHEN fp.loan_rule_code_num BETWEEN 13 AND 23 OR fp.loan_rule_code_num BETWEEN 510 AND 518 THEN 'Arlington'
          WHEN fp.loan_rule_code_num BETWEEN 24 AND 34 OR fp.loan_rule_code_num BETWEEN 519 AND 527 THEN 'Ashland'
          WHEN fp.loan_rule_code_num BETWEEN 35 AND 45 OR fp.loan_rule_code_num BETWEEN 528 AND 536 THEN 'Bedford'
          WHEN fp.loan_rule_code_num BETWEEN 46 AND 56 OR fp.loan_rule_code_num BETWEEN 537 AND 545 THEN 'Belmont'
          WHEN fp.loan_rule_code_num BETWEEN 57 AND 67 OR fp.loan_rule_code_num BETWEEN 546 AND 554 THEN 'Brookline'
          WHEN fp.loan_rule_code_num BETWEEN 68 AND 78 OR fp.loan_rule_code_num BETWEEN 555 AND 563 THEN 'Cambridge'
          WHEN fp.loan_rule_code_num BETWEEN 79 AND 89 OR fp.loan_rule_code_num BETWEEN 564 AND 572 THEN 'Concord'
          WHEN fp.loan_rule_code_num BETWEEN 90 AND 100 OR fp.loan_rule_code_num BETWEEN 573 AND 581 THEN 'Dedham'
          WHEN fp.loan_rule_code_num BETWEEN 101 AND 111 OR fp.loan_rule_code_num BETWEEN 582 AND 590 THEN 'Dean'
          WHEN fp.loan_rule_code_num BETWEEN 112 AND 122 OR fp.loan_rule_code_num BETWEEN 591 AND 599 THEN 'Dover'
          WHEN fp.loan_rule_code_num BETWEEN 123 AND 133 OR fp.loan_rule_code_num BETWEEN 600 AND 608 THEN 'Framingham'
          WHEN fp.loan_rule_code_num BETWEEN 134 AND 144 OR fp.loan_rule_code_num BETWEEN 609 AND 617 THEN 'Franklin'
          WHEN fp.loan_rule_code_num BETWEEN 145 AND 155 OR fp.loan_rule_code_num BETWEEN 618 AND 626 THEN 'Framingham State'
          WHEN fp.loan_rule_code_num BETWEEN 156 AND 166 OR fp.loan_rule_code_num BETWEEN 627 AND 635 THEN 'Holliston'
          WHEN fp.loan_rule_code_num BETWEEN 167 AND 177 OR fp.loan_rule_code_num BETWEEN 636 AND 644 THEN 'Lasell'
          WHEN fp.loan_rule_code_num BETWEEN 178 AND 188 OR fp.loan_rule_code_num BETWEEN 645 AND 653 THEN 'Lexington'
          WHEN fp.loan_rule_code_num BETWEEN 189 AND 199 OR fp.loan_rule_code_num BETWEEN 654 AND 662 THEN 'Lincoln'
          WHEN fp.loan_rule_code_num BETWEEN 200 AND 210 OR fp.loan_rule_code_num BETWEEN 663 AND 671 THEN 'Maynard'
          WHEN fp.loan_rule_code_num BETWEEN 222 AND 232 OR fp.loan_rule_code_num BETWEEN 681 AND 689 THEN 'Medford'
          WHEN fp.loan_rule_code_num BETWEEN 233 AND 243 OR fp.loan_rule_code_num BETWEEN 690 AND 698 THEN 'Millis'
          WHEN fp.loan_rule_code_num BETWEEN 244 AND 254 OR fp.loan_rule_code_num BETWEEN 699 AND 707 THEN 'Medfield'
          WHEN fp.loan_rule_code_num BETWEEN 255 AND 265 OR fp.loan_rule_code_num BETWEEN 708 AND 716 THEN 'Mount Ida'
          WHEN fp.loan_rule_code_num BETWEEN 266 AND 276 OR fp.loan_rule_code_num BETWEEN 717 AND 725 THEN 'Medway'
          WHEN fp.loan_rule_code_num BETWEEN 277 AND 287 OR fp.loan_rule_code_num BETWEEN 726 AND 734 THEN 'Natick'
          WHEN fp.loan_rule_code_num BETWEEN 289 AND 298 OR fp.loan_rule_code_num BETWEEN 734 AND 743 THEN 'Olin'
          WHEN fp.loan_rule_code_num BETWEEN 299 AND 309 OR fp.loan_rule_code_num BETWEEN 744 AND 752 THEN 'Needham'
          WHEN fp.loan_rule_code_num BETWEEN 310 AND 320 OR fp.loan_rule_code_num BETWEEN 753 AND 761 THEN 'Norwood'
          WHEN fp.loan_rule_code_num BETWEEN 321 AND 331 OR fp.loan_rule_code_num BETWEEN 762 AND 770 THEN 'Newton'
          WHEN fp.loan_rule_code_num BETWEEN 332 AND 342 OR fp.loan_rule_code_num BETWEEN 771 AND 779 THEN 'Somerville'
          WHEN fp.loan_rule_code_num BETWEEN 343 AND 353 OR fp.loan_rule_code_num BETWEEN 780 AND 788 THEN 'Stow'
          WHEN fp.loan_rule_code_num BETWEEN 354 AND 364 OR fp.loan_rule_code_num BETWEEN 789 AND 797 THEN 'Sudbury'
          WHEN fp.loan_rule_code_num BETWEEN 365 AND 375 OR fp.loan_rule_code_num BETWEEN 798 AND 806 THEN 'Watertown'
          WHEN fp.loan_rule_code_num BETWEEN 376 AND 386 OR fp.loan_rule_code_num BETWEEN 807 AND 815 THEN 'Wellesley'
          WHEN fp.loan_rule_code_num BETWEEN 387 AND 397 OR fp.loan_rule_code_num BETWEEN 816 AND 824 THEN 'Winchester'
          WHEN fp.loan_rule_code_num BETWEEN 398 AND 408 OR fp.loan_rule_code_num BETWEEN 825 AND 833 THEN 'Waltham'
          WHEN fp.loan_rule_code_num BETWEEN 409 AND 419 OR fp.loan_rule_code_num BETWEEN 834 AND 842 THEN 'Woburn'
          WHEN fp.loan_rule_code_num BETWEEN 420 AND 430 OR fp.loan_rule_code_num BETWEEN 843 AND 851 THEN 'Weston'
          WHEN fp.loan_rule_code_num BETWEEN 431 AND 441 OR fp.loan_rule_code_num BETWEEN 852 AND 860 THEN 'Westwood'
          WHEN fp.loan_rule_code_num BETWEEN 442 AND 452 OR fp.loan_rule_code_num BETWEEN 861 AND 869 THEN 'Wayland'
          WHEN fp.loan_rule_code_num BETWEEN 453 AND 463 OR fp.loan_rule_code_num BETWEEN 870 AND 878 THEN 'Pine Manor'
          WHEN fp.loan_rule_code_num BETWEEN 464 AND 474 OR fp.loan_rule_code_num BETWEEN 879 AND 887 THEN 'Regis'
          WHEN fp.loan_rule_code_num BETWEEN 475 AND 485 OR fp.loan_rule_code_num BETWEEN 888 AND 896 THEN 'Sherborn'
          Else 'Other'
        END AS checkout_location,
        CASE
          WHEN fp.charge_type_code = '1' THEN 'Manual Charge'
          WHEN fp.charge_type_code IN ('2','4','6') THEN 'Overdue'
          WHEN fp.charge_type_code = '3' THEN 'Replacement'
          WHEN fp.charge_type_code = '5' THEN 'Lost Book'
        END AS charge_type,
        CASE
          WHEN fp.payment_type_code = 'e' THEN true
          ELSE false
        END AS paid_online,
        COUNT(fp.id) FILTER (WHERE fp.payment_status_code NOT IN ('3','0')) AS fines_paid_count,
        COUNT(fp.id) FILTER (WHERE fp.payment_status_code = '3') AS fines_waived_count,
        COUNT(fp.id) FILTER (WHERE fp.payment_status_code = '0') AS fines_removed_count,
        COALESCE(SUM(
          CASE
            WHEN fp.payment_status_code = '2' THEN fp.paid_now_amt
            ELSE fp.item_charge_amt + fp.processing_fee_amt + fp.processing_fee_amt
          END
        ) FILTER (WHERE fp.payment_status_code != '3'),'0')::MONEY AS fines_paid_total,
        COALESCE(SUM(fp.item_charge_amt + fp.processing_fee_amt + fp.processing_fee_amt) FILTER (WHERE fp.payment_status_code = '3'),'0')::MONEY AS fines_waived_total,
        COALESCE(SUM(fp.item_charge_amt + fp.processing_fee_amt + fp.processing_fee_amt) FILTER (WHERE fp.payment_status_code = '0'),'0')::MONEY AS fines_removed_total

      FROM sierra_view.fines_paid fp

      WHERE fp.paid_date_gmt::DATE = CURRENT_DATE - INTERVAL '1 day'
        AND fp.charge_type_code IN ('1','2','3','4','5','6')

      GROUP BY 1,2,3,4
      ORDER BY 1,2,3,4
      """

    fines_paid = runquery(fines_paid_query)
    appendToSheet(config["gsheet"]["fines_paid"], fines_paid)

    # query to gather counts of items coming to, leaving from, or sitting on the holdshelf at each location
    library_transit_counts_query = """
      WITH transit AS (
        SELECT
          v.id AS id,
          SUBSTRING(SPLIT_PART(SPLIT_PART(v.field_content,'from ',2),' to',1)FROM 1 FOR 3) AS origin_loc,
          SUBSTRING(SPLIT_PART(v.field_content,'to ',2) FROM 1 FOR 3) AS destination_loc
        FROM sierra_view.item_record i
        JOIN sierra_view.varfield v
        ON i.id = v.record_id
          AND v.varfield_type_code = 'm'
          AND v.field_content LIKE '%IN TRANSIT%'
        WHERE i.item_status_code = 't'
      )

      SELECT
        TO_CHAR(CURRENT_DATE,'YYYY-MM-DD') AS "date",
        l.name AS library,
        COUNT(DISTINCT t.id) FILTER (WHERE t.origin_loc = l.code) AS transit_from,
        COUNT(DISTINCT t.id) FILTER (WHERE t.destination_loc = l.code) AS transit_to,
        (
          SELECT 
            COUNT(h.id) AS on_holdshelf
          FROM sierra_view.hold h
          WHERE SUBSTRING(h.pickup_location_code,1,3) = l.code
            AND h.status IN ('b','i')
        )

      FROM sierra_view.location_myuser l
      JOIN transit t
      ON (l.code = t.origin_loc OR l.code = t.destination_loc)
        AND l.code NOT IN ('trn','mti')

      GROUP BY 1,2,l.code
      ORDER BY 1,2
      """
    library_transit_counts = runquery(library_transit_counts_query)
    appendToSheet(config["gsheet"]["library_transit_counts"], library_transit_counts)

    # query gathers daily counts of each transaction type at each stat group
    circ_trans_snapshot_query = """
      SELECT
        CASE
	        WHEN C.op_code = 'i' AND C.transaction_gmt::TIME = '04:00:00' THEN to_char(i.last_checkin_gmt,'YYYY-MM-DD')
	        ELSE to_char(c.transaction_gmt,'YYYY-MM-DD')
        END AS DATE,
        CASE
	      --account for internal use counts made via mobile worklists
	        WHEN c.stat_group_code_num IN (992,0) AND c.op_code = 'u' AND u.name IS NOT NULL THEN u.statistic_group_code_num
	        ELSE c.stat_group_code_num 
        END AS stat_group,
        COUNT(DISTINCT c.id) FILTER(WHERE c.op_code = 'o') AS checkouts,
        COUNT(DISTINCT c.id) FILTER(WHERE c.op_code = 'i') AS checkins,
        COUNT(DISTINCT c.id) FILTER(WHERE c.op_code = 'r') AS renewals,
        COUNT(DISTINCT c.id) FILTER(WHERE c.op_code = 'u') AS use_count,
        COUNT(DISTINCT c.id) FILTER(WHERE c.op_code = 'f') AS filled_hold,
        COUNT(DISTINCT c.id) FILTER(WHERE c.op_code IN ('n','nb','ni','h','hb','hi')) AS hold_placed

      FROM sierra_view.circ_trans C
      LEFT JOIN sierra_view.item_record i
        ON C.item_record_id = i.id
      --account for internal use counts made via mobile worklists
      LEFT JOIN sierra_view.iii_user u
        ON SUBSTRING(i.location_code,1,3) = SUBSTRING(u.name,1,3) AND u.name ~ 'lists'

      WHERE (c.transaction_gmt::DATE >= CURRENT_DATE - INTERVAL '1 day'
        AND c.transaction_gmt::DATE != CURRENT_DATE)
        --accomodate backdating
        OR (C.op_code = 'i' AND C.transaction_gmt::TIME = '04:00:00' AND i.last_checkin_gmt::DATE >= CURRENT_DATE - INTERVAL '1 day'
        AND i.last_checkin_gmt::DATE != CURRENT_DATE)

      GROUP BY 1,2
      ORDER BY 1,2
    """
    circ_trans_snapshot = runquery(circ_trans_snapshot_query)
    appendToSheet(config["gsheet"]["circ_trans_snapshot"], circ_trans_snapshot)

    # query gathers count of unique patrons to place a hold or checkout at each location each day
    unique_patron_counts_query = """
      SELECT
        TO_CHAR(c.transaction_gmt, 'MM-DD-YY') AS "date",
        l.name AS library,
        COUNT(DISTINCT c.id) AS checkouts_items,
        COUNT(DISTINCT c.patron_record_id) AS checkouts_patrons,
        COUNT(DISTINCT h.id) AS holds_placed_items,
        COUNT(DISTINCT h.patron_record_id) AS holds_placed_patrons

      FROM sierra_view.circ_trans c
      JOIN sierra_view.statistic_group_myuser s
        ON c.stat_group_code_num = s.code
      JOIN sierra_view.location_myuser l
        ON s.location_code = l.code
      LEFT JOIN sierra_view.hold h
        ON s.location_code = SUBSTRING(h.pickup_location_code,1,3)
        AND c.transaction_gmt::DATE = h.placed_gmt::DATE

      WHERE c.op_code = 'o'
        AND c.transaction_gmt::DATE = CURRENT_DATE - INTERVAL '1 day'

      GROUP BY 1,2
      ORDER BY 1,2
      """

    unique_patron_counts = runquery(unique_patron_counts_query)
    appendToSheet(config["gsheet"]["unique_patron_counts"], unique_patron_counts)

    # query gathers the total value of items checked out at each stat group daily by itype and mat type
    checkout_value_query = """
      SELECT
        TO_CHAR(o.transaction_gmt,'YYYY-MM-DD') AS "Date",
        SUBSTRING(s.name,1,3) AS "Stat Group",
        loc.name AS Library,
        it.name AS iType,
        m.name AS "Mat Type",
        SUM(i.price)::MONEY AS VALUE,
        COUNT(DISTINCT o.id) AS checkouts

      FROM sierra_view.circ_trans o
      JOIN sierra_view.item_record i
        ON o.item_record_id = i.id
      JOIN sierra_view.bib_record_item_record_link l
        ON i.id = l.item_record_id
      JOIN sierra_view.bib_record_property b
        ON l.bib_record_id = b.bib_record_id
      JOIN sierra_view.statistic_group_myuser s
        ON i.checkout_statistic_group_code_num = s.code
      JOIN sierra_view.location_myuser loc
        ON SUBSTRING(s.name,1,3) = loc.code
      JOIN sierra_view.itype_property_myuser it
        ON i.itype_code_num = it.code
      JOIN sierra_view.material_property_myuser m
        ON b.material_code = m.code

      WHERE o.op_code = 'o'
        AND o.transaction_gmt::DATE = CURRENT_DATE - INTERVAL '1 day'

      GROUP BY 1,2,3,4,5
      """
    checkout_value = runquery(checkout_value_query)
    appendToSheet(config["gsheet"]["checkout_value"], checkout_value)

    # query to gather daily counts of holds on hold, in transit or on holdshelf at each location
    daily_holds_query = """
      SELECT
        TO_CHAR(CURRENT_DATE,'YYYY-MM-DD') AS "date",
        l.name AS library,
        CASE
          WHEN h.status = '0' THEN 'on hold'
          WHEN h.status = 't' THEN 'in transit'
          ELSE 'on holdshelf'
        END AS status,
        COUNT(h.id)

      FROM sierra_view.hold h
      JOIN sierra_view.location_myuser l
        ON SUBSTRING(h.pickup_location_code,1,3) = l.code

      GROUP BY 1,2,3
      ORDER BY 2,3
      """

    daily_holds = runquery(daily_holds_query)
    appendToSheet(config["gsheet"]["daily_holds"], daily_holds)

    # query to gather hourly transaction counts at each location
    hourly_transactions_query = """
      SELECT
        TO_CHAR(t.transaction_gmt,'YYYY-MM-DD') AS "Date",
        EXTRACT(HOUR FROM t.transaction_gmt)::INT AS "Hour",
        CASE
	        WHEN t.op_code = 'i' THEN 'Check In'
	        WHEN t.op_code = 'o' THEN 'Check Out'
	        WHEN t.op_code = 'f' THEN 'Filled Hold'
	        WHEN t.op_code = 'r' THEN 'Renewal'
	        ELSE 'Hold Placed'
        END AS "Transaction Type",
        CASE
	        WHEN s.name IS NULL THEN t.stat_group_code_num::VARCHAR
	        ELSE s.name
        END AS "Stat Group",
        COUNT(t.id) AS "Total"

      FROM sierra_view.circ_trans t
      LEFT JOIN sierra_view.statistic_group_myuser s
        ON t.stat_group_code_num = s.code
      WHERE t.transaction_gmt::DATE = CURRENT_DATE - INTERVAL '1 day'

      GROUP BY 1,2,3,4

      UNION

      SELECT
        TO_CHAR(t.transaction_gmt,'YYYY-MM-DD') AS "Date",
        EXTRACT(HOUR FROM t.transaction_gmt)::INT AS "Hour",
        'Non-Hold Checkout' AS "Transaction Type",
        CASE
	        WHEN s.name IS NULL THEN t.stat_group_code_num::VARCHAR
	        ELSE s.name
        END AS "Stat Group",
        CASE
	        WHEN COUNT(t.id) FILTER(WHERE t.op_code = 'o') = 0 THEN COUNT(t.id) FILTER(WHERE t.op_code = 'f')
	        WHEN COUNT(t.id) FILTER(WHERE t.op_code = 'o') - COUNT(t.id) FILTER(WHERE t.op_code = 'f') < 0 THEN 0
	        ELSE COUNT(t.id) FILTER(WHERE t.op_code = 'o') - COUNT(t.id) FILTER(WHERE t.op_code = 'f') 
        END AS "Total"

      FROM sierra_view.circ_trans t
      LEFT JOIN sierra_view.statistic_group_myuser s
        ON t.stat_group_code_num = s.code

      WHERE t.transaction_gmt::DATE = CURRENT_DATE - INTERVAL '1 day'
        AND t.op_code IN ('o','f')

      GROUP BY 1,2,3,4
      ORDER BY 1,2,3,4,5
      """
    hourly_transactions = runquery(hourly_transactions_query)
    appendToSheet(config["gsheet"]["hourly_transactions"], hourly_transactions)

    # gathers daily count of how many online registrations occured for each library
    online_registrations_query = """
      SELECT
        TO_CHAR(r.creation_date_gmt::DATE,'YYYY-MM-DD'),
        pt.name,
        COUNT(p.id)

      FROM sierra_view.patron_record p
      JOIN sierra_view.record_metadata r
        ON p.id = r.id
        AND p.patron_agency_code_num = '47'
        AND r.creation_date_gmt::DATE = CURRENT_DATE - INTERVAL '1 day'
      JOIN sierra_view.ptype_property_myuser pt
        ON p.ptype_code = pt.value

      GROUP BY 1,2
      ORDER BY 1,2
      """
    online_registrations = runquery(online_registrations_query)
    appendToSheet(config["gsheet"]["online_registrations"], online_registrations)

    # Gathers daily totals of items in transit
    in_transit_snapshot_query = """
      SELECT
        --date is offset by 1 to account for query runtime having shifted from before midnight to after
        TO_CHAR(CURRENT_DATE - INTERVAL '1 DAY','YYYY-MM-DD') AS transit_date,
        COUNT(i.id) FILTER(WHERE i.item_status_code = 't') AS count,
        COUNT(i.id) FILTER(WHERE i.item_status_code = 't' AND h.id IS NOT NULL) AS count_hold,
        COUNT(i.id) FILTER(WHERE i.item_status_code = 't' AND h.id IS NULL) AS count_return,
        COUNT(h.id) FILTER(WHERE h.status IN ('b','i')) AS count_holdshelf

      FROM sierra_view.item_record i
      LEFT JOIN sierra_view.hold h
        ON i.id = h.record_id
      """
    in_transit_snapshot = runquery(in_transit_snapshot_query)
    appendToSheet(config["gsheet"]["in_transit_snapshot"], in_transit_snapshot)


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
        email_subject = "circulation dashboard script error"
        email_message = (
            "Your script failed with the following error:\n\n" + traceback.format_exc()
        )

        send_email_error(email_subject, email_message, emailto)
        raise
