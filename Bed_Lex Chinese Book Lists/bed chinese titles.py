#!/usr/bin/env python3
# run in py338

"""
Jeremy Goldstein
Minuteman Library Network

Generates custom Chinese title lists per specifications of Lexington
Originally requested by Shiouh-Lin Chang
"""

import psycopg2
import xlsxwriter
import os
import configparser
import sys
import time
from datetime import date
import configparser


# connect to Sierra-db and store results of an sql query
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
    workbook = xlsxwriter.Workbook(excel_file, {"remove_timezone": True})

    worksheet = workbook.add_worksheet("New Titles")

    # Formatting our Excel worksheet
    worksheet.set_landscape()
    worksheet.hide_gridlines(0)

    # Formatting Cells
    eformat = workbook.add_format(
        {
            "text_wrap": True,
            "valign": "bottom",
            "font_name": "Arial",
            "font_size": "10",
            "top": 1,
            "bottom": 1,
        }
    )
    eformaturl = workbook.add_format(
        {
            "text_wrap": True,
            "valign": "bottom",
            "font_name": "Arial",
            "font_size": "10",
            "font_color": "blue",
            "top": 1,
            "bottom": 1,
            "right": 1,
        }
    )
    eformatlabel = workbook.add_format(
        {
            "text_wrap": True,
            "valign": "bottom",
            "bold": True,
            "font_name": "Arial",
            "font_size": "10",
            "top": 1,
            "bottom": 1,
        }
    )
    eformatdate = workbook.add_format(
        {
            "text_wrap": True,
            "valign": "bottom",
            "font_name": "Arial",
            "font_size": "10",
            "top": 1,
            "bottom": 1,
            "num_format": "mm/dd/yy",
        }
    )

    # includes commented out lines from more detailed report version
    # Setting the column widths
    worksheet.set_column(0, 0, 25.71)
    worksheet.set_column(1, 1, 14)
    worksheet.set_column(2, 2, 80.43)
    worksheet.set_column(3, 3, 36.29)
    # worksheet.set_column(3,3,20)
    # worksheet.set_column(4,4,36.29)
    # worksheet.set_column(5,5,80.43)
    worksheet.set_column(4, 4, 14)
    worksheet.set_column(5, 5, 9)
    worksheet.set_column(6, 6, 9)
    worksheet.set_column(7, 7, 13.33)
    worksheet.set_column(8, 8, 13.33)
    # worksheet.set_column(12,12,13.33)
    worksheet.set_column(9, 9, 13.33)

    # Inserting a header
    worksheet.set_header("Lex New Titles")

    # Adding column labels
    worksheet.write(0, 0, "Call Number", eformatlabel)
    worksheet.write(0, 1, "Barcode", eformatlabel)
    worksheet.write(0, 2, "Title", eformatlabel)
    worksheet.write(0, 3, "Author", eformatlabel)
    # worksheet.write(0,3,'ISBN', eformatlabel)
    # worksheet.write(0,4,'Author_English', eformatlabel)
    # worksheet.write(0,5,'Title_English', eformatlabel)
    worksheet.write(0, 4, "Checkout_Total", eformatlabel)
    worksheet.write(0, 5, "YTD_Circ", eformatlabel)
    worksheet.write(0, 6, "LYR_Circ", eformatlabel)
    worksheet.write(0, 7, "Created_Date", eformatlabel)
    worksheet.write(0, 8, "Last_Checkin", eformatlabel)
    # worksheet.write(0,12,'Due_Date', eformatlabel)
    worksheet.write(0, 9, "Status", eformatlabel)

    # Writing the report for staff to the Excel worksheet
    row0 = 1

    for rownum, row in enumerate(query_results):
        worksheet.write(row0, 0, row[0], eformat)
        worksheet.write(row0, 1, row[7], eformat)
        worksheet.write_url(row0, 2, row[5], eformaturl, string=row[1])
        worksheet.write(row0, 3, row[3], eformat)
        # worksheet.write(row0,3,row[6], eformat)
        # worksheet.write(row0,4,row[4], eformat)
        # worksheet.write(row0,5,row[2], eformat)
        worksheet.write(row0, 4, row[8], eformat)
        worksheet.write(row0, 5, row[9], eformat)
        worksheet.write(row0, 6, row[10], eformat)
        worksheet.write(row0, 7, row[11], eformatdate)
        worksheet.write(row0, 8, row[12], eformatdate)
        # worksheet.write(row0,12,row[13], eformat)
        worksheet.write(row0, 9, row[14], eformat)
        row0 += 1

    workbook.close()


def main(file, date_limit, scat_code):

    query = (
        """
    SELECT
      DISTINCT(REGEXP_REPLACE(i.call_number,'\|[a-z]',' ','g')) AS call_number,
      CASE
	    WHEN vt.field_content IS NULL THEN b.best_title
        ELSE REGEXP_REPLACE(SPLIT_PART(REGEXP_REPLACE(vt.field_content,'^.*\|a',''),'|',1),'\s?(\.|\,|\:|\/|\;|\=)\s?$','')
      END AS title,
      b.best_title AS title_english,
      CASE
	    WHEN va.field_content IS NULL THEN REPLACE(SPLIT_PART(SPLIT_PART(b.best_author,' (',1),', ',2),'.','')||' '||SPLIT_PART(b.best_author,', ',1)
        ELSE REGEXP_REPLACE(REPLACE(REPLACE(REGEXP_REPLACE(SPLIT_PART(va.field_content,'|e',1),'^.*\|a',''),'|d',' '),'|q',' '),'\s?(\.|\,|\:|\/|\;|\=)\s?$','')
      END AS author,
      b.best_author AS author_english,
      'https://catalog.minlib.net/Record/.'||rmb.record_type_code||rmb.record_num||
        COALESCE(
          CAST(
            NULLIF(
              (
                ( rmb.record_num % 10 ) * 2 +
                ( rmb.record_num / 10 % 10 ) * 3 +
                ( rmb.record_num / 100 % 10 ) * 4 +
                ( rmb.record_num / 1000 % 10 ) * 5 +
                ( rmb.record_num / 10000 % 10 ) * 6 +
                ( rmb.record_num / 100000 % 10 ) * 7 +
                ( rmb.record_num / 1000000 ) * 8
               ) % 11,
             10
             )
          AS CHAR(1)
        ), 'x'
      ) AS url,
      (
        SELECT
          SUBSTRING(s.content FROM '[0-9]+')
        FROM sierra_view.subfield s
        WHERE b.bib_record_id = s.record_id
          AND s.marc_tag = '020'
	      AND s.tag = 'a'
        ORDER BY s.occ_num
        LIMIT 1
      ) AS ISBN,
      i.barcode,
      ir.checkout_total,
      ir.year_to_date_checkout_total,
      ir.last_year_to_date_checkout_total,
      m.creation_date_gmt::DATE AS created_date,
      ir.last_checkin_gmt::DATE AS last_checkin,
      o.due_gmt::DATE AS due_date,
      CASE
        WHEN o.id IS NOT NULL THEN 'CHECKED OUT'
        ELSE isp.name
      END AS status

    FROM sierra_view.item_record ir
    JOIN sierra_view.record_metadata m
      ON ir.id = m.id 
    JOIN sierra_view.bib_record_item_record_link bi 
      ON ir.id = bi.item_record_id
    JOIN sierra_view.record_metadata rmb
      ON bi.bib_record_id = rmb.id
    JOIN sierra_view.bib_record_property b
      ON bi.bib_record_id = b.bib_record_id
    JOIN sierra_view.item_record_property i
      ON ir.id = i.item_record_id
    JOIN sierra_view.item_status_property_myuser AS isp
      ON ir.item_status_code = isp.code
    LEFT JOIN sierra_view.varfield vt
      ON b.bib_record_id = vt.record_id
      AND vt.marc_tag = '880'
      AND vt.field_content ~ '^/|6245'
    LEFT JOIN sierra_view.varfield va
      ON b.bib_record_id = va.record_id
      AND va.marc_tag = '880'
      AND va.field_content ~ '^/|6100'
    LEFT JOIN sierra_view.checkout o
      ON ir.id = o.item_record_id
    /*
    JOIN sierra_view.bib_record br
      ON b.bib_record_id = br.id AND br.language_code = 'chi'
    */

    WHERE ir.location_code ~ '^bed'
      AND ir.icode1 IN ("""
        + scat_code
        + """)
      AND m.creation_date_gmt::DATE """
        + date_limit
        + """
    --GROUP BY b.bib_record_id,1,2,3,b.material_code 
    ORDER BY 1

"""
    )
    excel_file = (
        "/Scripts/Bed_Lex Chinese Book Lists/Temp Files/BEDChineseTitles"
        + file
        + ".xlsx"
    )
    query_results = run_query(query)
    excel_writer(query_results, excel_file)


main("New", "BETWEEN '2025-10-01' AND '2026-01-16'", "'107','109'")
main("JuvNew", "BETWEEN '2025-10-01' AND '2026-01-16'", "'108'")
main("107", "<= current_date", "'107'")
main("109", "<= current_date", "'109'")
main("Juv108", "<= current_date", "'108'")
