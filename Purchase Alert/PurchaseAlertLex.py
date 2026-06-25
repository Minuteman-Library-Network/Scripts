#!/usr/bin/env python3

#run in py313

"""
Create weekly purchase alert customized to Lexington's parameters
Upload file to staff site via sftp for distribution to staff
"""

import psycopg2
import xlsxwriter
import os
import pysftp
import configparser
import sys
import time
from datetime import date
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

#convert sql query results into formatted excel file
def excel_writer(query_results,excel_file):

    #Creating the Excel file for staff
    workbook = xlsxwriter.Workbook(excel_file,{'remove_timezone': True})
    worksheet_adf = workbook.add_worksheet('Adult Fic')
    worksheet_adnf = workbook.add_worksheet('Adult NF')
    worksheet_adlp = workbook.add_worksheet('Large Print')
    worksheet_adunknown = workbook.add_worksheet('Adult Unknown')
    worksheet_adav = workbook.add_worksheet('Adult AV')
    worksheet_j = workbook.add_worksheet('Juv')
    worksheet_ya = workbook.add_worksheet('YA')
    worksheet_other = workbook.add_worksheet('Other')


    #Formatting our Excel worksheet
    worksheet_adf.set_landscape()
    worksheet_adf.hide_gridlines(0)

    #Formatting Cells
    eformat= workbook.add_format({'text_wrap': True, 'valign': 'top'})
    eformatlabel= workbook.add_format({'text_wrap': True, 'valign': 'top', 'bold': True})
    link_format = workbook.add_format({'color': 'blue', 'underline': 1})

    # Setting the column widths
    worksheet_adf.set_column(0,0,14.14)
    worksheet_adf.set_column(1,1,42.43)
    worksheet_adf.set_column(2,2,26.57)
    worksheet_adf.set_column(3,3,10.14)
    worksheet_adf.set_column(4,4,13)
    worksheet_adf.set_column(5,5,9.14)
    worksheet_adf.set_column(6,6,12.71)
    worksheet_adf.set_column(7,7,8.71)
    worksheet_adf.set_column(8,8,12.43)
    worksheet_adf.set_column(9,9,12.71)
    worksheet_adf.set_column(10,10,12.71)
    worksheet_adf.set_column(11,11,9.57)
    worksheet_adf.set_column(12,12,22)
    worksheet_adf.set_column(13,13,8.57)
    worksheet_adf.set_column(14,14,12.43)
    worksheet_adf.set_column(15,15,17.14)
    worksheet_adf.set_column(16,16,13.57)
    worksheet_adf.set_column(17,17,45.71)
    worksheet_adnf.set_column(0,0,14.14)
    worksheet_adnf.set_column(1,1,42.43)
    worksheet_adnf.set_column(2,2,26.57)
    worksheet_adnf.set_column(3,3,10.14)
    worksheet_adnf.set_column(4,4,13)
    worksheet_adnf.set_column(5,5,9.14)
    worksheet_adnf.set_column(6,6,12.71)
    worksheet_adnf.set_column(7,7,8.71)
    worksheet_adnf.set_column(8,8,12.43)
    worksheet_adnf.set_column(9,9,12.71)
    worksheet_adnf.set_column(10,10,12.71)
    worksheet_adnf.set_column(11,11,9.57)
    worksheet_adnf.set_column(12,12,22)
    worksheet_adnf.set_column(13,13,8.57)
    worksheet_adnf.set_column(14,14,12.43)
    worksheet_adnf.set_column(15,15,17.14)
    worksheet_adnf.set_column(16,16,13.57)
    worksheet_adnf.set_column(17,17,45.71)
    worksheet_adlp.set_column(0,0,14.14)
    worksheet_adlp.set_column(1,1,42.43)
    worksheet_adlp.set_column(2,2,26.57)
    worksheet_adlp.set_column(3,3,10.14)
    worksheet_adlp.set_column(4,4,13)
    worksheet_adlp.set_column(5,5,9.14)
    worksheet_adlp.set_column(6,6,12.71)
    worksheet_adlp.set_column(7,7,8.71)
    worksheet_adlp.set_column(8,8,12.43)
    worksheet_adlp.set_column(9,9,12.71)
    worksheet_adlp.set_column(10,10,12.71)
    worksheet_adlp.set_column(11,11,9.57)
    worksheet_adlp.set_column(12,12,22)
    worksheet_adlp.set_column(13,13,8.57)
    worksheet_adlp.set_column(14,14,12.43)
    worksheet_adlp.set_column(15,15,17.14)
    worksheet_adlp.set_column(16,16,13.57)
    worksheet_adlp.set_column(17,17,45.71)
    worksheet_adunknown.set_column(0,0,14.14)
    worksheet_adunknown.set_column(1,1,42.43)
    worksheet_adunknown.set_column(2,2,26.57)
    worksheet_adunknown.set_column(3,3,10.14)
    worksheet_adunknown.set_column(4,4,13)
    worksheet_adunknown.set_column(5,5,9.14)
    worksheet_adunknown.set_column(6,6,12.71)
    worksheet_adunknown.set_column(7,7,8.71)
    worksheet_adunknown.set_column(8,8,12.43)
    worksheet_adunknown.set_column(9,9,12.71)
    worksheet_adunknown.set_column(10,10,12.71)
    worksheet_adunknown.set_column(11,11,9.57)
    worksheet_adunknown.set_column(12,12,22)
    worksheet_adunknown.set_column(13,13,8.57)
    worksheet_adunknown.set_column(14,14,12.43)
    worksheet_adunknown.set_column(15,15,17.14)
    worksheet_adunknown.set_column(16,16,13.57)
    worksheet_adunknown.set_column(17,17,45.71)
    worksheet_adav.set_column(0,0,14.14)
    worksheet_adav.set_column(1,1,42.43)
    worksheet_adav.set_column(2,2,26.57)
    worksheet_adav.set_column(3,3,10.14)
    worksheet_adav.set_column(4,4,13)
    worksheet_adav.set_column(5,5,9.14)
    worksheet_adav.set_column(6,6,12.71)
    worksheet_adav.set_column(7,7,8.71)
    worksheet_adav.set_column(8,8,12.43)
    worksheet_adav.set_column(9,9,12.71)
    worksheet_adav.set_column(10,10,12.71)
    worksheet_adav.set_column(11,11,9.57)
    worksheet_adav.set_column(12,12,22)
    worksheet_adav.set_column(13,13,8.57)
    worksheet_adav.set_column(14,14,12.43)
    worksheet_adav.set_column(15,15,17.14)
    worksheet_adav.set_column(16,16,13.57)
    worksheet_adav.set_column(17,17,45.71)
    worksheet_j.set_column(0,0,14.14)
    worksheet_j.set_column(1,1,42.43)
    worksheet_j.set_column(2,2,26.57)
    worksheet_j.set_column(3,3,10.14)
    worksheet_j.set_column(4,4,13)
    worksheet_j.set_column(5,5,9.14)
    worksheet_j.set_column(6,6,12.71)
    worksheet_j.set_column(7,7,8.71)
    worksheet_j.set_column(8,8,12.43)
    worksheet_j.set_column(9,9,12.71)
    worksheet_j.set_column(10,10,12.71)
    worksheet_j.set_column(11,11,9.57)
    worksheet_j.set_column(12,12,22)
    worksheet_j.set_column(13,13,8.57)
    worksheet_j.set_column(14,14,12.43)
    worksheet_j.set_column(15,15,17.14)
    worksheet_j.set_column(16,16,13.57)
    worksheet_j.set_column(17,17,45.71)
    worksheet_ya.set_column(0,0,14.14)
    worksheet_ya.set_column(1,1,42.43)
    worksheet_ya.set_column(2,2,26.57)
    worksheet_ya.set_column(3,3,10.14)
    worksheet_ya.set_column(4,4,13)
    worksheet_ya.set_column(5,5,9.14)
    worksheet_ya.set_column(6,6,12.71)
    worksheet_ya.set_column(7,7,8.71)
    worksheet_ya.set_column(8,8,12.43)
    worksheet_ya.set_column(9,9,12.71)
    worksheet_ya.set_column(10,10,12.71)
    worksheet_ya.set_column(11,11,9.57)
    worksheet_ya.set_column(12,12,22)
    worksheet_ya.set_column(13,13,8.57)
    worksheet_ya.set_column(14,14,12.43)
    worksheet_ya.set_column(15,15,17.14)
    worksheet_ya.set_column(16,16,13.57)
    worksheet_ya.set_column(17,17,45.71)
    worksheet_other.set_column(0,0,14.14)
    worksheet_other.set_column(1,1,42.43)
    worksheet_other.set_column(2,2,26.57)
    worksheet_other.set_column(3,3,10.14)
    worksheet_other.set_column(4,4,13)
    worksheet_other.set_column(5,5,9.14)
    worksheet_other.set_column(6,6,12.71)
    worksheet_other.set_column(7,7,8.71)
    worksheet_other.set_column(8,8,12.43)
    worksheet_other.set_column(9,9,12.71)
    worksheet_other.set_column(10,10,12.71)
    worksheet_other.set_column(11,11,9.57)
    worksheet_other.set_column(12,12,22)
    worksheet_other.set_column(13,13,8.57)
    worksheet_other.set_column(14,14,12.43)
    worksheet_other.set_column(15,15,17.14)
    worksheet_other.set_column(16,16,13.57)
    worksheet_other.set_column(17,17,45.71)

    #Inserting a header
    worksheet_adf.set_header('Purchase Alert')

    # Adding column labels
    worksheet_adf.write(0,0,'Record_number', eformatlabel)
    worksheet_adf.write(0,1,'Title', eformatlabel)
    worksheet_adf.write(0,2,'Author', eformatlabel)
    worksheet_adf.write(0,3,'PublicationYear', eformatlabel)
    worksheet_adf.write(0,4,'MatType', eformatlabel)
    worksheet_adf.write(0,5,'TotalItemCount', eformatlabel)
    worksheet_adf.write(0,6,'TotalAvailableItemCount', eformatlabel)
    worksheet_adf.write(0,7,'TotalHoldCount', eformatlabel)
    worksheet_adf.write(0,8,'TotalDemandRatio', eformatlabel)
    worksheet_adf.write(0,9,'LocalHoldableItemCount', eformatlabel)
    worksheet_adf.write(0,10,'LocalSpeedItemCount', eformatlabel)
    worksheet_adf.write(0,11,'LocalOrderCopies', eformatlabel)
    worksheet_adf.write(0,12,'LocalCopiesInProcess', eformatlabel)
    worksheet_adf.write(0,13,'LocalHoldCount', eformatlabel)
    worksheet_adf.write(0,14,'LocalDemandRatio', eformatlabel)
    worksheet_adf.write(0,15,'SuggestedPurchaseQty (3)', eformatlabel)
    worksheet_adf.write(0,16,'OrderLocations', eformatlabel)
    worksheet_adf.write(0,17,'IsbnUPC', eformatlabel)
    worksheet_adnf.write(0,0,'Record_number', eformatlabel)
    worksheet_adnf.write(0,1,'Title', eformatlabel)
    worksheet_adnf.write(0,2,'Author', eformatlabel)
    worksheet_adnf.write(0,3,'PublicationYear', eformatlabel)
    worksheet_adnf.write(0,4,'MatType', eformatlabel)
    worksheet_adnf.write(0,5,'TotalItemCount', eformatlabel)
    worksheet_adnf.write(0,6,'TotalAvailableItemCount', eformatlabel)
    worksheet_adnf.write(0,7,'TotalHoldCount', eformatlabel)
    worksheet_adnf.write(0,8,'TotalDemandRatio', eformatlabel)
    worksheet_adnf.write(0,9,'LocalHoldableItemCount', eformatlabel)
    worksheet_adnf.write(0,10,'LocalSpeedItemCount', eformatlabel)
    worksheet_adnf.write(0,11,'LocalOrderCopies', eformatlabel)
    worksheet_adnf.write(0,12,'LocalCopiesInProcess', eformatlabel)
    worksheet_adnf.write(0,13,'LocalHoldCount', eformatlabel)
    worksheet_adnf.write(0,14,'LocalDemandRatio', eformatlabel)
    worksheet_adnf.write(0,15,'SuggestedPurchaseQty (3)', eformatlabel)
    worksheet_adnf.write(0,16,'OrderLocations', eformatlabel)
    worksheet_adnf.write(0,17,'IsbnUPC', eformatlabel)
    worksheet_adlp.write(0,0,'Record_number', eformatlabel)
    worksheet_adlp.write(0,1,'Title', eformatlabel)
    worksheet_adlp.write(0,2,'Author', eformatlabel)
    worksheet_adlp.write(0,3,'PublicationYear', eformatlabel)
    worksheet_adlp.write(0,4,'MatType', eformatlabel)
    worksheet_adlp.write(0,5,'TotalItemCount', eformatlabel)
    worksheet_adlp.write(0,6,'TotalAvailableItemCount', eformatlabel)
    worksheet_adlp.write(0,7,'TotalHoldCount', eformatlabel)
    worksheet_adlp.write(0,8,'TotalDemandRatio', eformatlabel)
    worksheet_adlp.write(0,9,'LocalHoldableItemCount', eformatlabel)
    worksheet_adlp.write(0,10,'LocalSpeedItemCount', eformatlabel)
    worksheet_adlp.write(0,11,'LocalOrderCopies', eformatlabel)
    worksheet_adlp.write(0,12,'LocalCopiesInProcess', eformatlabel)
    worksheet_adlp.write(0,13,'LocalHoldCount', eformatlabel)
    worksheet_adlp.write(0,14,'LocalDemandRatio', eformatlabel)
    worksheet_adlp.write(0,15,'SuggestedPurchaseQty (3)', eformatlabel)
    worksheet_adlp.write(0,16,'OrderLocations', eformatlabel)
    worksheet_adlp.write(0,17,'IsbnUPC', eformatlabel)
    worksheet_adunknown.write(0,0,'Record_number', eformatlabel)
    worksheet_adunknown.write(0,1,'Title', eformatlabel)
    worksheet_adunknown.write(0,2,'Author', eformatlabel)
    worksheet_adunknown.write(0,3,'PublicationYear', eformatlabel)
    worksheet_adunknown.write(0,4,'MatType', eformatlabel)
    worksheet_adunknown.write(0,5,'TotalItemCount', eformatlabel)
    worksheet_adunknown.write(0,6,'TotalAvailableItemCount', eformatlabel)
    worksheet_adunknown.write(0,7,'TotalHoldCount', eformatlabel)
    worksheet_adunknown.write(0,8,'TotalDemandRatio', eformatlabel)
    worksheet_adunknown.write(0,9,'LocalHoldableItemCount', eformatlabel)
    worksheet_adunknown.write(0,10,'LocalSpeedItemCount', eformatlabel)
    worksheet_adunknown.write(0,11,'LocalOrderCopies', eformatlabel)
    worksheet_adunknown.write(0,12,'LocalCopiesInProcess', eformatlabel)
    worksheet_adunknown.write(0,13,'LocalHoldCount', eformatlabel)
    worksheet_adunknown.write(0,14,'LocalDemandRatio', eformatlabel)
    worksheet_adunknown.write(0,15,'SuggestedPurchaseQty (3)', eformatlabel)
    worksheet_adunknown.write(0,16,'OrderLocations', eformatlabel)
    worksheet_adunknown.write(0,17,'IsbnUPC', eformatlabel)
    worksheet_adav.write(0,0,'Record_number', eformatlabel)
    worksheet_adav.write(0,1,'Title', eformatlabel)
    worksheet_adav.write(0,2,'Author', eformatlabel)
    worksheet_adav.write(0,3,'PublicationYear', eformatlabel)
    worksheet_adav.write(0,4,'MatType', eformatlabel)
    worksheet_adav.write(0,5,'TotalItemCount', eformatlabel)
    worksheet_adav.write(0,6,'TotalAvailableItemCount', eformatlabel)
    worksheet_adav.write(0,7,'TotalHoldCount', eformatlabel)
    worksheet_adav.write(0,8,'TotalDemandRatio', eformatlabel)
    worksheet_adav.write(0,9,'LocalHoldableItemCount', eformatlabel)
    worksheet_adav.write(0,10,'LocalSpeedItemCount', eformatlabel)
    worksheet_adav.write(0,11,'LocalOrderCopies', eformatlabel)
    worksheet_adav.write(0,12,'LocalCopiesInProcess', eformatlabel)
    worksheet_adav.write(0,13,'LocalHoldCount', eformatlabel)
    worksheet_adav.write(0,14,'LocalDemandRatio', eformatlabel)
    worksheet_adav.write(0,15,'SuggestedPurchaseQty (3)', eformatlabel)
    worksheet_adav.write(0,16,'OrderLocations', eformatlabel)
    worksheet_adav.write(0,17,'IsbnUPC', eformatlabel)
    worksheet_j.write(0,0,'Record_number', eformatlabel)
    worksheet_j.write(0,1,'Title', eformatlabel)
    worksheet_j.write(0,2,'Author', eformatlabel)
    worksheet_j.write(0,3,'PublicationYear', eformatlabel)
    worksheet_j.write(0,4,'MatType', eformatlabel)
    worksheet_j.write(0,5,'TotalItemCount', eformatlabel)
    worksheet_j.write(0,6,'TotalAvailableItemCount', eformatlabel)
    worksheet_j.write(0,7,'TotalHoldCount', eformatlabel)
    worksheet_j.write(0,8,'TotalDemandRatio', eformatlabel)
    worksheet_j.write(0,9,'LocalHoldableItemCount', eformatlabel)
    worksheet_j.write(0,10,'LocalSpeedItemCount', eformatlabel)
    worksheet_j.write(0,11,'LocalOrderCopies', eformatlabel)
    worksheet_j.write(0,12,'LocalCopiesInProcess', eformatlabel)
    worksheet_j.write(0,13,'LocalHoldCount', eformatlabel)
    worksheet_j.write(0,14,'LocalDemandRatio', eformatlabel)
    worksheet_j.write(0,15,'SuggestedPurchaseQty (3)', eformatlabel)
    worksheet_j.write(0,16,'OrderLocations', eformatlabel)
    worksheet_j.write(0,17,'IsbnUPC', eformatlabel)
    worksheet_ya.write(0,0,'Record_number', eformatlabel)
    worksheet_ya.write(0,1,'Title', eformatlabel)
    worksheet_ya.write(0,2,'Author', eformatlabel)
    worksheet_ya.write(0,3,'PublicationYear', eformatlabel)
    worksheet_ya.write(0,4,'MatType', eformatlabel)
    worksheet_ya.write(0,5,'TotalItemCount', eformatlabel)
    worksheet_ya.write(0,6,'TotalAvailableItemCount', eformatlabel)
    worksheet_ya.write(0,7,'TotalHoldCount', eformatlabel)
    worksheet_ya.write(0,8,'TotalDemandRatio', eformatlabel)
    worksheet_ya.write(0,9,'LocalHoldableItemCount', eformatlabel)
    worksheet_ya.write(0,10,'LocalSpeedItemCount', eformatlabel)
    worksheet_ya.write(0,11,'LocalOrderCopies', eformatlabel)
    worksheet_ya.write(0,12,'LocalCopiesInProcess', eformatlabel)
    worksheet_ya.write(0,13,'LocalHoldCount', eformatlabel)
    worksheet_ya.write(0,14,'LocalDemandRatio', eformatlabel)
    worksheet_ya.write(0,15,'SuggestedPurchaseQty (3)', eformatlabel)
    worksheet_ya.write(0,16,'OrderLocations', eformatlabel)
    worksheet_ya.write(0,17,'IsbnUPC', eformatlabel)
    worksheet_other.write(0,0,'Record_number', eformatlabel)
    worksheet_other.write(0,1,'Title', eformatlabel)
    worksheet_other.write(0,2,'Author', eformatlabel)
    worksheet_other.write(0,3,'PublicationYear', eformatlabel)
    worksheet_other.write(0,4,'MatType', eformatlabel)
    worksheet_other.write(0,5,'TotalItemCount', eformatlabel)
    worksheet_other.write(0,6,'TotalAvailableItemCount', eformatlabel)
    worksheet_other.write(0,7,'TotalHoldCount', eformatlabel)
    worksheet_other.write(0,8,'TotalDemandRatio', eformatlabel)
    worksheet_other.write(0,9,'LocalHoldableItemCount', eformatlabel)
    worksheet_other.write(0,10,'LocalSpeedItemCount', eformatlabel)
    worksheet_other.write(0,11,'LocalOrderCopies', eformatlabel)
    worksheet_other.write(0,12,'LocalCopiesInProcess', eformatlabel)
    worksheet_other.write(0,13,'LocalHoldCount', eformatlabel)
    worksheet_other.write(0,14,'LocalDemandRatio', eformatlabel)
    worksheet_other.write(0,15,'SuggestedPurchaseQty (3)', eformatlabel)
    worksheet_other.write(0,16,'OrderLocations', eformatlabel)
    worksheet_other.write(0,17,'IsbnUPC', eformatlabel)

    row_adf = 1
    row_adnf = 1
    row_adlp = 1
    row_adunknown = 1
    row_adav = 1
    row_j = 1
    row_ya = 1
    row_other = 1
    # suggested_purchase_formula = '=IF(N2/(VALUE(MID($P$1,SEARCH("(",$P$1)+1,SEARCH(")",$P$1)-SEARCH("(",$P$1)-1)+0))-J2-L2-M2<0,0,ROUND(N2/(VALUE(MID($P$1,SEARCH("(",$P$1)+1,SEARCH(")",$P$1)-SEARCH("(",$P$1)-1)+0))-J2-L2-M2,1))'
    suggested_purchase_formula = '=if(N{}/(VALUE(MID($P$1,SEARCH("(",$P$1)+1,SEARCH(")",$P$1)-SEARCH("(",$P$1)-1)+0))-J{}-L{}-M{}<0,0,round(N{}/(VALUE(MID($P$1,SEARCH("(",$P$1)+1,SEARCH(")",$P$1)-SEARCH("(",$P$1)-1)+0))-J{}-L{}-M{},1))'

    # Writing the report for staff to the Excel worksheet
    for rownum, row in enumerate(query_results):
        if row[4] == 'LARGE PRINT':
            worksheet_adlp.write(row_adlp,0,row[0], eformat)
            worksheet_adlp.write_url(row_adlp,1,row[13], link_format, row[1])
            worksheet_adlp.write(row_adlp,2,row[2], eformat)
            worksheet_adlp.write(row_adlp,3,row[3], eformat)
            worksheet_adlp.write(row_adlp,4,row[4], eformat)
            worksheet_adlp.write(row_adlp,5,row[5], eformat)
            worksheet_adlp.write(row_adlp,6,row[6], eformat)
            worksheet_adlp.write(row_adlp,7,row[7], eformat)
            worksheet_adlp.write(row_adlp,8,row[8], eformat)
            worksheet_adlp.write(row_adlp,9,row[9], eformat)
            worksheet_adlp.write(row_adlp,10,row[19], eformat)
            worksheet_adlp.write(row_adlp,11,row[10], eformat)
            worksheet_adlp.write(row_adlp,12,row[18], eformat)
            worksheet_adlp.write(row_adlp,13,row[11], eformat)
            worksheet_adlp.write(row_adlp,14,row[12], eformat)
            worksheet_adlp.write(row_adlp,15,suggested_purchase_formula.format(row_adlp+1,row_adlp+1,row_adlp+1,row_adlp+1,row_adlp+1,row_adlp+1,row_adlp+1,row_adlp+1), eformat)
            worksheet_adlp.write(row_adlp,16,row[14], eformat)
            worksheet_adlp.write(row_adlp,17,row[15], eformat)
            row_adlp += 1
        elif row[16] == 'ADULT' and row[17] == 'TRUE' and row[4] == 'BOOK':
            worksheet_adf.write(row_adf,0,row[0], eformat)
            worksheet_adf.write_url(row_adf,1,row[13], link_format, row[1])
            worksheet_adf.write(row_adf,2,row[2], eformat)
            worksheet_adf.write(row_adf,3,row[3], eformat)
            worksheet_adf.write(row_adf,4,row[4], eformat)
            worksheet_adf.write(row_adf,5,row[5], eformat)
            worksheet_adf.write(row_adf,6,row[6], eformat)
            worksheet_adf.write(row_adf,7,row[7], eformat)
            worksheet_adf.write(row_adf,8,row[8], eformat)
            worksheet_adf.write(row_adf,9,row[9], eformat)
            worksheet_adf.write(row_adf,10,row[19], eformat)
            worksheet_adf.write(row_adf,11,row[10], eformat)
            worksheet_adf.write(row_adf,12,row[18], eformat)
            worksheet_adf.write(row_adf,13,row[11], eformat)
            worksheet_adf.write(row_adf,14,row[12], eformat)
            worksheet_adf.write(row_adf,15,suggested_purchase_formula.format(row_adf+1,row_adf+1,row_adf+1,row_adf+1,row_adf+1,row_adf+1,row_adf+1,row_adf+1), eformat)
            worksheet_adf.write(row_adf,16,row[14], eformat)
            worksheet_adf.write(row_adf,17,row[15], eformat)
            row_adf += 1
        elif row[16] == 'ADULT' and row[17] == 'FALSE' and row[4] in ['BOOK','MUSIC SCORE']:
            worksheet_adnf.write(row_adnf,0,row[0], eformat)
            worksheet_adnf.write_url(row_adnf,1,row[13], link_format, row[1])
            worksheet_adnf.write(row_adnf,2,row[2], eformat)
            worksheet_adnf.write(row_adnf,3,row[3], eformat)
            worksheet_adnf.write(row_adnf,4,row[4], eformat)
            worksheet_adnf.write(row_adnf,5,row[5], eformat)
            worksheet_adnf.write(row_adnf,6,row[6], eformat)
            worksheet_adnf.write(row_adnf,7,row[7], eformat)
            worksheet_adnf.write(row_adnf,8,row[8], eformat)
            worksheet_adnf.write(row_adnf,9,row[9], eformat)
            worksheet_adnf.write(row_adnf,10,row[19], eformat)
            worksheet_adnf.write(row_adnf,11,row[10], eformat)
            worksheet_adnf.write(row_adnf,12,row[18], eformat)            
            worksheet_adnf.write(row_adnf,13,row[11], eformat)
            worksheet_adnf.write(row_adnf,14,row[12], eformat)
            worksheet_adnf.write(row_adnf,15,suggested_purchase_formula.format(row_adnf+1,row_adnf+1,row_adnf+1,row_adnf+1,row_adnf+1,row_adnf+1,row_adnf+1,row_adnf+1), eformat)
            worksheet_adnf.write(row_adnf,16,row[14], eformat)
            worksheet_adnf.write(row_adnf,17,row[15], eformat)
            row_adnf += 1
        elif row[16] == 'ADULT' and row[17] == 'UNKNOWN' and row[4] in ['BOOK','MUSIC SCORE']:
            worksheet_adunknown.write(row_adunknown,0,row[0], eformat)
            worksheet_adunknown.write_url(row_adunknown,1,row[13], link_format, row[1])
            worksheet_adunknown.write(row_adunknown,2,row[2], eformat)
            worksheet_adunknown.write(row_adunknown,3,row[3], eformat)
            worksheet_adunknown.write(row_adunknown,4,row[4], eformat)
            worksheet_adunknown.write(row_adunknown,5,row[5], eformat)
            worksheet_adunknown.write(row_adunknown,6,row[6], eformat)
            worksheet_adunknown.write(row_adunknown,7,row[7], eformat)
            worksheet_adunknown.write(row_adunknown,8,row[8], eformat)
            worksheet_adunknown.write(row_adunknown,9,row[9], eformat)
            worksheet_adunknown.write(row_adunknown,10,row[19], eformat)
            worksheet_adunknown.write(row_adunknown,11,row[10], eformat)
            worksheet_adunknown.write(row_adunknown,12,row[18], eformat)
            worksheet_adunknown.write(row_adunknown,13,row[11], eformat)
            worksheet_adunknown.write(row_adunknown,14,row[12], eformat)
            worksheet_adunknown.write(row_adunknown,15,suggested_purchase_formula.format(row_adunknown+1,row_adunknown+1,row_adunknown+1,row_adunknown+1,row_adunknown+1,row_adunknown+1,row_adunknown+1,row_adunknown+1), eformat)
            worksheet_adunknown.write(row_adunknown,16,row[14], eformat)
            worksheet_adunknown.write(row_adunknown,17,row[15], eformat)
            row_adunknown += 1
        elif row[16] == 'JUV':
            worksheet_j.write(row_j,0,row[0], eformat)
            worksheet_j.write_url(row_j,1,row[13], link_format, row[1])
            worksheet_j.write(row_j,2,row[2], eformat)
            worksheet_j.write(row_j,3,row[3], eformat)
            worksheet_j.write(row_j,4,row[4], eformat)
            worksheet_j.write(row_j,5,row[5], eformat)
            worksheet_j.write(row_j,6,row[6], eformat)
            worksheet_j.write(row_j,7,row[7], eformat)
            worksheet_j.write(row_j,8,row[8], eformat)
            worksheet_j.write(row_j,9,row[9], eformat)
            worksheet_j.write(row_j,10,row[19], eformat)
            worksheet_j.write(row_j,11,row[10], eformat)
            worksheet_j.write(row_j,12,row[18], eformat)
            worksheet_j.write(row_j,13,row[11], eformat)
            worksheet_j.write(row_j,14,row[12], eformat)
            worksheet_j.write(row_j,15,suggested_purchase_formula.format(row_j+1,row_j+1,row_j+1,row_j+1,row_j+1,row_j+1,row_j+1,row_j+1), eformat)
            worksheet_j.write(row_j,16,row[14], eformat)
            worksheet_j.write(row_j,17,row[15], eformat)
            row_j += 1
        elif row[16] == 'YA':
            worksheet_ya.write(row_ya,0,row[0], eformat)
            worksheet_ya.write_url(row_ya,1,row[13], link_format, row[1])
            worksheet_ya.write(row_ya,2,row[2], eformat)
            worksheet_ya.write(row_ya,3,row[3], eformat)
            worksheet_ya.write(row_ya,4,row[4], eformat)
            worksheet_ya.write(row_ya,5,row[5], eformat)
            worksheet_ya.write(row_ya,6,row[6], eformat)
            worksheet_ya.write(row_ya,7,row[7], eformat)
            worksheet_ya.write(row_ya,8,row[8], eformat)
            worksheet_ya.write(row_ya,9,row[9], eformat)
            worksheet_ya.write(row_ya,10,row[19], eformat)
            worksheet_ya.write(row_ya,11,row[10], eformat)
            worksheet_ya.write(row_ya,12,row[18], eformat)
            worksheet_ya.write(row_ya,13,row[11], eformat)
            worksheet_ya.write(row_ya,14,row[12], eformat)
            worksheet_ya.write(row_ya,15,suggested_purchase_formula.format(row_ya+1,row_ya+1,row_ya+1,row_ya+1,row_ya+1,row_ya+1,row_ya+1,row_ya+1), eformat)
            worksheet_ya.write(row_ya,16,row[14], eformat)
            worksheet_ya.write(row_ya,17,row[15], eformat)
            row_ya += 1
        elif row[4] == '3-D OBJECT' or row[4] == 'BLU-RAY' or row[4] == 'CONSOLE GAME' or row[4] == 'DVD OR VCD' or row[4] == 'MUSIC CD' or row[4] == 'PLAYAWAY AUDIOBOOK' or row[4] == 'SPOKEN CD':
            worksheet_adav.write(row_adav,0,row[0], eformat)
            worksheet_adav.write_url(row_adav,1,row[13], link_format, row[1])
            worksheet_adav.write(row_adav,2,row[2], eformat)
            worksheet_adav.write(row_adav,3,row[3], eformat)
            worksheet_adav.write(row_adav,4,row[4], eformat)
            worksheet_adav.write(row_adav,5,row[5], eformat)
            worksheet_adav.write(row_adav,6,row[6], eformat)
            worksheet_adav.write(row_adav,7,row[7], eformat)
            worksheet_adav.write(row_adav,8,row[8], eformat)
            worksheet_adav.write(row_adav,9,row[9], eformat)
            worksheet_adav.write(row_adav,10,row[19], eformat)
            worksheet_adav.write(row_adav,11,row[10], eformat)
            worksheet_adav.write(row_adav,12,row[18], eformat)
            worksheet_adav.write(row_adav,13,row[11], eformat)
            worksheet_adav.write(row_adav,14,row[12], eformat)
            worksheet_adav.write(row_adav,15,suggested_purchase_formula.format(row_adav+1,row_adav+1,row_adav+1,row_adav+1,row_adav+1,row_adav+1,row_adav+1,row_adav+1), eformat)
            worksheet_adav.write(row_adav,16,row[14], eformat)
            worksheet_adav.write(row_adav,17,row[15], eformat)
            row_adav += 1
        else:
            worksheet_other.write(row_other,0,row[0], eformat)
            worksheet_other.write_url(row_other,1,row[13], link_format, row[1])
            worksheet_other.write(row_other,2,row[2], eformat)
            worksheet_other.write(row_other,3,row[3], eformat)
            worksheet_other.write(row_other,4,row[4], eformat)
            worksheet_other.write(row_other,5,row[5], eformat)
            worksheet_other.write(row_other,6,row[6], eformat)
            worksheet_other.write(row_other,7,row[7], eformat)
            worksheet_other.write(row_other,8,row[8], eformat)
            worksheet_other.write(row_other,9,row[9], eformat)
            worksheet_other.write(row_other,10,row[19], eformat)
            worksheet_other.write(row_other,11,row[10], eformat)
            worksheet_other.write(row_other,12,row[18], eformat)
            worksheet_other.write(row_other,13,row[11], eformat)
            worksheet_other.write(row_other,14,row[12], eformat)
            worksheet_other.write(row_other,15,suggested_purchase_formula.format(row_other+1,row_other+1,row_other+1,row_other+1,row_other+1,row_other+1,row_other+1,row_other+1), eformat)
            worksheet_other.write(row_other,16,row[14], eformat)
            worksheet_other.write(row_other,17,row[15], eformat)
            row_other += 1
    
    workbook.close()
    
    return excel_file

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
        + "/Purchase Alert/"
    )
    srv.put(local_file)

    for fname in srv.listdir_attr():
        fullpath = (
            "/reports/Library-Specific Reports/"
            + library
            + "/Purchase Alert/{}".format(fname.filename)
        )
        # time tracked in seconds, st_mtime is time last modified
        name = str(fname.filename)
        if (name != "meta.json") and (
            (time.time() - fname.st_mtime) // (24 * 3600) >= 90
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

def main():
	
    query = r"""
    WITH orders AS (
  SELECT
    COUNT(oc.order_record_id) FILTER(WHERE o.order_status_code = 'o') AS order_count,
    SUM(oc.copies) FILTER (WHERE o.order_status_code = 'o') AS order_copies,
    SUM(oc.copies) FILTER(WHERE o.order_status_code = 'a' AND o.received_date_gmt::DATE >= CURRENT_DATE - INTERVAL '14 days') AS processing_copies,
    bro.bib_record_id AS bib_id,
    STRING_AGG(DISTINCT(oc.location_code), ',') AS order_locations
        
	FROM sierra_view.order_record o
	JOIN sierra_view.order_record_cmf oc
	  ON o.id = oc.order_record_id
	JOIN sierra_view.bib_record_order_record_link bro
	  ON o.id=bro.order_record_id
        
	WHERE o.order_status_code IN ('o','a')
	  AND oc.location_code ~ '^lex'	
	  --location will take the form ^oln, which in this example looks for all locations starting with the string oln.
   GROUP BY bro.bib_record_id
),

hold_data AS(
SELECT 
	b.id AS bib_id, 
	COUNT(DISTINCT h.id) AS hold_count,
	orders.order_locations AS order_locations,
	COUNT(DISTINCT h.id) FILTER(WHERE h.pickup_location_code ~ '^lex') AS local_holds,
	COUNT(DISTINCT i.id) AS item_count,
	COUNT(DISTINCT ia.id) AS avail_item_count, 
	COUNT(DISTINCT ia.id) FILTER(WHERE ia.location_code ~ '^lex' AND rmia.creation_date_gmt::DATE >= CURRENT_DATE - INTERVAL '14 days') AS in_process_item_count,
	COUNT(DISTINCT ia.id) FILTER(WHERE ia.location_code ~ '^lex' AND ia.itype_code_num NOT IN ('5','21','109','133','160')) AS local_holdable_item_count,
	COUNT(DISTINCT ia.id) FILTER(WHERE ia.location_code ~ '^lex' AND ia.itype_code_num IN ('5','21','109','133','160')) AS local_speed_item_count,
	COUNT(DISTINCT ia.id) FILTER(WHERE ia.location_code ~ '^lex') AS local_avail_item_count,
	MAX(orders.order_count) AS order_count,
	CASE
      WHEN MAX(orders.order_copies) IS NULL THEN 0
      ELSE MAX(orders.order_copies)
   END AS order_copies,
   CASE
    	WHEN MAX(orders.processing_copies) IS NULL THEN 0
    	ELSE MAX(orders.processing_copies)
   END AS processing_copies,
	MODE() WITHIN GROUP (ORDER BY SUBSTRING(i.location_code,4,1)) AS age_level

  FROM sierra_view.bib_record b
  LEFT JOIN sierra_view.bib_record_item_record_link bri
    ON b.id = bri.bib_record_id
  JOIN sierra_view.hold h
    ON b.id = h.record_id
	 OR bri.item_record_id = h.record_id
  LEFT JOIN sierra_view.item_record i
    ON bri.item_record_id = i.id
  --secondary join to only available items
  LEFT JOIN sierra_view.item_record ia
    ON ia.id=bri.item_record_id
	 AND ia.item_status_code IN ('-','t','p','!')
	 AND (
	   (ia.location_code !~ '^lex'
	     AND ia.itype_code_num NOT IN ('5','21','109','133','160','183','239','240','241','244','248','249')) 
	   OR ia.location_code ~ '^lex'
		--location will take the form ^oln, which in this example looks for all locations starting with the string oln.
	   )
  LEFT JOIN sierra_view.record_metadata rmia
    ON ia.id = rmia.id
  LEFT JOIN orders 
    ON orders.bib_id=b.id

  WHERE h.status='0'
		
  GROUP BY b.id,3
  HAVING COUNT(DISTINCT h.id)>0
)

SELECT * FROM (
  SELECT
	id2reckey(hd.bib_id)||'a' AS RecordNumber,
	brp.best_title AS Title, 
	brp.best_author AS Author, 
   --publish_year logic from Rebecca King, Thousand Oaks Library
   CASE
     -- valid 4-digit year
	  WHEN brp.publish_year BETWEEN 1000 AND 2099 THEN brp.publish_year
     -- possible corrupted YYYYMMDD → try first 2 digits as year
	  WHEN brp.publish_year BETWEEN 10000000 AND 99999999 THEN LEFT(brp.publish_year::TEXT, 4)::INTEGER
     -- possible truncated YYMM like 2603 → interpret as 2026
	  WHEN brp.publish_year BETWEEN 0 AND 9999 AND LENGTH(brp.publish_year::TEXT) = 4
	    THEN 2000 + LEFT(brp.publish_year::TEXT, 2)::INTEGER
	  ELSE NULL
   END AS publication_year,
	mp.name AS mattype,
	hd.item_count AS TotalItemCount, 
	MAX(hd.avail_item_count) AS AvailableItemCount,
	hd.hold_count AS TotalHoldCount,
	CASE
  	  WHEN MAX(hd.avail_item_count) + MAX(hd.order_copies) + (
	    CASE
		   WHEN MAX(hd.processing_copies) > 0 AND MAX(hd.in_process_item_count) < MAX(hd.processing_copies) THEN MAX(hd.processing_copies) - MAX(hd.in_process_item_count)
			ELSE 0
		 END)
	  = 0 THEN hd.hold_count
  	  ELSE ROUND(CAST((hd.hold_count) AS NUMERIC(12, 2))/CAST((MAX(hd.avail_item_count) + MAX(hd.order_copies) + (
		 CASE
			WHEN MAX(hd.processing_copies) > 0 AND MAX(hd.in_process_item_count) < MAX(hd.processing_copies) THEN MAX(hd.processing_copies) - MAX(hd.in_process_item_count)
			ELSE 0
		 END)
	  ) AS NUMERIC(12,2)),2)
  	END AS TotalRatio,
	MAX(hd.local_holdable_item_count) AS LocalHoldableItemCount,
	MAX(hd.order_copies) AS LocalOrderCopies,
	hd.local_holds AS LocalHoldCount,
   CASE
     WHEN MAX(hd.local_avail_item_count) + MAX(hd.order_copies) + (
	    CASE
			WHEN MAX(hd.processing_copies) > 0 AND MAX(hd.in_process_item_count) < MAX(hd.processing_copies) THEN MAX(hd.processing_copies) - MAX(hd.in_process_item_count)
			ELSE 0
	    END)
	  = 0 THEN hd.local_holds
     ELSE ROUND(CAST((hd.local_holds) AS NUMERIC(12, 2))/CAST((MAX(hd.local_avail_item_count) + MAX(hd.order_copies) + (
		 CASE
			WHEN MAX(hd.processing_copies) > 0 AND MAX(hd.in_process_item_count) < MAX(hd.processing_copies) THEN MAX(hd.processing_copies) - MAX(hd.in_process_item_count)
			ELSE 0
		 END)
	  ) AS NUMERIC(12,2)),2)
   END AS LocalRatio,
	'https://catalog.minlib.net/Record/'||id2reckey(hd.bib_id) AS URL,
	hd.order_locations AS OrderLocations,
	(SELECT
		COALESCE(STRING_AGG(REGEXP_REPLACE(REPLACE(REGEXP_REPLACE(v.field_content,'(\|a|:)','','g'),'|q',' '),'(\|c|\|2|\|d).*?(\||$)',''),', '),'') AS isbns
	 FROM sierra_view.varfield v
	 WHERE brp.bib_record_id = v.record_id
	   AND v.marc_tag IN ('020','024')
	)AS isbns,
	CASE
	  WHEN hd.age_level = 'j' THEN 'JUV'
	  WHEN hd.age_level = 'y' THEN 'YA'
	  WHEN hd.age_level IS NULL THEN 'UNKNOWN'
	  ELSE 'ADULT'
	END AS age_level,
	MODE() WITHIN GROUP (ORDER BY CASE
		WHEN d.index_entry ~ '((\yfiction)|(pictorial works)|(tales)|(novels)|(^\y(?!\w*biography)\w*(comic books strips etc))|(^\y(?!\w*biography)\w*(graphic novels))|(\ydrama)|((?<!hi)stories))(( [a-z]+)?)(( translations into [a-z]+)?)$'
			AND brp.material_code NOT IN ('7','8','b','e','j','k','m','n')
			AND NOT (ml.bib_level_code = 'm'
			AND ml.record_type_code = 'a'
			AND f.p33 IN ('0','e','i','p','s','','c')) THEN 'TRUE'
		WHEN d.index_entry IS NULL THEN 'UNKNOWN'
		ELSE 'FALSE'
	END) AS is_fiction,
	CASE
			WHEN MAX(hd.processing_copies) > 0 AND MAX(hd.in_process_item_count) < MAX(hd.processing_copies) THEN MAX(hd.processing_copies) - MAX(hd.in_process_item_count)
			ELSE 0
	END AS LocalCopiesInProcess,
	MAX(hd.local_speed_item_count) AS LocalSpeedItemCount

FROM sierra_view.bib_record_property brp
	JOIN hold_data hd
		ON brp.bib_record_id = hd.bib_id
	LEFT JOIN sierra_view.bib_record_item_record_link bri
		ON brp.bib_record_id=bri.bib_record_id
	LEFT JOIN sierra_view.item_record ir
		ON bri.item_record_id = ir.id
	JOIN sierra_view.material_property_myuser mp
		ON brp.material_code = mp.code
	LEFT JOIN sierra_view.phrase_entry d
		ON hd.bib_id = d.record_id AND d.index_tag = 'd' AND d.is_permuted = FALSE
	LEFT JOIN sierra_view.leader_field ml
		ON hd.bib_id = ml.record_id
	LEFT JOIN sierra_view.control_field f
		ON hd.bib_id = f.record_id
        AND f.control_num = 8

GROUP BY 1, 2, 3, 4, 5, 6, 8, 12, 14, 15, 16, 17
HAVING hd.local_holds > 0
)a

ORDER BY 5, 
	CASE
		WHEN CAST(a.LocalHoldCount AS NUMERIC(12, 2))/3.0 - a.LocalHoldableItemCount - a.LocalOrderCopies - a.LocalCopiesInProcess < 0 THEN 0
		ELSE CAST(a.LocalHoldCount AS NUMERIC(12, 2))/3.0 - a.LocalHoldableItemCount - a.LocalOrderCopies - a.LocalCopiesInProcess
	END DESC,
	a.LocalRatio DESC
    """
    query_results = run_query(query)
    #Name of Excel File
    excel_file =  "/Scripts/Purchase Alert/Temp Files/LEXPurchaseAlertCustom{}.xlsx".format(date.today())
    excel_writer(query_results,excel_file)
    sftp_file("C:\\Scripts\\Purchase Alert\\Temp Files\\LEXPurchaseAlertCustom{}.xlsx".format(date.today()), 'Lexington')

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
        email_subject = "Lexington Custom Purchase Alert script error"
        email_message = (
            "Your script failed with the following error:\n\n" + traceback.format_exc()
        )

        send_email_error(email_subject, email_message, emailto)
        raise
