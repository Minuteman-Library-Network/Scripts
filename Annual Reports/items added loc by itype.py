#!/usr/bin/env python3

"""
Jeremy Goldstein
Minuteman Library Network

Generates crosstab report of new items in the last year
broken out by shelving location and item type
"""
# run in py38

import psycopg2
import configparser
import xlsxwriter
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

    # Creating the Excel file for staff
    workbook = xlsxwriter.Workbook(excel_file)
    worksheet = workbook.add_worksheet()

    # Formatting our Excel worksheet
    worksheet.set_landscape()
    worksheet.hide_gridlines(0)

    # Formatting Cells
    eformat = workbook.add_format({"text_wrap": True, "valign": "top"})
    eformatlabel = workbook.add_format(
        {"text_wrap": True, "valign": "top", "bold": True, "rotation": -90}
    )

    # Setting the column widths
    worksheet.set_column(0, 0, 49.86)
    worksheet.set_column(1, 1, 16.43)
    worksheet.set_column(2, 2, 16.43)
    worksheet.set_column(3, 3, 16.43)
    worksheet.set_column(4, 4, 16.43)
    worksheet.set_column(5, 5, 16.43)
    worksheet.set_column(6, 6, 16.43)
    worksheet.set_column(7, 7, 16.43)
    worksheet.set_column(8, 8, 16.43)
    worksheet.set_column(9, 9, 16.43)
    worksheet.set_column(10, 10, 16.43)
    worksheet.set_column(11, 11, 16.43)
    worksheet.set_column(12, 12, 16.43)
    worksheet.set_column(13, 13, 16.43)
    worksheet.set_column(14, 14, 16.43)
    worksheet.set_column(15, 15, 16.43)
    worksheet.set_column(16, 16, 16.43)
    worksheet.set_column(17, 17, 16.43)
    worksheet.set_column(18, 18, 16.43)
    worksheet.set_column(19, 19, 16.43)
    worksheet.set_column(20, 20, 16.43)
    worksheet.set_column(21, 21, 16.43)
    worksheet.set_column(22, 22, 16.43)
    worksheet.set_column(23, 23, 16.43)
    worksheet.set_column(24, 24, 16.43)
    worksheet.set_column(25, 25, 16.43)
    worksheet.set_column(26, 26, 16.43)
    worksheet.set_column(27, 27, 16.43)
    worksheet.set_column(28, 28, 16.43)
    worksheet.set_column(29, 29, 16.43)
    worksheet.set_column(30, 30, 16.43)
    worksheet.set_column(31, 31, 16.43)
    worksheet.set_column(32, 32, 16.43)
    worksheet.set_column(33, 33, 16.43)
    worksheet.set_column(34, 34, 16.43)
    worksheet.set_column(35, 35, 16.43)
    worksheet.set_column(36, 36, 16.43)
    worksheet.set_column(37, 37, 16.43)
    worksheet.set_column(38, 38, 16.43)
    worksheet.set_column(39, 39, 16.43)
    worksheet.set_column(40, 40, 16.43)
    worksheet.set_column(41, 41, 16.43)
    worksheet.set_column(42, 42, 16.43)
    worksheet.set_column(43, 43, 16.43)
    worksheet.set_column(44, 44, 16.43)
    worksheet.set_column(45, 45, 16.43)
    worksheet.set_column(46, 46, 16.43)
    worksheet.set_column(47, 47, 16.43)
    worksheet.set_column(48, 48, 16.43)
    worksheet.set_column(49, 49, 16.43)
    worksheet.set_column(50, 50, 16.43)
    worksheet.set_column(51, 51, 16.43)
    worksheet.set_column(52, 52, 16.43)
    worksheet.set_column(53, 53, 16.43)
    worksheet.set_column(54, 54, 16.43)
    worksheet.set_column(55, 55, 16.43)
    worksheet.set_column(56, 56, 16.43)
    worksheet.set_column(57, 57, 16.43)
    worksheet.set_column(58, 58, 16.43)
    worksheet.set_column(59, 59, 16.43)
    worksheet.set_column(60, 60, 16.43)
    worksheet.set_column(61, 61, 16.43)
    worksheet.set_column(62, 62, 16.43)
    worksheet.set_column(63, 63, 16.43)
    worksheet.set_column(64, 64, 16.43)
    worksheet.set_column(65, 65, 16.43)
    worksheet.set_column(66, 66, 16.43)
    worksheet.set_column(67, 67, 16.43)
    worksheet.set_column(68, 68, 16.43)
    worksheet.set_column(69, 69, 16.43)
    worksheet.set_column(70, 70, 16.43)
    worksheet.set_column(71, 71, 16.43)
    worksheet.set_column(72, 72, 16.43)
    worksheet.set_column(73, 73, 16.43)
    worksheet.set_column(74, 74, 16.43)
    worksheet.set_column(75, 75, 16.43)
    worksheet.set_column(76, 76, 16.43)
    worksheet.set_column(77, 77, 16.43)
    worksheet.set_column(78, 78, 16.43)
    worksheet.set_column(79, 79, 16.43)
    worksheet.set_column(80, 80, 16.43)
    worksheet.set_column(81, 81, 16.43)
    worksheet.set_column(82, 82, 16.43)
    worksheet.set_column(83, 83, 16.43)
    worksheet.set_column(84, 84, 16.43)
    worksheet.set_column(85, 85, 16.43)
    worksheet.set_column(86, 86, 16.43)
    worksheet.set_column(87, 87, 16.43)
    worksheet.set_column(88, 88, 16.43)
    worksheet.set_column(89, 89, 16.43)
    worksheet.set_column(90, 90, 16.43)
    worksheet.set_column(91, 91, 16.43)
    worksheet.set_column(92, 92, 16.43)
    worksheet.set_column(93, 93, 16.43)
    worksheet.set_column(94, 94, 16.43)
    worksheet.set_column(95, 95, 16.43)
    worksheet.set_column(96, 96, 16.43)
    worksheet.set_column(97, 97, 16.43)
    worksheet.set_column(98, 98, 16.43)
    worksheet.set_column(99, 99, 16.43)
    worksheet.set_column(100, 100, 16.43)
    worksheet.set_column(101, 101, 16.43)
    worksheet.set_column(102, 102, 16.43)
    worksheet.set_column(103, 103, 16.43)
    worksheet.set_column(104, 104, 16.43)
    worksheet.set_column(105, 105, 16.43)
    worksheet.set_column(106, 106, 16.43)
    worksheet.set_column(107, 107, 16.43)
    worksheet.set_column(108, 108, 16.43)
    worksheet.set_column(109, 109, 16.43)
    worksheet.set_column(110, 110, 16.43)
    worksheet.set_column(111, 111, 16.43)
    worksheet.set_column(112, 112, 16.43)
    worksheet.set_column(113, 113, 16.43)
    worksheet.set_column(114, 114, 16.43)
    worksheet.set_column(115, 115, 16.43)
    worksheet.set_column(116, 116, 16.43)
    worksheet.set_column(117, 117, 16.43)
    worksheet.set_column(118, 118, 16.43)
    worksheet.set_column(119, 119, 16.43)
    worksheet.set_column(120, 120, 16.43)
    worksheet.set_column(121, 121, 16.43)
    worksheet.set_column(122, 122, 16.43)
    worksheet.set_column(123, 123, 16.43)
    worksheet.set_column(124, 124, 16.43)
    worksheet.set_column(125, 125, 16.43)
    worksheet.set_column(126, 126, 16.43)
    worksheet.set_column(127, 127, 16.43)
    worksheet.set_column(128, 128, 16.43)
    worksheet.set_column(129, 129, 16.43)
    worksheet.set_column(130, 130, 16.43)
    worksheet.set_column(131, 131, 16.43)
    worksheet.set_column(132, 132, 16.43)

    # Inserting a header
    worksheet.set_header("Items Added Loc by Itype")

    # Adding column labels
    worksheet.write(0, 0, "Location", eformatlabel)
    worksheet.write(0, 1, "0  Book", eformatlabel)
    worksheet.write(0, 2, "1  Paperback", eformatlabel)
    worksheet.write(0, 3, "2  Large Print", eformatlabel)
    worksheet.write(0, 4, "3  Ref Book", eformatlabel)
    worksheet.write(0, 5, "4  HD Book", eformatlabel)
    worksheet.write(0, 6, "5  Speed Read", eformatlabel)
    worksheet.write(0, 7, "6  Special Collection", eformatlabel)
    worksheet.write(0, 8, "7  Reading List", eformatlabel)
    worksheet.write(0, 9, "8  Book+Disk", eformatlabel)
    worksheet.write(0, 10, "9  Literacy", eformatlabel)
    worksheet.write(0, 11, "10  Periodical", eformatlabel)
    worksheet.write(0, 12, "11  Map", eformatlabel)
    worksheet.write(0, 13, "12  Score", eformatlabel)
    worksheet.write(0, 14, "13  Misc", eformatlabel)
    worksheet.write(0, 15, "19  New TV Series", eformatlabel)
    worksheet.write(0, 16, "20  TV Series", eformatlabel)
    worksheet.write(0, 17, "21  Speed View", eformatlabel)
    worksheet.write(0, 18, "22  Speed View TV Series", eformatlabel)
    worksheet.write(0, 19, "23  NEW DVD F", eformatlabel)
    worksheet.write(0, 20, "25  VHS F", eformatlabel)
    worksheet.write(0, 21, "26  VHS NF", eformatlabel)
    worksheet.write(0, 22, "27  DVD F", eformatlabel)
    worksheet.write(0, 23, "28  DVD NF", eformatlabel)
    worksheet.write(0, 24, "29  VCD F", eformatlabel)
    worksheet.write(0, 25, "30  VCD NF", eformatlabel)
    worksheet.write(0, 26, "31  CD-ROM Ref", eformatlabel)
    worksheet.write(0, 27, "32  CD-ROM Circ", eformatlabel)
    worksheet.write(0, 28, "33  Music CD", eformatlabel)
    worksheet.write(0, 29, "34  Music CD HD", eformatlabel)
    worksheet.write(0, 30, "35  Music Cass", eformatlabel)
    worksheet.write(0, 31, "36  Word CD", eformatlabel)
    worksheet.write(0, 32, "37  Word CD HD", eformatlabel)
    worksheet.write(0, 33, "38  Word Cass", eformatlabel)
    worksheet.write(0, 34, "40  LP Record", eformatlabel)
    worksheet.write(0, 35, "41  Talking Book", eformatlabel)
    worksheet.write(0, 36, "42  Software", eformatlabel)
    worksheet.write(0, 37, "43  Console Game", eformatlabel)
    worksheet.write(0, 38, "44  Film", eformatlabel)
    worksheet.write(0, 39, "45  Slide", eformatlabel)
    worksheet.write(0, 40, "46  Microform", eformatlabel)
    worksheet.write(0, 41, "47  Media Kit", eformatlabel)
    worksheet.write(0, 42, "48  Game", eformatlabel)
    worksheet.write(0, 43, "50  Book on Player", eformatlabel)
    worksheet.write(0, 44, "51  eReader", eformatlabel)
    worksheet.write(0, 45, "52  Video On Player", eformatlabel)
    worksheet.write(0, 46, "100  YA Book", eformatlabel)
    worksheet.write(0, 47, "101  YA Paperback", eformatlabel)
    worksheet.write(0, 48, "102  YA Large Print", eformatlabel)
    worksheet.write(0, 49, "103  YA Ref Book", eformatlabel)
    worksheet.write(0, 50, "104  YA HD Book", eformatlabel)
    worksheet.write(0, 51, "105  YA Read List", eformatlabel)
    worksheet.write(0, 52, "106  YA Book + Disk", eformatlabel)
    worksheet.write(0, 53, "107  YA Magazine", eformatlabel)
    worksheet.write(0, 54, "108  YA Misc", eformatlabel)
    worksheet.write(0, 55, "109  YA Speed Read", eformatlabel)
    worksheet.write(0, 56, "113  YA NEW DVD F", eformatlabel)
    worksheet.write(0, 57, "115  YA VHS F", eformatlabel)
    worksheet.write(0, 58, "116  YA VHS NF", eformatlabel)
    worksheet.write(0, 59, "117  YA DVD F", eformatlabel)
    worksheet.write(0, 60, "118  YA DVD NF", eformatlabel)
    worksheet.write(0, 61, "119  YA VCD F", eformatlabel)
    worksheet.write(0, 62, "120  YA VCD NF", eformatlabel)
    worksheet.write(0, 63, "121  YA CD_ROM Ref", eformatlabel)
    worksheet.write(0, 64, "122  YA CD-ROM Circ", eformatlabel)
    worksheet.write(0, 65, "123  YA Music CD", eformatlabel)
    worksheet.write(0, 66, "124  YA Music Cass", eformatlabel)
    worksheet.write(0, 67, "125  YA Word CD", eformatlabel)
    worksheet.write(0, 68, "126  YA Word Cass", eformatlabel)
    worksheet.write(0, 69, "127  YA Media Kit", eformatlabel)
    worksheet.write(0, 70, "128  YA Game", eformatlabel)
    worksheet.write(0, 71, "129  YA Console Game", eformatlabel)
    worksheet.write(0, 72, "130  YA Book on Player", eformatlabel)
    worksheet.write(0, 73, "131  YA eReader", eformatlabel)
    worksheet.write(0, 74, "132  YA VideoOnPlayer", eformatlabel)
    worksheet.write(0, 75, "133  YA Speed View", eformatlabel)
    worksheet.write(0, 76, "150  J Book", eformatlabel)
    worksheet.write(0, 77, "151  J Paperback", eformatlabel)
    worksheet.write(0, 78, "152  J Holiday", eformatlabel)
    worksheet.write(0, 79, "153  J Large Print", eformatlabel)
    worksheet.write(0, 80, "154  J Ref Book", eformatlabel)
    worksheet.write(0, 81, "155  J HD Book", eformatlabel)
    worksheet.write(0, 82, "156  J Read List", eformatlabel)
    worksheet.write(0, 83, "157  J Book+Disk", eformatlabel)
    worksheet.write(0, 84, "158  J Magazine", eformatlabel)
    worksheet.write(0, 85, "159  J Misc", eformatlabel)
    worksheet.write(0, 86, "160  J Speed Read", eformatlabel)
    worksheet.write(0, 87, "163  J TV Series", eformatlabel)
    worksheet.write(0, 88, "164  J NEW DVD F", eformatlabel)
    worksheet.write(0, 89, "165  J VHS F", eformatlabel)
    worksheet.write(0, 90, "166  J VHS NF", eformatlabel)
    worksheet.write(0, 91, "167  J DVD F", eformatlabel)
    worksheet.write(0, 92, "168  J DVD NF", eformatlabel)
    worksheet.write(0, 93, "169  J CD-ROM Ref", eformatlabel)
    worksheet.write(0, 94, "170  J CD-ROM Circ", eformatlabel)
    worksheet.write(0, 95, "171  J Music CD", eformatlabel)
    worksheet.write(0, 96, "172  J Music Cass", eformatlabel)
    worksheet.write(0, 97, "173  J Word CD", eformatlabel)
    worksheet.write(0, 98, "174  J Word Cass", eformatlabel)
    worksheet.write(0, 99, "175  J Record", eformatlabel)
    worksheet.write(0, 100, "176  J Media Kit", eformatlabel)
    worksheet.write(0, 101, "177  J Game/Toy", eformatlabel)
    worksheet.write(0, 102, "178  J Console Game", eformatlabel)
    worksheet.write(0, 103, "179  J Filmstrip", eformatlabel)
    worksheet.write(0, 104, "180  J Book on Player", eformatlabel)
    worksheet.write(0, 105, "181  J eReader", eformatlabel)
    worksheet.write(0, 106, "182  J VideoOnPlayer", eformatlabel)
    worksheet.write(0, 107, "183  J Speed View", eformatlabel)
    worksheet.write(0, 108, "186  J Equipment1", eformatlabel)
    worksheet.write(0, 109, "187  J Equipment2", eformatlabel)
    worksheet.write(0, 110, "188  J Equipment3", eformatlabel)
    worksheet.write(0, 111, "189  J Equipment4", eformatlabel)
    worksheet.write(0, 112, "221  Reserve 1", eformatlabel)
    worksheet.write(0, 113, "222  Reserve 2", eformatlabel)
    worksheet.write(0, 114, "223  Reserve 3", eformatlabel)
    worksheet.write(0, 115, "224  Reserve 4", eformatlabel)
    worksheet.write(0, 116, "239  ComCat", eformatlabel)
    worksheet.write(0, 117, "241  Non-MLN ILL", eformatlabel)
    worksheet.write(0, 118, "242  On Order", eformatlabel)
    worksheet.write(0, 119, "243  Museum Pass", eformatlabel)
    worksheet.write(0, 120, "244  Subscribed DB", eformatlabel)
    worksheet.write(0, 121, "245  Equipment1", eformatlabel)
    worksheet.write(0, 122, "246  HD Equipment", eformatlabel)
    worksheet.write(0, 123, "247  Art", eformatlabel)
    worksheet.write(0, 124, "248  E-Resource", eformatlabel)
    worksheet.write(0, 125, "249  OverDrive ADV", eformatlabel)
    worksheet.write(0, 126, "250  In-House Equipment", eformatlabel)
    worksheet.write(0, 127, "251  Equipment2", eformatlabel)
    worksheet.write(0, 128, "252  Equipment3", eformatlabel)
    worksheet.write(0, 129, "253  Equipment4", eformatlabel)
    worksheet.write(0, 130, "255  Unknown", eformatlabel)
    worksheet.write(0, 131, "256  Hotspot", eformatlabel)
    worksheet.write(0, 132, "257  Hotspot HD", eformatlabel)

    # Writing the report for staff to the Excel worksheet
    for rownum, row in enumerate(query_results):
        worksheet.write(rownum + 1, 0, row[0], eformat)
        worksheet.write(rownum + 1, 1, row[1], eformat)
        worksheet.write(rownum + 1, 2, row[2], eformat)
        worksheet.write(rownum + 1, 3, row[3], eformat)
        worksheet.write(rownum + 1, 4, row[4], eformat)
        worksheet.write(rownum + 1, 5, row[5], eformat)
        worksheet.write(rownum + 1, 6, row[6], eformat)
        worksheet.write(rownum + 1, 7, row[7], eformat)
        worksheet.write(rownum + 1, 8, row[8], eformat)
        worksheet.write(rownum + 1, 9, row[9], eformat)
        worksheet.write(rownum + 1, 10, row[10], eformat)
        worksheet.write(rownum + 1, 11, row[11], eformat)
        worksheet.write(rownum + 1, 12, row[12], eformat)
        worksheet.write(rownum + 1, 13, row[13], eformat)
        worksheet.write(rownum + 1, 14, row[14], eformat)
        worksheet.write(rownum + 1, 15, row[15], eformat)
        worksheet.write(rownum + 1, 16, row[16], eformat)
        worksheet.write(rownum + 1, 17, row[17], eformat)
        worksheet.write(rownum + 1, 18, row[18], eformat)
        worksheet.write(rownum + 1, 19, row[19], eformat)
        worksheet.write(rownum + 1, 20, row[20], eformat)
        worksheet.write(rownum + 1, 21, row[21], eformat)
        worksheet.write(rownum + 1, 22, row[22], eformat)
        worksheet.write(rownum + 1, 23, row[23], eformat)
        worksheet.write(rownum + 1, 24, row[24], eformat)
        worksheet.write(rownum + 1, 25, row[25], eformat)
        worksheet.write(rownum + 1, 26, row[26], eformat)
        worksheet.write(rownum + 1, 27, row[27], eformat)
        worksheet.write(rownum + 1, 28, row[28], eformat)
        worksheet.write(rownum + 1, 29, row[29], eformat)
        worksheet.write(rownum + 1, 30, row[30], eformat)
        worksheet.write(rownum + 1, 31, row[31], eformat)
        worksheet.write(rownum + 1, 32, row[32], eformat)
        worksheet.write(rownum + 1, 33, row[33], eformat)
        worksheet.write(rownum + 1, 34, row[34], eformat)
        worksheet.write(rownum + 1, 35, row[35], eformat)
        worksheet.write(rownum + 1, 36, row[36], eformat)
        worksheet.write(rownum + 1, 37, row[37], eformat)
        worksheet.write(rownum + 1, 38, row[38], eformat)
        worksheet.write(rownum + 1, 39, row[39], eformat)
        worksheet.write(rownum + 1, 40, row[40], eformat)
        worksheet.write(rownum + 1, 41, row[41], eformat)
        worksheet.write(rownum + 1, 42, row[42], eformat)
        worksheet.write(rownum + 1, 43, row[43], eformat)
        worksheet.write(rownum + 1, 44, row[44], eformat)
        worksheet.write(rownum + 1, 45, row[45], eformat)
        worksheet.write(rownum + 1, 46, row[46], eformat)
        worksheet.write(rownum + 1, 47, row[47], eformat)
        worksheet.write(rownum + 1, 48, row[48], eformat)
        worksheet.write(rownum + 1, 49, row[49], eformat)
        worksheet.write(rownum + 1, 50, row[50], eformat)
        worksheet.write(rownum + 1, 51, row[51], eformat)
        worksheet.write(rownum + 1, 52, row[52], eformat)
        worksheet.write(rownum + 1, 53, row[53], eformat)
        worksheet.write(rownum + 1, 54, row[54], eformat)
        worksheet.write(rownum + 1, 55, row[55], eformat)
        worksheet.write(rownum + 1, 56, row[56], eformat)
        worksheet.write(rownum + 1, 57, row[57], eformat)
        worksheet.write(rownum + 1, 58, row[58], eformat)
        worksheet.write(rownum + 1, 59, row[59], eformat)
        worksheet.write(rownum + 1, 60, row[60], eformat)
        worksheet.write(rownum + 1, 61, row[61], eformat)
        worksheet.write(rownum + 1, 62, row[62], eformat)
        worksheet.write(rownum + 1, 63, row[63], eformat)
        worksheet.write(rownum + 1, 64, row[64], eformat)
        worksheet.write(rownum + 1, 65, row[65], eformat)
        worksheet.write(rownum + 1, 66, row[66], eformat)
        worksheet.write(rownum + 1, 67, row[67], eformat)
        worksheet.write(rownum + 1, 68, row[68], eformat)
        worksheet.write(rownum + 1, 69, row[69], eformat)
        worksheet.write(rownum + 1, 70, row[70], eformat)
        worksheet.write(rownum + 1, 71, row[71], eformat)
        worksheet.write(rownum + 1, 72, row[72], eformat)
        worksheet.write(rownum + 1, 73, row[73], eformat)
        worksheet.write(rownum + 1, 74, row[74], eformat)
        worksheet.write(rownum + 1, 75, row[75], eformat)
        worksheet.write(rownum + 1, 76, row[76], eformat)
        worksheet.write(rownum + 1, 77, row[77], eformat)
        worksheet.write(rownum + 1, 78, row[78], eformat)
        worksheet.write(rownum + 1, 79, row[79], eformat)
        worksheet.write(rownum + 1, 80, row[80], eformat)
        worksheet.write(rownum + 1, 81, row[81], eformat)
        worksheet.write(rownum + 1, 82, row[82], eformat)
        worksheet.write(rownum + 1, 83, row[83], eformat)
        worksheet.write(rownum + 1, 84, row[84], eformat)
        worksheet.write(rownum + 1, 85, row[85], eformat)
        worksheet.write(rownum + 1, 86, row[86], eformat)
        worksheet.write(rownum + 1, 87, row[87], eformat)
        worksheet.write(rownum + 1, 88, row[88], eformat)
        worksheet.write(rownum + 1, 89, row[89], eformat)
        worksheet.write(rownum + 1, 90, row[90], eformat)
        worksheet.write(rownum + 1, 91, row[91], eformat)
        worksheet.write(rownum + 1, 92, row[92], eformat)
        worksheet.write(rownum + 1, 93, row[93], eformat)
        worksheet.write(rownum + 1, 94, row[94], eformat)
        worksheet.write(rownum + 1, 95, row[95], eformat)
        worksheet.write(rownum + 1, 96, row[96], eformat)
        worksheet.write(rownum + 1, 97, row[97], eformat)
        worksheet.write(rownum + 1, 98, row[98], eformat)
        worksheet.write(rownum + 1, 99, row[99], eformat)
        worksheet.write(rownum + 1, 100, row[100], eformat)
        worksheet.write(rownum + 1, 101, row[101], eformat)
        worksheet.write(rownum + 1, 102, row[102], eformat)
        worksheet.write(rownum + 1, 103, row[103], eformat)
        worksheet.write(rownum + 1, 104, row[104], eformat)
        worksheet.write(rownum + 1, 105, row[105], eformat)
        worksheet.write(rownum + 1, 106, row[106], eformat)
        worksheet.write(rownum + 1, 107, row[107], eformat)
        worksheet.write(rownum + 1, 108, row[108], eformat)
        worksheet.write(rownum + 1, 109, row[109], eformat)
        worksheet.write(rownum + 1, 110, row[110], eformat)
        worksheet.write(rownum + 1, 111, row[111], eformat)
        worksheet.write(rownum + 1, 112, row[112], eformat)
        worksheet.write(rownum + 1, 113, row[113], eformat)
        worksheet.write(rownum + 1, 114, row[114], eformat)
        worksheet.write(rownum + 1, 115, row[115], eformat)
        worksheet.write(rownum + 1, 116, row[116], eformat)
        worksheet.write(rownum + 1, 117, row[117], eformat)
        worksheet.write(rownum + 1, 118, row[118], eformat)
        worksheet.write(rownum + 1, 119, row[119], eformat)
        worksheet.write(rownum + 1, 120, row[120], eformat)
        worksheet.write(rownum + 1, 121, row[121], eformat)
        worksheet.write(rownum + 1, 122, row[122], eformat)
        worksheet.write(rownum + 1, 123, row[123], eformat)
        worksheet.write(rownum + 1, 124, row[124], eformat)
        worksheet.write(rownum + 1, 125, row[125], eformat)
        worksheet.write(rownum + 1, 126, row[126], eformat)
        worksheet.write(rownum + 1, 127, row[127], eformat)
        worksheet.write(rownum + 1, 128, row[128], eformat)
        worksheet.write(rownum + 1, 129, row[129], eformat)
        worksheet.write(rownum + 1, 130, row[130], eformat)
        worksheet.write(rownum + 1, 131, row[131], eformat)
        worksheet.write(rownum + 1, 132, row[132], eformat)

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
    query = """\
            SELECT
              i.location_code||'    '||l.name AS location,
              COUNT(i.id) FILTER(WHERE i.itype_code_num = 0) AS "0",
              COUNT(i.id) FILTER(WHERE i.itype_code_num = 1) AS "1",
              COUNT(i.id) FILTER(WHERE i.itype_code_num = 2) AS "2",
              COUNT(i.id) FILTER(WHERE i.itype_code_num = 3) AS "3",
              COUNT(i.id) FILTER(WHERE i.itype_code_num = 4) AS "4",
              COUNT(i.id) FILTER(WHERE i.itype_code_num = 5) AS "5",
              COUNT(i.id) FILTER(WHERE i.itype_code_num = 6) AS "6",
              COUNT(i.id) FILTER(WHERE i.itype_code_num = 7) AS "7",
              COUNT(i.id) FILTER(WHERE i.itype_code_num = 8) AS "8",
              COUNT(i.id) FILTER(WHERE i.itype_code_num = 9) AS "9",
              COUNT(i.id) FILTER(WHERE i.itype_code_num = 10) AS "10",
              COUNT(i.id) FILTER(WHERE i.itype_code_num = 11) AS "11",
              COUNT(i.id) FILTER(WHERE i.itype_code_num = 12) AS "12",
              COUNT(i.id) FILTER(WHERE i.itype_code_num = 13) AS "13",
              COUNT(i.id) FILTER(WHERE i.itype_code_num = 19) AS "19",
              COUNT(i.id) FILTER(WHERE i.itype_code_num = 20) AS "20",
              COUNT(i.id) FILTER(WHERE i.itype_code_num = 21) AS "21",
              COUNT(i.id) FILTER(WHERE i.itype_code_num = 22) AS "22",
              COUNT(i.id) FILTER(WHERE i.itype_code_num = 23) AS "23",
              COUNT(i.id) FILTER(WHERE i.itype_code_num = 25) AS "25",
              COUNT(i.id) FILTER(WHERE i.itype_code_num = 26) AS "26",
              COUNT(i.id) FILTER(WHERE i.itype_code_num = 27) AS "27",
              COUNT(i.id) FILTER(WHERE i.itype_code_num = 28) AS "28",
              COUNT(i.id) FILTER(WHERE i.itype_code_num = 29) AS "29",
              COUNT(i.id) FILTER(WHERE i.itype_code_num = 30) AS "30",
              COUNT(i.id) FILTER(WHERE i.itype_code_num = 31) AS "31",
              COUNT(i.id) FILTER(WHERE i.itype_code_num = 32) AS "32",
              COUNT(i.id) FILTER(WHERE i.itype_code_num = 33) AS "33",
              COUNT(i.id) FILTER(WHERE i.itype_code_num = 34) AS "34",
              COUNT(i.id) FILTER(WHERE i.itype_code_num = 35) AS "35",
              COUNT(i.id) FILTER(WHERE i.itype_code_num = 36) AS "36",
              COUNT(i.id) FILTER(WHERE i.itype_code_num = 37) AS "37",
              COUNT(i.id) FILTER(WHERE i.itype_code_num = 38) AS "38",
              COUNT(i.id) FILTER(WHERE i.itype_code_num = 40) AS "40",
              COUNT(i.id) FILTER(WHERE i.itype_code_num = 41) AS "41",
              COUNT(i.id) FILTER(WHERE i.itype_code_num = 42) AS "42",
              COUNT(i.id) FILTER(WHERE i.itype_code_num = 43) AS "43",
              COUNT(i.id) FILTER(WHERE i.itype_code_num = 44) AS "44",
              COUNT(i.id) FILTER(WHERE i.itype_code_num = 45) AS "45",
              COUNT(i.id) FILTER(WHERE i.itype_code_num = 46) AS "46",
              COUNT(i.id) FILTER(WHERE i.itype_code_num = 47) AS "47",
              COUNT(i.id) FILTER(WHERE i.itype_code_num = 48) AS "48",
              COUNT(i.id) FILTER(WHERE i.itype_code_num = 50) AS "50",
              COUNT(i.id) FILTER(WHERE i.itype_code_num = 51) AS "51",
              COUNT(i.id) FILTER(WHERE i.itype_code_num = 52) AS "52",
              COUNT(i.id) FILTER(WHERE i.itype_code_num = 100) AS "100",
              COUNT(i.id) FILTER(WHERE i.itype_code_num = 101) AS "101",
              COUNT(i.id) FILTER(WHERE i.itype_code_num = 102) AS "102",
              COUNT(i.id) FILTER(WHERE i.itype_code_num = 103) AS "103",
              COUNT(i.id) FILTER(WHERE i.itype_code_num = 104) AS "104",
              COUNT(i.id) FILTER(WHERE i.itype_code_num = 105) AS "105",
              COUNT(i.id) FILTER(WHERE i.itype_code_num = 106) AS "106",
              COUNT(i.id) FILTER(WHERE i.itype_code_num = 107) AS "107",
              COUNT(i.id) FILTER(WHERE i.itype_code_num = 108) AS "108",
              COUNT(i.id) FILTER(WHERE i.itype_code_num = 109) AS "109",
              COUNT(i.id) FILTER(WHERE i.itype_code_num = 113) AS "113",
              COUNT(i.id) FILTER(WHERE i.itype_code_num = 115) AS "115",
              COUNT(i.id) FILTER(WHERE i.itype_code_num = 116) AS "116",
              COUNT(i.id) FILTER(WHERE i.itype_code_num = 117) AS "117",
              COUNT(i.id) FILTER(WHERE i.itype_code_num = 118) AS "118",
              COUNT(i.id) FILTER(WHERE i.itype_code_num = 119) AS "119",
              COUNT(i.id) FILTER(WHERE i.itype_code_num = 120) AS "120",
              COUNT(i.id) FILTER(WHERE i.itype_code_num = 121) AS "121",
              COUNT(i.id) FILTER(WHERE i.itype_code_num = 122) AS "122",
              COUNT(i.id) FILTER(WHERE i.itype_code_num = 123) AS "123",
              COUNT(i.id) FILTER(WHERE i.itype_code_num = 124) AS "124",
              COUNT(i.id) FILTER(WHERE i.itype_code_num = 125) AS "125",
              COUNT(i.id) FILTER(WHERE i.itype_code_num = 126) AS "126",
              COUNT(i.id) FILTER(WHERE i.itype_code_num = 127) AS "127",
              COUNT(i.id) FILTER(WHERE i.itype_code_num = 128) AS "128",
              COUNT(i.id) FILTER(WHERE i.itype_code_num = 129) AS "129",
              COUNT(i.id) FILTER(WHERE i.itype_code_num = 130) AS "130",
              COUNT(i.id) FILTER(WHERE i.itype_code_num = 131) AS "131",
              COUNT(i.id) FILTER(WHERE i.itype_code_num = 132) AS "132",
              COUNT(i.id) FILTER(WHERE i.itype_code_num = 133) AS "133",
              COUNT(i.id) FILTER(WHERE i.itype_code_num = 150) AS "150",
              COUNT(i.id) FILTER(WHERE i.itype_code_num = 151) AS "151",
              COUNT(i.id) FILTER(WHERE i.itype_code_num = 152) AS "152",
              COUNT(i.id) FILTER(WHERE i.itype_code_num = 153) AS "153",
              COUNT(i.id) FILTER(WHERE i.itype_code_num = 154) AS "154",
              COUNT(i.id) FILTER(WHERE i.itype_code_num = 155) AS "155",
              COUNT(i.id) FILTER(WHERE i.itype_code_num = 156) AS "156",
              COUNT(i.id) FILTER(WHERE i.itype_code_num = 157) AS "157",
              COUNT(i.id) FILTER(WHERE i.itype_code_num = 158) AS "158",
              COUNT(i.id) FILTER(WHERE i.itype_code_num = 159) AS "159",
              COUNT(i.id) FILTER(WHERE i.itype_code_num = 160) AS "160",
              COUNT(i.id) FILTER(WHERE i.itype_code_num = 163) AS "163",
              COUNT(i.id) FILTER(WHERE i.itype_code_num = 164) AS "164",
              COUNT(i.id) FILTER(WHERE i.itype_code_num = 165) AS "165",
              COUNT(i.id) FILTER(WHERE i.itype_code_num = 166) AS "166",
              COUNT(i.id) FILTER(WHERE i.itype_code_num = 167) AS "167",
              COUNT(i.id) FILTER(WHERE i.itype_code_num = 168) AS "168",
              COUNT(i.id) FILTER(WHERE i.itype_code_num = 169) AS "169",
              COUNT(i.id) FILTER(WHERE i.itype_code_num = 170) AS "170",
              COUNT(i.id) FILTER(WHERE i.itype_code_num = 171) AS "171",
              COUNT(i.id) FILTER(WHERE i.itype_code_num = 172) AS "172",
              COUNT(i.id) FILTER(WHERE i.itype_code_num = 173) AS "173",
              COUNT(i.id) FILTER(WHERE i.itype_code_num = 174) AS "174",
              COUNT(i.id) FILTER(WHERE i.itype_code_num = 175) AS "175",
              COUNT(i.id) FILTER(WHERE i.itype_code_num = 176) AS "176",
              COUNT(i.id) FILTER(WHERE i.itype_code_num = 177) AS "177",
              COUNT(i.id) FILTER(WHERE i.itype_code_num = 178) AS "178",
              COUNT(i.id) FILTER(WHERE i.itype_code_num = 179) AS "179",
              COUNT(i.id) FILTER(WHERE i.itype_code_num = 180) AS "180",
              COUNT(i.id) FILTER(WHERE i.itype_code_num = 181) AS "181",
              COUNT(i.id) FILTER(WHERE i.itype_code_num = 182) AS "182",
              COUNT(i.id) FILTER(WHERE i.itype_code_num = 183) AS "183",
              COUNT(i.id) FILTER(WHERE i.itype_code_num = 183) AS "186",
              COUNT(i.id) FILTER(WHERE i.itype_code_num = 183) AS "187",
              COUNT(i.id) FILTER(WHERE i.itype_code_num = 183) AS "188",
              COUNT(i.id) FILTER(WHERE i.itype_code_num = 183) AS "189",
              COUNT(i.id) FILTER(WHERE i.itype_code_num = 221) AS "221",
              COUNT(i.id) FILTER(WHERE i.itype_code_num = 222) AS "222",
              COUNT(i.id) FILTER(WHERE i.itype_code_num = 223) AS "223",
              COUNT(i.id) FILTER(WHERE i.itype_code_num = 224) AS "224",
              COUNT(i.id) FILTER(WHERE i.itype_code_num = 239) AS "239",
              COUNT(i.id) FILTER(WHERE i.itype_code_num = 241) AS "241",
              COUNT(i.id) FILTER(WHERE i.itype_code_num = 242) AS "242",
              COUNT(i.id) FILTER(WHERE i.itype_code_num = 243) AS "243",
              COUNT(i.id) FILTER(WHERE i.itype_code_num = 244) AS "244",
              COUNT(i.id) FILTER(WHERE i.itype_code_num = 245) AS "245",
              COUNT(i.id) FILTER(WHERE i.itype_code_num = 246) AS "246",
              COUNT(i.id) FILTER(WHERE i.itype_code_num = 247) AS "247",
              COUNT(i.id) FILTER(WHERE i.itype_code_num = 248) AS "248",
              COUNT(i.id) FILTER(WHERE i.itype_code_num = 249) AS "249",
              COUNT(i.id) FILTER(WHERE i.itype_code_num = 250) AS "250",
              COUNT(i.id) FILTER(WHERE i.itype_code_num = 251) AS "251",
              COUNT(i.id) FILTER(WHERE i.itype_code_num = 252) AS "252",
              COUNT(i.id) FILTER(WHERE i.itype_code_num = 253) AS "253",
              COUNT(i.id) FILTER(WHERE i.itype_code_num = 255) AS "255",
              COUNT(i.id) FILTER(WHERE i.itype_code_num = 255) AS "256",
              COUNT(i.id) FILTER(WHERE i.itype_code_num = 255) AS "257"

            FROM sierra_view.item_record as i
            JOIN sierra_view.location_myuser l
              ON i.location_code = l.code
            JOIN sierra_view.record_metadata rm
              ON i.id = rm.id

            WHERE rm.creation_date_gmt::date >= (localtimestamp::date - INTERVAL '1 year')
            GROUP BY 1
            ORDER BY 1
            """
    query_results = run_query(query)

    # generate excel file from those query results
    excel_file = (
        "/Scripts/Annual Reports/Archive/items added loc by itype{}.xlsx".format(
            date.today()
        )
    )
    excel_writer(query_results, excel_file)

    # send email with attached file
    email_subject = "Items Added Loc By IType"
    email_message = """***This is an automated email***


The Items Added Loc By Itype report has been attached."""
    send_email(email_subject, email_message, excel_file)


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
        email_subject = "annual reports: items added loc by itype script error"
        email_message = (
            "Your script failed with the following error:\n\n" + traceback.format_exc()
        )

        send_email_error(email_subject, email_message, emailto)
        raise
