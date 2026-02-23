#!/usr/bin/env python3

"""
Jeremy Goldstein
Minuteman Library Network

Generates holdings profile report for use with annual reports
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
    workbook = xlsxwriter.Workbook(excel_file, {"remove_timezone": True})
    worksheet = workbook.add_worksheet()

    # #Formatting our Excel worksheet
    worksheet.set_landscape()
    worksheet.hide_gridlines(0)

    # Formatting Cells
    eformat = workbook.add_format(
        {"text_wrap": True, "valign": "top", "align": "center"}
    )
    eformatlabel = workbook.add_format(
        {"text_wrap": True, "valign": "top", "bold": True, "align": "center"}
    )

    # Setting the column widths
    worksheet.set_column(0, 0, 9.14)
    worksheet.set_column(1, 1, 28)
    worksheet.set_column(2, 2, 7.29)
    worksheet.set_column(3, 3, 7.29)
    worksheet.set_column(4, 4, 7.29)
    worksheet.set_column(5, 5, 7.29)
    worksheet.set_column(6, 6, 7.29)
    worksheet.set_column(7, 7, 7.29)
    worksheet.set_column(8, 8, 7.29)
    worksheet.set_column(9, 9, 7.29)
    worksheet.set_column(10, 10, 7.29)
    worksheet.set_column(11, 11, 7.29)
    worksheet.set_column(12, 12, 7.29)
    worksheet.set_column(13, 13, 7.29)
    worksheet.set_column(14, 14, 7.29)
    worksheet.set_column(15, 15, 7.29)
    worksheet.set_column(16, 16, 7.29)
    worksheet.set_column(17, 17, 7.29)
    worksheet.set_column(18, 18, 7.29)
    worksheet.set_column(19, 19, 7.29)
    worksheet.set_column(20, 20, 7.29)
    worksheet.set_column(21, 21, 7.29)
    worksheet.set_column(22, 22, 7.29)
    worksheet.set_column(23, 23, 7.29)
    worksheet.set_column(24, 24, 7.29)
    worksheet.set_column(25, 25, 7.29)
    worksheet.set_column(26, 26, 7.29)
    worksheet.set_column(27, 27, 7.29)
    worksheet.set_column(28, 28, 7.29)
    worksheet.set_column(29, 29, 7.29)
    worksheet.set_column(30, 30, 7.29)
    worksheet.set_column(31, 31, 7.29)
    worksheet.set_column(32, 32, 7.29)
    worksheet.set_column(33, 33, 7.29)
    worksheet.set_column(34, 34, 7.29)
    worksheet.set_column(35, 35, 7.29)
    worksheet.set_column(36, 36, 7.29)
    worksheet.set_column(37, 37, 7.29)
    worksheet.set_column(38, 38, 7.29)
    worksheet.set_column(39, 39, 7.29)
    worksheet.set_column(40, 40, 7.29)
    worksheet.set_column(41, 41, 7.29)
    worksheet.set_column(42, 42, 7.29)
    worksheet.set_column(43, 43, 7.29)
    worksheet.set_column(44, 44, 7.29)
    worksheet.set_column(45, 45, 7.29)
    worksheet.set_column(46, 46, 7.29)
    worksheet.set_column(47, 47, 7.29)
    worksheet.set_column(48, 48, 7.29)
    worksheet.set_column(49, 49, 7.29)
    worksheet.set_column(50, 50, 7.29)
    worksheet.set_column(51, 51, 7.29)
    worksheet.set_column(52, 52, 7.29)
    worksheet.set_column(53, 53, 7.29)
    worksheet.set_column(54, 54, 7.29)
    worksheet.set_column(55, 55, 7.29)
    worksheet.set_column(56, 56, 7.29)
    worksheet.set_column(57, 57, 7.29)
    worksheet.set_column(58, 58, 7.29)
    worksheet.set_column(59, 59, 7.29)
    worksheet.set_column(60, 60, 7.29)
    worksheet.set_column(61, 61, 7.29)
    worksheet.set_column(62, 62, 7.29)
    worksheet.set_column(63, 63, 7.29)
    worksheet.set_column(64, 64, 7.29)
    worksheet.set_column(65, 65, 7.29)
    worksheet.set_column(66, 66, 7.29)
    worksheet.set_column(67, 67, 7.29)
    worksheet.set_column(68, 68, 7.29)
    worksheet.set_column(69, 69, 7.29)
    worksheet.set_column(70, 70, 7.29)
    worksheet.set_column(71, 71, 7.29)
    worksheet.set_column(72, 72, 7.29)
    worksheet.set_column(73, 73, 7.29)
    worksheet.set_column(74, 74, 7.29)
    worksheet.set_column(75, 75, 7.29)
    worksheet.set_column(76, 76, 7.29)
    worksheet.set_column(77, 77, 7.29)
    worksheet.set_column(78, 78, 7.29)
    worksheet.set_column(79, 79, 7.29)
    worksheet.set_column(80, 80, 7.29)
    worksheet.set_column(81, 81, 7.29)
    worksheet.set_column(82, 82, 7.29)
    worksheet.set_column(83, 83, 7.29)
    worksheet.set_column(84, 84, 7.29)
    worksheet.set_column(85, 85, 7.29)
    worksheet.set_column(86, 86, 7.29)
    worksheet.set_column(87, 87, 7.29)
    worksheet.set_column(88, 88, 7.29)
    worksheet.set_column(89, 89, 7.29)
    worksheet.set_column(90, 90, 7.29)
    worksheet.set_column(91, 91, 7.29)
    worksheet.set_column(92, 92, 7.29)
    worksheet.set_column(93, 93, 7.29)
    worksheet.set_column(94, 94, 7.29)
    worksheet.set_column(95, 95, 7.29)
    worksheet.set_column(96, 96, 7.29)
    worksheet.set_column(97, 97, 7.29)
    worksheet.set_column(98, 98, 7.29)
    worksheet.set_column(99, 99, 7.29)
    worksheet.set_column(100, 100, 7.29)
    worksheet.set_column(101, 101, 7.29)
    worksheet.set_column(102, 102, 7.29)
    worksheet.set_column(103, 103, 7.29)
    worksheet.set_column(104, 104, 7.29)
    worksheet.set_column(105, 105, 7.29)
    worksheet.set_column(106, 106, 7.29)
    worksheet.set_column(107, 107, 7.29)
    worksheet.set_column(108, 108, 7.29)
    worksheet.set_column(109, 109, 7.29)
    worksheet.set_column(110, 110, 7.29)
    worksheet.set_column(111, 111, 7.29)
    worksheet.set_column(112, 112, 7.29)
    worksheet.set_column(113, 113, 7.29)
    worksheet.set_column(114, 114, 7.29)
    worksheet.set_column(115, 115, 7.29)
    worksheet.set_column(116, 116, 7.29)
    worksheet.set_column(117, 117, 7.29)
    worksheet.set_column(118, 118, 7.29)
    worksheet.set_column(119, 119, 7.29)
    worksheet.set_column(120, 120, 7.29)
    worksheet.set_column(121, 121, 7.29)
    worksheet.set_column(122, 122, 7.29)
    worksheet.set_column(123, 123, 7.29)
    worksheet.set_column(124, 124, 7.29)
    worksheet.set_column(125, 125, 7.29)
    worksheet.set_column(126, 126, 7.29)
    worksheet.set_column(127, 127, 7.29)
    worksheet.set_column(128, 128, 7.29)
    worksheet.set_column(129, 129, 7.29)
    worksheet.set_column(130, 130, 7.29)
    worksheet.set_column(131, 131, 7.29)
    worksheet.set_column(132, 132, 7.29)
    worksheet.set_column(133, 133, 7.29)

    # Inserting a header
    worksheet.set_header("Holdings Profile")

    # Adding column labels
    worksheet.write(0, 0, "-", eformatlabel)
    worksheet.write(0, 1, "-", eformatlabel)
    worksheet.write(0, 2, "H1", eformatlabel)
    worksheet.write(0, 3, "H1", eformatlabel)
    worksheet.write(0, 4, "H1", eformatlabel)
    worksheet.write(0, 5, "H1", eformatlabel)
    worksheet.write(0, 6, "H1", eformatlabel)
    worksheet.write(0, 7, "H1", eformatlabel)
    worksheet.write(0, 8, "H1", eformatlabel)
    worksheet.write(0, 9, "H1", eformatlabel)
    worksheet.write(0, 10, "H1", eformatlabel)
    worksheet.write(0, 11, "H1", eformatlabel)
    worksheet.write(0, 12, "*", eformatlabel)
    worksheet.write(0, 13, "H7", eformatlabel)
    worksheet.write(0, 14, "H1", eformatlabel)
    worksheet.write(0, 15, "H7", eformatlabel)
    worksheet.write(0, 16, "H4", eformatlabel)
    worksheet.write(0, 17, "H4", eformatlabel)
    worksheet.write(0, 18, "H4", eformatlabel)
    worksheet.write(0, 19, "H4", eformatlabel)
    worksheet.write(0, 20, "H4", eformatlabel)
    worksheet.write(0, 21, "H4", eformatlabel)
    worksheet.write(0, 22, "H4", eformatlabel)
    worksheet.write(0, 23, "H4", eformatlabel)
    worksheet.write(0, 24, "H4", eformatlabel)
    worksheet.write(0, 25, "H4", eformatlabel)
    worksheet.write(0, 26, "H4", eformatlabel)
    worksheet.write(0, 27, "H5", eformatlabel)
    worksheet.write(0, 28, "H5", eformatlabel)
    worksheet.write(0, 29, "H3", eformatlabel)
    worksheet.write(0, 30, "H3", eformatlabel)
    worksheet.write(0, 31, "H3", eformatlabel)
    worksheet.write(0, 32, "H3", eformatlabel)
    worksheet.write(0, 33, "H3", eformatlabel)
    worksheet.write(0, 34, "H3", eformatlabel)
    worksheet.write(0, 35, "H3", eformatlabel)
    worksheet.write(0, 36, "H3", eformatlabel)
    worksheet.write(0, 37, "H5", eformatlabel)
    worksheet.write(0, 38, "H5", eformatlabel)
    worksheet.write(0, 39, "H7", eformatlabel)
    worksheet.write(0, 40, "H7", eformatlabel)
    worksheet.write(0, 41, "H6", eformatlabel)
    worksheet.write(0, 42, "H7", eformatlabel)
    worksheet.write(0, 43, "H7", eformatlabel)
    worksheet.write(0, 44, "H3", eformatlabel)
    worksheet.write(0, 45, "H7", eformatlabel)
    worksheet.write(0, 46, "H4", eformatlabel)
    worksheet.write(0, 47, "H9", eformatlabel)
    worksheet.write(0, 48, "H9", eformatlabel)
    worksheet.write(0, 49, "H9", eformatlabel)
    worksheet.write(0, 50, "H9", eformatlabel)
    worksheet.write(0, 51, "H9", eformatlabel)
    worksheet.write(0, 52, "H9", eformatlabel)
    worksheet.write(0, 53, "H9", eformatlabel)
    worksheet.write(0, 54, "*", eformatlabel)
    worksheet.write(0, 55, "H15", eformatlabel)
    worksheet.write(0, 56, "H9", eformatlabel)
    worksheet.write(0, 57, "H12", eformatlabel)
    worksheet.write(0, 58, "H12", eformatlabel)
    worksheet.write(0, 59, "H12", eformatlabel)
    worksheet.write(0, 60, "H12", eformatlabel)
    worksheet.write(0, 61, "H12", eformatlabel)
    worksheet.write(0, 62, "H12", eformatlabel)
    worksheet.write(0, 63, "H12", eformatlabel)
    worksheet.write(0, 64, "H13", eformatlabel)
    worksheet.write(0, 65, "H13", eformatlabel)
    worksheet.write(0, 66, "H11", eformatlabel)
    worksheet.write(0, 67, "H11", eformatlabel)
    worksheet.write(0, 68, "H11", eformatlabel)
    worksheet.write(0, 69, "H11", eformatlabel)
    worksheet.write(0, 70, "H15", eformatlabel)
    worksheet.write(0, 71, "H15", eformatlabel)
    worksheet.write(0, 72, "H13", eformatlabel)
    worksheet.write(0, 73, "H11", eformatlabel)
    worksheet.write(0, 74, "H15", eformatlabel)
    worksheet.write(0, 75, "H12", eformatlabel)
    worksheet.write(0, 76, "H12", eformatlabel)
    worksheet.write(0, 77, "H17", eformatlabel)
    worksheet.write(0, 78, "H17", eformatlabel)
    worksheet.write(0, 79, "H17", eformatlabel)
    worksheet.write(0, 80, "H17", eformatlabel)
    worksheet.write(0, 81, "H17", eformatlabel)
    worksheet.write(0, 82, "H17", eformatlabel)
    worksheet.write(0, 83, "H17", eformatlabel)
    worksheet.write(0, 84, "H17", eformatlabel)
    worksheet.write(0, 85, "*", eformatlabel)
    worksheet.write(0, 86, "H23", eformatlabel)
    worksheet.write(0, 87, "H17", eformatlabel)
    worksheet.write(0, 88, "H20", eformatlabel)
    worksheet.write(0, 89, "H20", eformatlabel)
    worksheet.write(0, 90, "H20", eformatlabel)
    worksheet.write(0, 91, "H20", eformatlabel)
    worksheet.write(0, 92, "H20", eformatlabel)
    worksheet.write(0, 93, "H20", eformatlabel)
    worksheet.write(0, 94, "H21", eformatlabel)
    worksheet.write(0, 95, "H21", eformatlabel)
    worksheet.write(0, 96, "H19", eformatlabel)
    worksheet.write(0, 97, "H19", eformatlabel)
    worksheet.write(0, 98, "H19", eformatlabel)
    worksheet.write(0, 99, "H19", eformatlabel)
    worksheet.write(0, 100, "H19", eformatlabel)
    worksheet.write(0, 101, "H23", eformatlabel)
    worksheet.write(0, 102, "H23", eformatlabel)
    worksheet.write(0, 103, "H21", eformatlabel)
    worksheet.write(0, 104, "H23", eformatlabel)
    worksheet.write(0, 105, "H19", eformatlabel)
    worksheet.write(0, 106, "H23", eformatlabel)
    worksheet.write(0, 107, "H20", eformatlabel)
    worksheet.write(0, 108, "H20", eformatlabel)
    worksheet.write(0, 109, "H23", eformatlabel)
    worksheet.write(0, 110, "H23", eformatlabel)
    worksheet.write(0, 111, "H23", eformatlabel)
    worksheet.write(0, 112, "H23", eformatlabel)
    worksheet.write(0, 113, "H1", eformatlabel)
    worksheet.write(0, 114, "H1", eformatlabel)
    worksheet.write(0, 115, "H1", eformatlabel)
    worksheet.write(0, 116, "H1", eformatlabel)
    worksheet.write(0, 117, "-", eformatlabel)
    worksheet.write(0, 118, "-", eformatlabel)
    worksheet.write(0, 119, "-", eformatlabel)
    worksheet.write(0, 120, "H7", eformatlabel)
    worksheet.write(0, 121, "-", eformatlabel)
    worksheet.write(0, 122, "H7", eformatlabel)
    worksheet.write(0, 123, "H7", eformatlabel)
    worksheet.write(0, 124, "H7", eformatlabel)
    worksheet.write(0, 125, "-", eformatlabel)
    worksheet.write(0, 126, "-", eformatlabel)
    worksheet.write(0, 127, "H7", eformatlabel)
    worksheet.write(0, 128, "H7", eformatlabel)
    worksheet.write(0, 129, "H7", eformatlabel)
    worksheet.write(0, 130, "H7", eformatlabel)
    worksheet.write(0, 131, "?", eformatlabel)
    worksheet.write(0, 132, "H7", eformatlabel)
    worksheet.write(0, 133, "H7", eformatlabel)

    worksheet.write(1, 0, "Location Code", eformatlabel)
    worksheet.write(1, 1, "Location", eformatlabel)
    worksheet.write(1, 2, "0 Book", eformatlabel)
    worksheet.write(1, 3, "1 Paperback", eformatlabel)
    worksheet.write(1, 4, "2 Large Print", eformatlabel)
    worksheet.write(1, 5, "3 Reference Book", eformatlabel)
    worksheet.write(1, 6, "4 High Demand Book", eformatlabel)
    worksheet.write(1, 7, "5 Speed Read Book", eformatlabel)
    worksheet.write(1, 8, "6 Special Collection", eformatlabel)
    worksheet.write(1, 9, "7 Reading List", eformatlabel)
    worksheet.write(1, 10, "8 Book Plus Computer Disk", eformatlabel)
    worksheet.write(1, 11, "9 Literacy", eformatlabel)
    worksheet.write(1, 12, "10 Periodical", eformatlabel)
    worksheet.write(1, 13, "11 Map", eformatlabel)
    worksheet.write(1, 14, "12 Score", eformatlabel)
    worksheet.write(1, 15, "13 Misc, Pamphlet, College Catalog", eformatlabel)
    worksheet.write(1, 16, "19 New TV Series", eformatlabel)
    worksheet.write(1, 17, "20 TV Series", eformatlabel)
    worksheet.write(1, 18, "21 Speed View", eformatlabel)
    worksheet.write(1, 19, "22 Speed View TV Series", eformatlabel)
    worksheet.write(1, 20, "23 New DVD Feature", eformatlabel)
    worksheet.write(1, 21, "25 VHS Feature", eformatlabel)
    worksheet.write(1, 22, "26 VHS Nonfiction", eformatlabel)
    worksheet.write(1, 23, "27 DVD Feature", eformatlabel)
    worksheet.write(1, 24, "28 DVD Nonfiction", eformatlabel)
    worksheet.write(1, 25, "29 VCD Feature", eformatlabel)
    worksheet.write(1, 26, "30 VCD Nonfiction", eformatlabel)
    worksheet.write(1, 27, "31 Reference CD-ROM", eformatlabel)
    worksheet.write(1, 28, "32 Circulating CD-ROM", eformatlabel)
    worksheet.write(1, 29, "33 CD Music", eformatlabel)
    worksheet.write(1, 30, "34 CD High Demand Music", eformatlabel)
    worksheet.write(1, 31, "35 Cassette Music", eformatlabel)
    worksheet.write(1, 32, "36 CD Spoken Word", eformatlabel)
    worksheet.write(1, 33, "37 CD High Demand Spoken Word", eformatlabel)
    worksheet.write(1, 34, "38 Cassette Spoken Word", eformatlabel)
    worksheet.write(1, 35, "40 LP Record", eformatlabel)
    worksheet.write(1, 36, "41 Talking Book (for the blind)", eformatlabel)
    worksheet.write(1, 37, "42 Software", eformatlabel)
    worksheet.write(1, 38, "43 Console Game", eformatlabel)
    worksheet.write(1, 39, "44 Film", eformatlabel)
    worksheet.write(1, 40, "45 Slides", eformatlabel)
    worksheet.write(1, 41, "46 Microform", eformatlabel)
    worksheet.write(1, 42, "47 Media Kit", eformatlabel)
    worksheet.write(1, 43, "48 Game", eformatlabel)
    worksheet.write(1, 44, "50 Book on Player", eformatlabel)
    worksheet.write(1, 45, "51 eReader", eformatlabel)
    worksheet.write(1, 46, "52 VideoOnPlayer", eformatlabel)
    worksheet.write(1, 47, "100 YA Book", eformatlabel)
    worksheet.write(1, 48, "101 YA Paperback", eformatlabel)
    worksheet.write(1, 49, "102 YA Large Print", eformatlabel)
    worksheet.write(1, 50, "103 YA Reference Book", eformatlabel)
    worksheet.write(1, 51, "104 YA High Demand Book", eformatlabel)
    worksheet.write(1, 52, "105 YA Reading List", eformatlabel)
    worksheet.write(1, 53, "106 YA Book Plus Computer Disk", eformatlabel)
    worksheet.write(1, 54, "107 YA Magazine", eformatlabel)
    worksheet.write(1, 55, "108 YA Miscellaneous", eformatlabel)
    worksheet.write(1, 56, "109 YA Speed Read", eformatlabel)
    worksheet.write(1, 57, "113 YA New DVD Feature", eformatlabel)
    worksheet.write(1, 58, "115 YA VHS Feature", eformatlabel)
    worksheet.write(1, 59, "116 YA VHS Nonfiction", eformatlabel)
    worksheet.write(1, 60, "117 YA DVD Feature", eformatlabel)
    worksheet.write(1, 61, "118 YA DVD Nonfiction", eformatlabel)
    worksheet.write(1, 62, "119 YA VCD Feature", eformatlabel)
    worksheet.write(1, 63, "120 YA VCD Nonfiction", eformatlabel)
    worksheet.write(1, 64, "121 YA Reference CD-ROM", eformatlabel)
    worksheet.write(1, 65, "122 YA Circulating CD-ROM", eformatlabel)
    worksheet.write(1, 66, "123 YA CD Music", eformatlabel)
    worksheet.write(1, 67, "124 YA Cassette Music", eformatlabel)
    worksheet.write(1, 68, "125 YA CD Spoken Word", eformatlabel)
    worksheet.write(1, 69, "126 YA Cassette Spoken Word", eformatlabel)
    worksheet.write(1, 70, "127 YA Media Kit", eformatlabel)
    worksheet.write(1, 71, "128 YA Game", eformatlabel)
    worksheet.write(1, 72, "129 YA Console Game", eformatlabel)
    worksheet.write(1, 73, "130 YA Book on Player", eformatlabel)
    worksheet.write(1, 74, "131 YA eReader", eformatlabel)
    worksheet.write(1, 75, "132 YA VideoOnPlayer", eformatlabel)
    worksheet.write(1, 76, "133 YA Speed View", eformatlabel)
    worksheet.write(1, 77, "150 Juvenile Book", eformatlabel)
    worksheet.write(1, 78, "151 Juvenile Paperback", eformatlabel)
    worksheet.write(1, 79, "152 Juvenile Holiday Book", eformatlabel)
    worksheet.write(1, 80, "153 Juvenile Large Print", eformatlabel)
    worksheet.write(1, 81, "154 Juvenile Reference Book", eformatlabel)
    worksheet.write(1, 82, "155 Juvenile High Demand Book", eformatlabel)
    worksheet.write(1, 83, "156 Juvenile Reading List", eformatlabel)
    worksheet.write(1, 84, "157 Juvenile Book Plus Disk", eformatlabel)
    worksheet.write(1, 85, "158 Juvenile Magazine", eformatlabel)
    worksheet.write(1, 86, "159 Juvenile Miscellaneous", eformatlabel)
    worksheet.write(1, 87, "160 Juvenile Speed Read", eformatlabel)
    worksheet.write(1, 88, "163 Juvenile TV Series", eformatlabel)
    worksheet.write(1, 89, "164 Juvenile New DVD Feature", eformatlabel)
    worksheet.write(1, 90, "165 Juvenile VHS Feature", eformatlabel)
    worksheet.write(1, 91, "166 Juvenile VHS Nonfiction", eformatlabel)
    worksheet.write(1, 92, "167 Juvenile DVD Feature", eformatlabel)
    worksheet.write(1, 93, "168 Juvenile DVD Nonfiction", eformatlabel)
    worksheet.write(1, 94, "169 Juvenile Reference CD-ROM", eformatlabel)
    worksheet.write(1, 95, "170 Juvenile Circulating CD-ROM", eformatlabel)
    worksheet.write(1, 96, "171 Juvenile CD Music", eformatlabel)
    worksheet.write(1, 97, "172 Juvenile Cassette Music", eformatlabel)
    worksheet.write(1, 98, "173 Juvenile CD Spoken Word", eformatlabel)
    worksheet.write(1, 99, "174 Juvenile Cassette Spoken Word", eformatlabel)
    worksheet.write(1, 100, "175 Juvenile Record", eformatlabel)
    worksheet.write(1, 101, "176 Juvenile Media Kit", eformatlabel)
    worksheet.write(1, 102, "177 Juvenile Game/Toy", eformatlabel)
    worksheet.write(1, 103, "178 Juvenile Console Game", eformatlabel)
    worksheet.write(1, 104, "179 Juvenile Filmstrip", eformatlabel)
    worksheet.write(1, 105, "180 Juvenile Book on Player", eformatlabel)
    worksheet.write(1, 106, "181 Juvenile eReader", eformatlabel)
    worksheet.write(1, 107, "182 Juvenile Video on Player", eformatlabel)
    worksheet.write(1, 108, "183 Juvenile Speed View", eformatlabel)
    worksheet.write(1, 109, "186 Juvenile Equipment 1", eformatlabel)
    worksheet.write(1, 110, "187 Juvenile Equipment 2", eformatlabel)
    worksheet.write(1, 111, "188 Juvenile Equipment 3", eformatlabel)
    worksheet.write(1, 112, "189 Juvenile Equipment 4", eformatlabel)
    worksheet.write(1, 113, "221 Reserve 1", eformatlabel)
    worksheet.write(1, 114, "222 Reserve 2", eformatlabel)
    worksheet.write(1, 115, "223 Reserve 3", eformatlabel)
    worksheet.write(1, 116, "224 Reserve 4", eformatlabel)
    worksheet.write(1, 117, "239 ComCat", eformatlabel)
    worksheet.write(1, 118, "241 Non-Minuteman ILL", eformatlabel)
    worksheet.write(1, 119, "242 On Order", eformatlabel)
    worksheet.write(1, 120, "243 Museum Pass", eformatlabel)
    worksheet.write(1, 121, "244 Subscribed Database", eformatlabel)
    worksheet.write(1, 122, "245 Equipment 1", eformatlabel)
    worksheet.write(1, 123, "246 High Demand Equipment", eformatlabel)
    worksheet.write(1, 124, "247 Art", eformatlabel)
    worksheet.write(1, 125, "248 E-Resource", eformatlabel)
    worksheet.write(1, 126, "249 OverDrive Advantage", eformatlabel)
    worksheet.write(1, 127, "250 In-House Equipment", eformatlabel)
    worksheet.write(1, 128, "251 Equipment 2", eformatlabel)
    worksheet.write(1, 129, "252 Equipment 3", eformatlabel)
    worksheet.write(1, 130, "253 Equipment 4", eformatlabel)
    worksheet.write(1, 131, "255 Unknown", eformatlabel)
    worksheet.write(1, 132, "256 Hotspot", eformatlabel)
    worksheet.write(1, 133, "257 High Demand Hotspot", eformatlabel)

    # Writing the report for staff to the Excel worksheet
    for rownum, row in enumerate(query_results):
        worksheet.write(rownum + 2, 0, row[0], eformat)
        worksheet.write(rownum + 2, 1, row[1], eformat)
        worksheet.write(rownum + 2, 2, row[2], eformat)
        worksheet.write(rownum + 2, 3, row[3], eformat)
        worksheet.write(rownum + 2, 4, row[4], eformat)
        worksheet.write(rownum + 2, 5, row[5], eformat)
        worksheet.write(rownum + 2, 6, row[6], eformat)
        worksheet.write(rownum + 2, 7, row[7], eformat)
        worksheet.write(rownum + 2, 8, row[8], eformat)
        worksheet.write(rownum + 2, 9, row[9], eformat)
        worksheet.write(rownum + 2, 10, row[10], eformat)
        worksheet.write(rownum + 2, 11, row[11], eformat)
        worksheet.write(rownum + 2, 12, row[12], eformat)
        worksheet.write(rownum + 2, 13, row[13], eformat)
        worksheet.write(rownum + 2, 14, row[14], eformat)
        worksheet.write(rownum + 2, 15, row[15], eformat)
        worksheet.write(rownum + 2, 16, row[16], eformat)
        worksheet.write(rownum + 2, 17, row[17], eformat)
        worksheet.write(rownum + 2, 18, row[18], eformat)
        worksheet.write(rownum + 2, 19, row[19], eformat)
        worksheet.write(rownum + 2, 20, row[20], eformat)
        worksheet.write(rownum + 2, 21, row[21], eformat)
        worksheet.write(rownum + 2, 22, row[22], eformat)
        worksheet.write(rownum + 2, 23, row[23], eformat)
        worksheet.write(rownum + 2, 24, row[24], eformat)
        worksheet.write(rownum + 2, 25, row[25], eformat)
        worksheet.write(rownum + 2, 26, row[26], eformat)
        worksheet.write(rownum + 2, 27, row[27], eformat)
        worksheet.write(rownum + 2, 28, row[28], eformat)
        worksheet.write(rownum + 2, 29, row[29], eformat)
        worksheet.write(rownum + 2, 30, row[30], eformat)
        worksheet.write(rownum + 2, 31, row[31], eformat)
        worksheet.write(rownum + 2, 32, row[32], eformat)
        worksheet.write(rownum + 2, 33, row[33], eformat)
        worksheet.write(rownum + 2, 34, row[34], eformat)
        worksheet.write(rownum + 2, 35, row[35], eformat)
        worksheet.write(rownum + 2, 36, row[36], eformat)
        worksheet.write(rownum + 2, 37, row[37], eformat)
        worksheet.write(rownum + 2, 38, row[38], eformat)
        worksheet.write(rownum + 2, 39, row[39], eformat)
        worksheet.write(rownum + 2, 40, row[40], eformat)
        worksheet.write(rownum + 2, 41, row[41], eformat)
        worksheet.write(rownum + 2, 42, row[42], eformat)
        worksheet.write(rownum + 2, 43, row[43], eformat)
        worksheet.write(rownum + 2, 44, row[44], eformat)
        worksheet.write(rownum + 2, 45, row[45], eformat)
        worksheet.write(rownum + 2, 46, row[46], eformat)
        worksheet.write(rownum + 2, 47, row[47], eformat)
        worksheet.write(rownum + 2, 48, row[48], eformat)
        worksheet.write(rownum + 2, 49, row[49], eformat)
        worksheet.write(rownum + 2, 50, row[50], eformat)
        worksheet.write(rownum + 2, 51, row[51], eformat)
        worksheet.write(rownum + 2, 52, row[52], eformat)
        worksheet.write(rownum + 2, 53, row[53], eformat)
        worksheet.write(rownum + 2, 54, row[54], eformat)
        worksheet.write(rownum + 2, 55, row[55], eformat)
        worksheet.write(rownum + 2, 56, row[56], eformat)
        worksheet.write(rownum + 2, 57, row[57], eformat)
        worksheet.write(rownum + 2, 58, row[58], eformat)
        worksheet.write(rownum + 2, 59, row[59], eformat)
        worksheet.write(rownum + 2, 60, row[60], eformat)
        worksheet.write(rownum + 2, 61, row[61], eformat)
        worksheet.write(rownum + 2, 62, row[62], eformat)
        worksheet.write(rownum + 2, 63, row[63], eformat)
        worksheet.write(rownum + 2, 64, row[64], eformat)
        worksheet.write(rownum + 2, 65, row[65], eformat)
        worksheet.write(rownum + 2, 66, row[66], eformat)
        worksheet.write(rownum + 2, 67, row[67], eformat)
        worksheet.write(rownum + 2, 68, row[68], eformat)
        worksheet.write(rownum + 2, 69, row[69], eformat)
        worksheet.write(rownum + 2, 70, row[70], eformat)
        worksheet.write(rownum + 2, 71, row[71], eformat)
        worksheet.write(rownum + 2, 72, row[72], eformat)
        worksheet.write(rownum + 2, 73, row[73], eformat)
        worksheet.write(rownum + 2, 74, row[74], eformat)
        worksheet.write(rownum + 2, 75, row[75], eformat)
        worksheet.write(rownum + 2, 76, row[76], eformat)
        worksheet.write(rownum + 2, 77, row[77], eformat)
        worksheet.write(rownum + 2, 78, row[78], eformat)
        worksheet.write(rownum + 2, 79, row[79], eformat)
        worksheet.write(rownum + 2, 80, row[80], eformat)
        worksheet.write(rownum + 2, 81, row[81], eformat)
        worksheet.write(rownum + 2, 82, row[82], eformat)
        worksheet.write(rownum + 2, 83, row[83], eformat)
        worksheet.write(rownum + 2, 84, row[84], eformat)
        worksheet.write(rownum + 2, 85, row[85], eformat)
        worksheet.write(rownum + 2, 86, row[86], eformat)
        worksheet.write(rownum + 2, 87, row[87], eformat)
        worksheet.write(rownum + 2, 88, row[88], eformat)
        worksheet.write(rownum + 2, 89, row[89], eformat)
        worksheet.write(rownum + 2, 90, row[90], eformat)
        worksheet.write(rownum + 2, 91, row[91], eformat)
        worksheet.write(rownum + 2, 92, row[92], eformat)
        worksheet.write(rownum + 2, 93, row[93], eformat)
        worksheet.write(rownum + 2, 94, row[94], eformat)
        worksheet.write(rownum + 2, 95, row[95], eformat)
        worksheet.write(rownum + 2, 96, row[96], eformat)
        worksheet.write(rownum + 2, 97, row[97], eformat)
        worksheet.write(rownum + 2, 98, row[98], eformat)
        worksheet.write(rownum + 2, 99, row[99], eformat)
        worksheet.write(rownum + 2, 100, row[100], eformat)
        worksheet.write(rownum + 2, 101, row[101], eformat)
        worksheet.write(rownum + 2, 102, row[102], eformat)
        worksheet.write(rownum + 2, 103, row[103], eformat)
        worksheet.write(rownum + 2, 104, row[104], eformat)
        worksheet.write(rownum + 2, 105, row[105], eformat)
        worksheet.write(rownum + 2, 106, row[106], eformat)
        worksheet.write(rownum + 2, 107, row[107], eformat)
        worksheet.write(rownum + 2, 108, row[108], eformat)
        worksheet.write(rownum + 2, 109, row[109], eformat)
        worksheet.write(rownum + 2, 110, row[110], eformat)
        worksheet.write(rownum + 2, 111, row[111], eformat)
        worksheet.write(rownum + 2, 112, row[112], eformat)
        worksheet.write(rownum + 2, 113, row[113], eformat)
        worksheet.write(rownum + 2, 114, row[114], eformat)
        worksheet.write(rownum + 2, 115, row[115], eformat)
        worksheet.write(rownum + 2, 116, row[116], eformat)
        worksheet.write(rownum + 2, 117, row[117], eformat)
        worksheet.write(rownum + 2, 118, row[118], eformat)
        worksheet.write(rownum + 2, 119, row[119], eformat)
        worksheet.write(rownum + 2, 120, row[120], eformat)
        worksheet.write(rownum + 2, 121, row[121], eformat)
        worksheet.write(rownum + 2, 122, row[122], eformat)
        worksheet.write(rownum + 2, 123, row[123], eformat)
        worksheet.write(rownum + 2, 124, row[124], eformat)
        worksheet.write(rownum + 2, 125, row[125], eformat)
        worksheet.write(rownum + 2, 126, row[126], eformat)
        worksheet.write(rownum + 2, 127, row[127], eformat)
        worksheet.write(rownum + 2, 128, row[128], eformat)
        worksheet.write(rownum + 2, 129, row[129], eformat)
        worksheet.write(rownum + 2, 130, row[130], eformat)
        worksheet.write(rownum + 2, 131, row[131], eformat)
        worksheet.write(rownum + 2, 132, row[132], eformat)
        worksheet.write(rownum + 2, 133, row[133], eformat)

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
    # query to count new items added at each location
    query = """\
            SELECT
              i.location_code AS "Location Code",
              l.name AS Location,
              COUNT(i.id) FILTER(WHERE i.itype_code_num = 0) AS "(H1) Book",
              COUNT(i.id) FILTER(WHERE i.itype_code_num = 1) AS "(H1) Paperback",
              COUNT(i.id) FILTER(WHERE i.itype_code_num = 2) AS "(H1) Large Print",
              COUNT(i.id) FILTER(WHERE i.itype_code_num = 3) AS "(H1) Reference",
              COUNT(i.id) FILTER(WHERE i.itype_code_num = 4) AS "(H1) High Demand Book",
              COUNT(i.id) FILTER(WHERE i.itype_code_num = 5) AS "(H1) Speed Read Book",
              COUNT(i.id) FILTER(WHERE i.itype_code_num = 6) AS "(H1) Rental Book",
              COUNT(i.id) FILTER(WHERE i.itype_code_num = 7) AS "(H1) Reading List",
              COUNT(i.id) FILTER(WHERE i.itype_code_num = 8) AS "(H1) Book Plus Computer Disk",
              COUNT(i.id) FILTER(WHERE i.itype_code_num = 9) AS "(H1) Literacy",
              COUNT(i.id) FILTER(WHERE i.itype_code_num = 10) AS "(*) Periodical",
              COUNT(i.id) FILTER(WHERE i.itype_code_num = 11) AS "(H10) Map",
              COUNT(i.id) FILTER(WHERE i.itype_code_num = 12) AS "(H1) Score",
              COUNT(i.id) FILTER(WHERE i.itype_code_num = 13) AS "(H10) Misc, Pamplet, College Catalog",
              COUNT(i.id) FILTER(WHERE i.itype_code_num = 19) AS "(H4) New TV Series",
              COUNT(i.id) FILTER(WHERE i.itype_code_num = 20) AS "(H4) TV Series",
              COUNT(i.id) FILTER(WHERE i.itype_code_num = 21) AS "(H4) Speed View",
              COUNT(i.id) FILTER(WHERE i.itype_code_num = 22) AS "(H4) Speed View TV Series",
              COUNT(i.id) FILTER(WHERE i.itype_code_num = 23) AS "(H4) New DVD Feature",
              --COUNT(i.id) FILTER(WHERE i.itype_code_num = 24) AS "(H4) New VCD Feature",
              COUNT(i.id) FILTER(WHERE i.itype_code_num = 25) AS "(H4) VHS Feature",
              COUNT(i.id) FILTER(WHERE i.itype_code_num = 26) AS "(H4) VHS Nonfiction",
              COUNT(i.id) FILTER(WHERE i.itype_code_num = 27) AS "(H4) DVD Feature",
              COUNT(i.id) FILTER(WHERE i.itype_code_num = 28) AS "(H4) DVD Nonfiction",
              COUNT(i.id) FILTER(WHERE i.itype_code_num = 29) AS "(H4) VCD Feature",
              COUNT(i.id) FILTER(WHERE i.itype_code_num = 30) AS "(H4) VCD Nonfiction",
              COUNT(i.id) FILTER(WHERE i.itype_code_num = 31) AS "(H8) Reference CD-ROM",
              COUNT(i.id) FILTER(WHERE i.itype_code_num = 32) AS "(H8) Circulating CD-ROM",
              COUNT(i.id) FILTER(WHERE i.itype_code_num = 33) AS "(H3) CD Music",
              COUNT(i.id) FILTER(WHERE i.itype_code_num = 34) AS "(H3) CD High Demand Music",
              COUNT(i.id) FILTER(WHERE i.itype_code_num = 35) AS "(H3) Casette Music",
              COUNT(i.id) FILTER(WHERE i.itype_code_num = 36) AS "(H3) CD Spoken Word",
              COUNT(i.id) FILTER(WHERE i.itype_code_num = 37) AS "(H3) CD High Demand Spoken Word",
              COUNT(i.id) FILTER(WHERE i.itype_code_num = 38) AS "(H3) Cassette Spoken Word",
              COUNT(i.id) FILTER(WHERE i.itype_code_num = 40) AS "(H3) LP Record",
              COUNT(i.id) FILTER(WHERE i.itype_code_num = 41) AS "(H3) Talking Book (for the blind)",
              COUNT(i.id) FILTER(WHERE i.itype_code_num = 42) AS "(H8) Software",
              COUNT(i.id) FILTER(WHERE i.itype_code_num = 43) AS "(H8) Console Game",
              COUNT(i.id) FILTER(WHERE i.itype_code_num = 44) AS "(H10) Film",
              COUNT(i.id) FILTER(WHERE i.itype_code_num = 45) AS "(H10) Slides",
              COUNT(i.id) FILTER(WHERE i.itype_code_num = 46) AS "(H9) Microfilm",
              COUNT(i.id) FILTER(WHERE i.itype_code_num = 47) AS "(H10) Media Kit",
              COUNT(i.id) FILTER(WHERE i.itype_code_num = 48) AS "(H10) Game",
              COUNT(i.id) FILTER(WHERE i.itype_code_num = 50) AS "(H3) Book on Player",
              COUNT(i.id) FILTER(WHERE i.itype_code_num = 51) AS "(H10) eReader",
              COUNT(i.id) FILTER(WHERE i.itype_code_num = 52) AS "(H4) VideoOnPlayer",
              COUNT(i.id) FILTER(WHERE i.itype_code_num = 100) AS "(H12) YA Book",
              COUNT(i.id) FILTER(WHERE i.itype_code_num = 101) AS "(H12) YA Paperback",
              COUNT(i.id) FILTER(WHERE i.itype_code_num = 102) AS "(H12) YA Large Print",
              COUNT(i.id) FILTER(WHERE i.itype_code_num = 103) AS "(H12) YA Reference Book",
              COUNT(i.id) FILTER(WHERE i.itype_code_num = 104) AS "(H12) YA High Demand Book",
              COUNT(i.id) FILTER(WHERE i.itype_code_num = 105) AS "(H12) YA Reading List",
              COUNT(i.id) FILTER(WHERE i.itype_code_num = 106) AS "(H12) YA YA Book Plus Computer Disk",
              COUNT(i.id) FILTER(WHERE i.itype_code_num = 107) AS "(*) YA Magazine",
              COUNT(i.id) FILTER(WHERE i.itype_code_num = 108) AS "(H21) YA Miscellaneous",
              COUNT(i.id) FILTER(WHERE i.itype_code_num = 109) AS "(H12) YA Speed Read",
              COUNT(i.id) FILTER(WHERE i.itype_code_num = 113) AS "() YA 113",
              COUNT(i.id) FILTER(WHERE i.itype_code_num = 115) AS "() YA 115",
              COUNT(i.id) FILTER(WHERE i.itype_code_num = 116) AS "() YA 116",
              COUNT(i.id) FILTER(WHERE i.itype_code_num = 117) AS "() YA 117",
              COUNT(i.id) FILTER(WHERE i.itype_code_num = 118) AS "() YA 118",
              COUNT(i.id) FILTER(WHERE i.itype_code_num = 119) AS "() YA 119",
              COUNT(i.id) FILTER(WHERE i.itype_code_num = 120) AS "() YA 120",
              COUNT(i.id) FILTER(WHERE i.itype_code_num = 121) AS "() YA 121",
              COUNT(i.id) FILTER(WHERE i.itype_code_num = 122) AS "() YA 122",
              COUNT(i.id) FILTER(WHERE i.itype_code_num = 123) AS "() YA 123",
              COUNT(i.id) FILTER(WHERE i.itype_code_num = 124) AS "() YA 124",
              COUNT(i.id) FILTER(WHERE i.itype_code_num = 125) AS "() YA 125",
              COUNT(i.id) FILTER(WHERE i.itype_code_num = 126) AS "() YA 126",
              COUNT(i.id) FILTER(WHERE i.itype_code_num = 127) AS "() YA 127",
              COUNT(i.id) FILTER(WHERE i.itype_code_num = 128) AS "() YA 128",
              COUNT(i.id) FILTER(WHERE i.itype_code_num = 129) AS "() YA 129",
              COUNT(i.id) FILTER(WHERE i.itype_code_num = 130) AS "() YA 130",
              COUNT(i.id) FILTER(WHERE i.itype_code_num = 131) AS "() YA 131",
              COUNT(i.id) FILTER(WHERE i.itype_code_num = 132) AS "() YA 132",
              COUNT(i.id) FILTER(WHERE i.itype_code_num = 133) AS "() YA 133",
              COUNT(i.id) FILTER(WHERE i.itype_code_num = 150) AS "() 150",
              COUNT(i.id) FILTER(WHERE i.itype_code_num = 151) AS "() 151",
              COUNT(i.id) FILTER(WHERE i.itype_code_num = 152) AS "() 152",
              COUNT(i.id) FILTER(WHERE i.itype_code_num = 153) AS "() 153",
              COUNT(i.id) FILTER(WHERE i.itype_code_num = 154) AS "() 154",
              COUNT(i.id) FILTER(WHERE i.itype_code_num = 155) AS "() 155",
              COUNT(i.id) FILTER(WHERE i.itype_code_num = 156) AS "() 156",
              COUNT(i.id) FILTER(WHERE i.itype_code_num = 157) AS "() 157",
              COUNT(i.id) FILTER(WHERE i.itype_code_num = 158) AS "() 158",
              COUNT(i.id) FILTER(WHERE i.itype_code_num = 159) AS "() 159",
              COUNT(i.id) FILTER(WHERE i.itype_code_num = 160) AS "() 160",
              COUNT(i.id) FILTER(WHERE i.itype_code_num = 163) AS "() 163",
              COUNT(i.id) FILTER(WHERE i.itype_code_num = 164) AS "() 164",
              COUNT(i.id) FILTER(WHERE i.itype_code_num = 165) AS "() 165",
              COUNT(i.id) FILTER(WHERE i.itype_code_num = 166) AS "() 166",
              COUNT(i.id) FILTER(WHERE i.itype_code_num = 167) AS "() 167",
              COUNT(i.id) FILTER(WHERE i.itype_code_num = 168) AS "() 168",
              COUNT(i.id) FILTER(WHERE i.itype_code_num = 169) AS "() 169",
              COUNT(i.id) FILTER(WHERE i.itype_code_num = 170) AS "() 170",
              COUNT(i.id) FILTER(WHERE i.itype_code_num = 171) AS "() 171",
              COUNT(i.id) FILTER(WHERE i.itype_code_num = 172) AS "() 172",
              COUNT(i.id) FILTER(WHERE i.itype_code_num = 173) AS "() 173",
              COUNT(i.id) FILTER(WHERE i.itype_code_num = 174) AS "() 174",
              COUNT(i.id) FILTER(WHERE i.itype_code_num = 175) AS "() 175",
              COUNT(i.id) FILTER(WHERE i.itype_code_num = 176) AS "() 176",
              COUNT(i.id) FILTER(WHERE i.itype_code_num = 177) AS "() 177",
              COUNT(i.id) FILTER(WHERE i.itype_code_num = 178) AS "() 178",
              COUNT(i.id) FILTER(WHERE i.itype_code_num = 179) AS "() 179",
              COUNT(i.id) FILTER(WHERE i.itype_code_num = 180) AS "() 180",
              COUNT(i.id) FILTER(WHERE i.itype_code_num = 181) AS "() 181",
              COUNT(i.id) FILTER(WHERE i.itype_code_num = 182) AS "() 182",
              COUNT(i.id) FILTER(WHERE i.itype_code_num = 183) AS "() 183",
              COUNT(i.id) FILTER(WHERE i.itype_code_num = 186) AS "() 186",
              COUNT(i.id) FILTER(WHERE i.itype_code_num = 187) AS "() 187",
              COUNT(i.id) FILTER(WHERE i.itype_code_num = 188) AS "() 188",
              COUNT(i.id) FILTER(WHERE i.itype_code_num = 189) AS "() 189",
              COUNT(i.id) FILTER(WHERE i.itype_code_num = 221) AS "() 221",
              COUNT(i.id) FILTER(WHERE i.itype_code_num = 222) AS "() 222",
              COUNT(i.id) FILTER(WHERE i.itype_code_num = 223) AS "() 223",
              COUNT(i.id) FILTER(WHERE i.itype_code_num = 224) AS "() 224",
              COUNT(i.id) FILTER(WHERE i.itype_code_num = 239) AS "() 239",
              COUNT(i.id) FILTER(WHERE i.itype_code_num = 241) AS "() 241",
              COUNT(i.id) FILTER(WHERE i.itype_code_num = 242) AS "() 242",
              COUNT(i.id) FILTER(WHERE i.itype_code_num = 243) AS "() 243",
              COUNT(i.id) FILTER(WHERE i.itype_code_num = 244) AS "() 244",
              COUNT(i.id) FILTER(WHERE i.itype_code_num = 245) AS "() 245",
              COUNT(i.id) FILTER(WHERE i.itype_code_num = 246) AS "() 246",
              COUNT(i.id) FILTER(WHERE i.itype_code_num = 247) AS "() 247",
              COUNT(i.id) FILTER(WHERE i.itype_code_num = 248) AS "() 248",
              COUNT(i.id) FILTER(WHERE i.itype_code_num = 249) AS "() 249",
              COUNT(i.id) FILTER(WHERE i.itype_code_num = 250) AS "() 250",
              COUNT(i.id) FILTER(WHERE i.itype_code_num = 251) AS "() 251",
              COUNT(i.id) FILTER(WHERE i.itype_code_num = 252) AS "() 252",
              COUNT(i.id) FILTER(WHERE i.itype_code_num = 253) AS "() 253",
              COUNT(i.id) FILTER(WHERE i.itype_code_num = 255) AS "() 255",
              COUNT(i.id) FILTER(WHERE i.itype_code_num = 256) AS "() 256",
              COUNT(i.id) FILTER(WHERE i.itype_code_num = 257) AS "() 257"
            FROM sierra_view.item_record i
            JOIN sierra_view.location_myuser l
              ON i.location_code = l.code

            GROUP BY 1,2
            ORDER BY 1
            """
    query_results = run_query(query)

    # generate excel file from those query results
    excel_file = "/Scripts/Annual Reports/Archive/holdings_profile{}.xlsx".format(
        date.today()
    )
    excel_writer(query_results, excel_file)

    # send email with attached file
    email_subject = "Holdings Profile"
    email_message = """***Holdings Profile***


The Holdings Profile report has been attached."""
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
        email_subject = "annual reports: holdings profile script error"
        email_message = (
            "Your script failed with the following error:\n\n" + traceback.format_exc()
        )

        send_email_error(email_subject, email_message, emailto)
        raise
