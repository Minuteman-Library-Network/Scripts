#!/usr/bin/env python3

# Run in maps

"""
Jeremy Goldstein
Minuteman Library Network

Generates html file containing an interactive choropleth maps with
the option to select from a number of metrics to display

The completed maps are then uploaded to the reports folder we maintain for each library within Minuteman
"""

import geopandas as gpd
import plotly.express as px
import plotly.io as pio
import pandas as pd
import numpy as np
import plotly.graph_objs as go
import json
import psycopg2
import configparser
import os
import pysftp
import sys
import time
from datetime import date


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


# Produce choropleth map as html file
def gen_map(patron_df):

    # generate pandas dataframe from provided tigerline GIS file
    zipfile = "zip://Data Sources//tl_2024_25_bg.zip"
    df = gpd.read_file(zipfile).to_crs("EPSG:4326")
    df.columns = df.columns.str.lower()

    # merge tigerline data with sierra patron data on geoid values
    df = df.merge(patron_df, on="geoid", how="inner")

    # generate df for census acs population totals
    pop_df = pd.read_csv(
        "/Scripts/Patron Maps/Data Sources/2024 acs pop estimate bg.csv",
        dtype={"geoid": str},
    )

    # merge population total dataframe with combined sierra & tigerline dataframe
    df = df.merge(pop_df, on="geoid", how="inner")

    # calculate percentage of population with a library card
    df["pct_cardholders"] = df.total_patrons / df.estimated_population * 100.00
    df["pct_cardholders"] = df["pct_cardholders"].round(decimals=2)

    # Convert dataframe to json for use with plotly
    zipjson = json.loads(df.to_json())

    cols_dd = [
        "total_patrons",
        "estimated_population",
        "pct_cardholders",
        "total_checkouts",
        "checkouts_per_patron",
        "total_new_patrons",
        "total_active_patrons",
        "pct_active",
        "total_blocked_patrons",
        "pct_blocked",
    ]

    # generate choropleth map with metric selector
    visible = np.array(cols_dd)

    traces = []
    buttons = []

    for value in cols_dd:
        traces.append(
            go.Choropleth(
                locations=df.index,
                geojson=zipjson,
                z=df[value],
                colorbar_title=value,
                colorscale="YlGnBu",
                hovertemplate="<b>"
                + df.geographic_area_name
                + "</b><br>"
                + "Total: "
                + df[value].astype(str)
                + "<extra></extra>",
                visible=True if value == cols_dd[0] else False,
            )
        )

        buttons.append(
            dict(
                label=value,
                method="update",
                args=[
                    {"visible": list(visible == value)},
                    {"title": f"<b>{value}</b>"},
                ],
            )
        )

    updatemenus = [{"active": 0, "buttons": buttons}]

    fig = go.Figure(data=traces, layout=dict(updatemenus=updatemenus))

    first_title = cols_dd[0]
    fig.update_layout(title=f"<b>{first_title}</b>", title_x=0.5, geo=dict(scope="usa"))

    fig.update_geos(fitbounds="locations", visible=True)

    # save map to html file
    pio.write_html(
        fig,
        file="/Scripts/Patron Maps/Temp Files/AllInOneMap{}.html".format(date.today()),
        auto_open=False,
    )


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

    srv.cwd("/reports/Library-Specific Reports/" + library + "/Patron Maps/")

    # remove old file
    for fname in srv.listdir_attr():
        fullpath = (
            "/reports/Library-Specific Reports/"
            + library
            + "/Patron Maps/{}".format(fname.filename)
        )
        name = str(fname.filename)
        if name.startswith("AllInOne") and (
            (time.time() - fname.st_mtime) // (24 * 3600) >= 90
        ):
            srv.remove(fullpath)

    # upload file
    srv.put(local_file)
    # close sftp connection
    srv.close()
    # remove local copy of file
    os.remove(local_file)


def main(library, tracts):
    query = (
        """
            /*Gather usage stats grouped by census block group for the given set of census tracts*/
            SELECT
              CASE 
	            WHEN v.field_content IS NULL THEN 'no data' 
	            WHEN v.field_content = '' THEN v.field_content 
	            ELSE SUBSTRING(REGEXP_REPLACE(v.field_content,'\|(s|c|t|b)','','g'),1,12) 
              END AS geoid, 
              COUNT(DISTINCT p.id) AS total_patrons,
              SUM(p.checkout_total) AS total_checkouts,
    		  COUNT(DISTINCT p.id) FILTER(WHERE rm.creation_date_gmt::DATE >= CURRENT_DATE - INTERVAL '1 year') AS total_new_patrons,
    		  COUNT(DISTINCT p.id) FILTER(WHERE p.activity_gmt::DATE >= CURRENT_DATE - INTERVAL '1 year' AND NOT ((p.mblock_code != '-') OR (p.owed_amt >= 100))) AS total_active_patrons,
              ROUND(100.0 * (CAST(COUNT(DISTINCT p.id) FILTER(WHERE p.activity_gmt::DATE >= CURRENT_DATE - INTERVAL '1 year' AND NOT ((p.mblock_code != '-') OR (p.owed_amt >= 100))) AS NUMERIC (12,2))) / CAST(COUNT(DISTINCT p.id) AS NUMERIC (12,2)), 2)::VARCHAR AS pct_active,
              COUNT(DISTINCT p.id) FILTER(WHERE ((p.mblock_code != '-') OR (p.owed_amt >= 100))) as total_blocked_patrons,
              ROUND(100.0 * (CAST(COUNT(DISTINCT p.id) FILTER(WHERE ((p.mblock_code != '-') OR (p.owed_amt >= 100))) as numeric (12,2)) / cast(COUNT(DISTINCT p.id) as numeric (12,2))),2)::VARCHAR AS pct_blocked,
              ROUND((100.0 * SUM(p.checkout_total))/(100.0 *COUNT(DISTINCT p.id)),2)::VARCHAR AS checkouts_per_patron
            
            FROM sierra_view.patron_record p
            JOIN sierra_view.record_metadata rm 
              ON p.id = rm.id 
            LEFT JOIN sierra_view.hold h 
              ON p.id = h.patron_record_id 
            LEFT JOIN sierra_view.varfield v 
              ON v.record_id = p.id
              AND v.varfield_type_code = 'k'
              AND v.field_content ~ '^\|s25' 
            WHERE SUBSTRING(REGEXP_REPLACE(v.field_content,'\|(s|c|t|b)','','g'),6,6) IN ("""
        + tracts
        + """\
            ) 
            GROUP BY 1
            ORDER BY 2 DESC
            """
    )

    query_results = run_query(query)

    # convert query results into a pandas dataframe
    column_names = [
        "geoid",
        "total_patrons",
        "total_checkouts",
        "total_new_patrons",
        "total_active_patrons",
        "pct_active",
        "total_blocked_patrons",
        "pct_blocked",
        "checkouts_per_patron",
    ]
    df = pd.DataFrame(query_results, columns=column_names)

    # generate map based on dataframe
    gen_map(df)

    # upload file and delete local copy
    sftp_file("C:\\Scripts\\Patron Maps\\Temp Files\\AllInOneMap{}.html".format(date.today()),library)


main("Acton", "'363102','363103','363104','363105','363106','363201','363202'")
main(
    "Arlington",
    "'356100','356200','356300','356400','356500','356601','356602','356701','356702','356703','356704'",
)
main("Ashland", "'385100','385101','385102','385201','385202','385203','385204'")
main("Bedford", "'359100','359300','359301','359302','359303'")
main(
    "Belmont", "'357100','357200','357300','357400','357500','357600','357700','357800'"
)
main(
    "Brookline",
    "'400100','400200','400201','400202','400300','400400','400401','400402','400500','400600','400700','400800','400900','401000','401100','401200','401201','401202'",
)
main(
    "Cambridge",
    "'352101','352102','352200','352300','352400','352500','352600','352700','352800','352900','353000','353101','353102','353200','353300','353400','353500','353600','353700','353800','353900','354000','354100','354200','354300','354400','354500','354600','354601','354602','354700','354800','354900','354901','354902','355000','359400'",
)
main("Concord", "'361100','361200','361300','359301'")
main("Dedham", "'402101','402102','402200','402300','402400','402500'")
main("Dover", "'405100'")
main(
    "Framingham Public",
    "'383101','383102','383200','383300','383400','383401','383402','383501','383502','383600','383700','383800','383900','383901','383902','383903','383904','384000','384001','384002','384003','384004'",
)
main(
    "Franklin",
    "'442101','442102','442103','442104','442105','442201','442202','442203','442204'",
)
main("Holliston", "'387100','387201','387202'")
main("Lexington", "'358100','358200','358300','358400','358500','358600','358700'")
main("Lincoln", "'360100','360200','360300','359302'")
main("Maynard", "'364101','364102'")
main("Medfield", "'406101','406102'")
main(
    "Medford",
    "'339100','339101','339102','339200','339300','339400','339500','339600','339700','339801','339802','339803','339804','339900','340000','340100'",
)
main("Medway", "'408101','408102','408103','408104'")
main("Millis", "'407100','407101','407102'")
main("Natick", "'382100','382200','382300','382400','382500','382601','382602'")
main("Needham", "'403100','403300','403400','403500','403501','403502','457200'")
main(
    "Newton",
    "'373100','373200','373300','373400','373500','373600','373700','373800','373900','373901','373902','374000','374100','374200','374300','374400','374500','374600','374700','374800'",
)
main(
    "Norwood", "'413100','413200','413201','413202','413300','413401','413402','413500'"
)
main("Sherborn", "'386100'")
main(
    "Somerville",
    "'350103','350104','350105','350106','350107','350108','350109','350200','350201','350202','350300','350400','350500','350600','350700','350701','350702','350800','350900','351000','351001','351002','351100','351101','351102','351203','351204','351300','351403','351404','351500'",
)
main("Stow", "'323100','323101','323102','980000'")
main("Sudbury", "'365100','365201','365202'")
main("Wayland", "'366100','366201','366202'")
main(
    "Waltham",
    "'368101','368102','368200','368300','368400','368500','368600','368700','368800','368901','368902','369000','369100'",
)
main(
    "Watertown",
    "'370101','370102','370103','370104','370201','370202','370300','370301','370302','370400','370401','370402','370403'",
)
main("Wellesley", "'404100','404201','404202','404301','404302','404400'")
main("Weston", "'367100','367200'")
main("Westwood", "'412100','412200','412300'")
main("Winchester", "'338100','338200','338300','338400','338500'")
main(
    "Woburn",
    "'333100','333200','333300','333400','333501','333502','333600','333601','333602'",
)