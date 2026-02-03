#!/usr/bin/env python3

# Run in maps

"""
Jeremy Goldstein
Minuteman Library Network

Generates html file containing two interactive choropleth maps showing
the number of active patrons and cardholders in each census block group with a town

The completed maps are then uploaded to the reports folder we maintain for each library within Minuteman
"""

import json
import pandas as pd
import geopandas as gpd
import plotly.io as pio
import psycopg2
import configparser
import plotly.graph_objects as go
from plotly.subplots import make_subplots
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
def gen_map(patron_df, lat, lon, mapzoom):

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

    # generate choropleth mapbox object
    fig1 = go.Choroplethmapbox(
        geojson=zipjson,
        locations=df.geoid,
        featureidkey="properties.geoid",
        z=df.pct_cardholders,
        colorscale="YlGnBu",
        hovertemplate="<b>"
        + df.geographic_area_name
        + "</b><br>"
        + df.pct_cardholders.astype(str)
        + "%"
        + "</b><br>"
        + "Patron Total: "
        + df.total_patrons.astype(str)
        + "</b><br>"
        + "Est. Pop: "
        + df.estimated_population.astype(str)
        + "<extra></extra>",
        # zmin=0, zmax=1,
        marker_opacity=0.65,
        marker_line_width=1,
        showlegend=False,
        showscale=False,
    )
    fig2 = go.Choroplethmapbox(
        geojson=zipjson,
        locations=df.geoid,
        featureidkey="properties.geoid",
        z=df.pct_active,
        colorscale="matter",
        hovertemplate="<b>"
        + df.geographic_area_name
        + "</b><br>"
        + df.pct_active.astype(str)
        + "%"
        + "</b><br>"
        + "Patron Total: "
        + df.total_patrons.astype(str)
        + "</b><br>"
        + "Active Patron Total: "
        + df.total_active_patrons.astype(str)
        + "</b><br>"
        + "<extra></extra>",
        # zmin=0, zmax=8000,
        marker_opacity=0.65,
        marker_line_width=1,
        showlegend=False,
        showscale=False,
    )

    fig = make_subplots(
        rows=1,
        cols=2,
        subplot_titles=("Cardholder Percentage", "Active Percentage"),
        specs=[[{"type": "choroplethmapbox"}, {"type": "choroplethmapbox"}]],
    )

    # Add first map
    fig.add_trace(fig1, row=1, col=1)

    # Add second map
    fig.add_trace(fig2, row=1, col=2)

    fig.update_layout(
        mapbox_style="open-street-map",
        mapbox2_style="open-street-map",
        mapbox_zoom=mapzoom,
        mapbox2_zoom=mapzoom,
        mapbox_center={"lat": lat, "lon": lon},
        mapbox2_center={"lat": lat, "lon": lon},
    )
    # save map to html file
    pio.write_html(
        fig,
        file="/Scripts/Patron Maps/Temp Files/ActivePatrons{}.html".format(
            date.today()
        ),
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
        if name.startswith("ActivePatrons") and (
            (time.time() - fname.st_mtime) // (24 * 3600) >= 90
        ):
            srv.remove(fullpath)

    # upload file
    srv.put(local_file)
    # close sftp connection
    srv.close()
    # remove local copy of file
    os.remove(local_file)


def main(library, tracts, lat, lon, mapzoom):
    query = (
        """
            /*Gather patron counts grouped by census block group for the given set of census tracts*/
            SELECT
              CASE 
	            WHEN v.field_content IS NULL THEN 'no data' 
	            WHEN v.field_content = '' THEN v.field_content 
	            ELSE SUBSTRING(REGEXP_REPLACE(v.field_content,'\|(s|c|t|b)','','g'),1,12) 
              END AS geoid, 
              COUNT(DISTINCT p.id) AS total_patrons,
    		  COUNT(DISTINCT p.id) FILTER(WHERE p.activity_gmt::DATE >= CURRENT_DATE - INTERVAL '1 year' AND NOT ((p.mblock_code != '-') OR (p.owed_amt >= 100))) AS total_active_patrons,
              ROUND(100.0 * (CAST(COUNT(DISTINCT p.id) FILTER(WHERE p.activity_gmt::DATE >= CURRENT_DATE - INTERVAL '1 year' AND NOT ((p.mblock_code != '-') OR (p.owed_amt >= 100))) AS NUMERIC (12,2))) / CAST(COUNT(DISTINCT p.id) AS NUMERIC (12,2)), 2)::VARCHAR AS pct_active
            
            FROM sierra_view.patron_record p 
            JOIN sierra_view.record_metadata rm 
              ON p.id = rm.id 
            LEFT JOIN sierra_view.varfield v 
              ON v.record_id = p.id
              AND v.varfield_type_code = 'k'
              AND v.field_content ~ '^\|s25' 
            
            WHERE SUBSTRING(REGEXP_REPLACE(v.field_content,'\|(s|c|t|b)','','g'),6,6) IN ("""
        + tracts
        + """) 
            
            GROUP BY 1
            ORDER BY 2 DESC
            """
    )

    query_results = run_query(query)

    # convert query results into a pandas dataframe
    column_names = [
        "geoid",
        "total_patrons",
        "total_active_patrons",
        "pct_active",
    ]
    df = pd.DataFrame(query_results, columns=column_names)

    # generate map based on dataframe and coordinates provided in order to center and zoom it appropriately
    gen_map(df, lat, lon, mapzoom)

    # upload file and delete local copy
    sftp_file("C:\\Scripts\\Patron Maps\\Temp Files\\ActivePatrons{}.html".format(date.today()),library)


# run for each municipality within Minuteman
main(
    "Acton",
    "'363102','363103','363104','363105','363106','363201','363202'",
    42.4831,
    -71.4451,
    11,
)
main(
    "Arlington",
    "'356100','356200','356300','356400','356500','356601','356602','356701','356702','356703','356704'",
    42.4168,
    -71.1674,
    12,
)
main(
    "Ashland",
    "'385100','385101','385102','385201','385202','385203','385204'",
    42.2597,
    -71.4705,
    12,
)
main("Bedford", "'359100','359300','359301','359302','359303'", 42.4932, -71.2766, 11)
main(
    "Belmont",
    "'357100','357200','357300','357400','357500','357600','357700','357800'",
    42.3952,
    -71.1821,
    12,
)
main(
    "Brookline",
    "'400100','400200','400201','400202','400300','400400','400401','400402','400500','400600','400700','400800','400900','401000','401100','401200','401201','401202'",
    42.3232,
    -71.1422,
    12,
)
main(
    "Cambridge",
    "'352101','352102','352200','352300','352400','352500','352600','352700','352800','352900','353000','353101','353102','353200','353300','353400','353500','353600','353700','353800','353900','354000','354100','354200','354300','354400','354500','354600','354601','354602','354700','354800','354900','354901','354902','355000','359400'",
    42.3736,
    -71.1097,
    12,
)
main("Concord", "'361100','361200','361300','359301'", 42.4586, -71.3598, 11)
main(
    "Dedham",
    "'402101','402102','402200','402300','402400','402500'",
    42.2448,
    -71.1812,
    11,
)
main("Dover", "'405100'", 42.2417, -71.2875, 11)
main(
    "Framingham Public",
    "'383101','383102','383200','383300','383400','383401','383402','383501','383502','383600','383700','383800','383900','383901','383902','383903','383904','384000','384001','384002','384003','384004'",
    42.3039,
    -71.4233,
    11,
)
main(
    "Franklin",
    "'442101','442102','442103','442104','442105','442201','442202','442203','442204'",
    42.0845,
    -71.4025,
    11,
)
main("Holliston", "'387100','387201','387202'", 42.1974, -71.4413, 11)
main(
    "Lexington",
    "'358100','358200','358300','358400','358500','358600','358700'",
    42.4482,
    -71.2253,
    11,
)
main("Lincoln", "'360100','360200','360300','359302'", 42.4293, -71.3162, 11)
main("Maynard", "'364101','364102'", 42.4265, -71.4543, 12)
main("Medfield", "'406101','406102'", 42.1821, -71.3101, 11)
main(
    "Medford",
    "'339100','339101','339102','339200','339300','339400','339500','339600','339700','339801','339802','339803','339804','339900','340000','340100'",
    42.4245,
    -71.1106,
    12,
)
main("Medway", "'408101','408102','408103','408104'", 42.1554, -71.4271, 12)
main("Millis", "'407100','407101','407102'", 42.1663, -71.3613, 11)
main(
    "Natick",
    "'382100','382200','382300','382400','382500','382601','382602'",
    42.2901,
    -71.3527,
    11,
)
main(
    "Needham",
    "'403100','403300','403400','403500','403501','403502','457200'",
    42.2857,
    -71.2448,
    11,
)
main(
    "Newton",
    "'373100','373200','373300','373400','373500','373600','373700','373800','373900','373901','373902','374000','374100','374200','374300','374400','374500','374600','374700','374800'",
    42.3253,
    -71.2139,
    11,
)
main(
    "Norwood",
    "'413100','413200','413201','413202','413300','413401','413402','413500'",
    42.1819,
    -71.1966,
    11,
)
main("Sherborn", "'386100'", 42.2321, -71.3747, 11)
main(
    "Somerville",
    "'350103','350104','350105','350106','350107','350108','350109','350200','350201','350202','350300','350400','350500','350600','350700','350701','350702','350800','350900','351000','351001','351002','351100','351101','351102','351203','351204','351300','351403','351404','351500'",
    42.3953,
    -71.1037,
    11,
)
main("Stow", "'323100','323101','323102','980000'", 42.4283, -71.5117, 11)
main("Sudbury", "'365100','365201','365202'", 42.3890, -71.4225, 11)
main("Wayland", "'366100','366201','366202'", 42.3613, -71.3634, 11)
main(
    "Waltham",
    "'368101','368102','368200','368300','368400','368500','368600','368700','368800','368901','368902','369000','369100'",
    42.3890,
    -71.2401,
    11,
)
main(
    "Watertown",
    "'370101','370102','370103','370104','370201','370202','370300','370301','370302','370400','370401','370402','370403'",
    42.3723,
    -71.1785,
    13,
)
main(
    "Wellesley",
    "'404100','404201','404202','404301','404302','404400'",
    42.2989,
    -71.2786,
    11,
)
main("Weston", "'367100','367200'", 42.3580, -71.2958, 11)
main("Westwood", "'412100','412200','412300'", 42.2210, -71.1985, 11)
main(
    "Winchester", "'338100','338200','338300','338400','338500'", 42.4547, -71.1496, 12
)
main(
    "Woburn",
    "'333100','333200','333300','333400','333501','333502','333600','333601','333602'",
    42.4899,
    -71.1595,
    11,
)