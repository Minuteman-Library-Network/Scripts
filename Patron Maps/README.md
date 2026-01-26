# Patron Maps

Python scripts to produce Choropleth maps based around data in our Sierra patron records and tables available from the U.S. Census Bureau.

Key to this script is the presence of a census field that we have added to our patron records using the Cenus Bureau Geocoder (https://geocoding.geo.census.gov/geocoder)
For records where an address has been successfully matched to the geocoder the census block number is added to the patron record.  This allows for our SQL scripts to group statistics at the state, county, tract or block group level for these maps.  The script used for that process can be found here ([[https://github.com/Minuteman-Library-Network/Patron-Geocoder](https://github.com/Minuteman-Library-Network/Scripts/tree/main/Geocoder)](https://github.com/Minuteman-Library-Network/Scripts/tree/main/Geocoder)).

Here is an example of the output from each script

Pct_Cardholders_single_map_with_baselayer.py:
![Framingham_Cardholder_Pct.png](https://github.com/Minuteman-Library-Network/Patron-Maps/blob/main/img/Framingham_Cardholder_Pct.png)

Pct_Active_duel_maps_with_baselayer.py:
![Framingham_Active_Pct.png](https://github.com/Minuteman-Library-Network/Patron-Maps/blob/main/img/Framingham_Active_Pct.png)

All_In_One_Map.py:
![All_In_One_Map.png](https://github.com/Minuteman-Library-Network/Patron-Maps/blob/main/img/All_In_One_Map.png)

Created and maintained by Jeremy Goldstein for test purposes only. Use at your own risk. Not supported by the Minuteman Library Network. 
