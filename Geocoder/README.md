This script will gather 10,000 patron records at a time via a SQL query and pass the addresses of those records through the US Census Bureau's geocoding service so that geoid values may in turn be added/updated in each patron record where an address match is found

Geoid and a last updated date are placed in a Census varfield (varfield tag k in our system) using the following structure |s state id |c county id |t tract id |b block id |d last updated date

Data from this field is then used for created choropleth maps that can utilize both demographic data from the census bureau and annoynmized statistical data from our patron records. Examples can be seen within the Patron Maps folder of this repo ([https://github.com/Minuteman-Library-Network/Scripts/tree/main/Patron-Maps]([https://github.com/Minuteman-Library-Network/Scripts/tree/main/Patron%20Maps](https://github.com/Minuteman-Library-Network/Scripts/tree/main/Patron%20Maps))).

Execution Plan
* Run query to find 10,000 highest priority patron records in need of geocoding, looking for patron records without a census field at all followed by the records with the oldest dates since the last check for this data among active patrons
* Generate csv file from query results, matching the needs of [the Census Burea's sample file](https://geocoding.geo.census.gov/geocoder/Addresses.csv)
* Run csv file through the geocoder API using censusgeocode Python library
* Loop through returned file and write census fields to each patron record using the Sierra API
* Remove all files once complete
