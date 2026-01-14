This script will gather 10,000 patron records at a time via a SQL query and pass the addresses of those records through the US Census Bureau's geocoding service so that geoid values may in turn be added/updated in each patron record where an address match is found

Geoid and a last updated date are placed in a Census varfield (varfield tag k in our system) using the following structure |s state id |c county id |t tract id |b block id |d last updated date

Data from this field is then used for created choropleth maps that can utilize both demographic data from the census bureau and annoynmized statistical data from our patron records. Examples can be seen within the Patron Maps folder of this repo ([https://github.com/Minuteman-Library-Network/Patron-Maps](https://github.com/Minuteman-Library-Network/Scripts/tree/main/Patron%20Maps)).
