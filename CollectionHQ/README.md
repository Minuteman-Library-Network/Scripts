Script used to extract Data from Sierra ILS that is needed for CollectionHQ to provide their services.  Script is scheduled to be run once a month.

Execution Plan:
* Run SQL queries against Sierra
* Generate csv files from query results
* FTP files to CollectionHQ
* Remove any local files
