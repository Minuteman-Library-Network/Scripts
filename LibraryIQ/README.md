Script used to extract Data from Sierra ILS that is needed for libraryIQ to provide their services.  Script is scheduled to be run once a day, with delta files being generated most days and full data extracts gathered once a week.

Execution Plan:
* Run SQL queries against Sierra
* Generate csv files from query results
* SFTP files to libraryIQ
* Remove any local files
