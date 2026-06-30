Script used to extract Data from Sierra ILS that is needed for Orangeboy's Savannah service.  Script is scheduled to be run once a week.  Script is run separately for each member library subscribing to the service

Execution Plan:
* Run SQL queries against Sierra
* Generate csv files from query results
* SFTP files to Orangeboy
* Remove any local files
