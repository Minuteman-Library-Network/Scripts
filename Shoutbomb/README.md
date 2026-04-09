Script used to extract Data from Sierra ILS that is needed for Shoutbomb to provide their services.  

Scripts run daily, with the shoutbomb_all script running twice and the shoutbomb_holds_only script running twice at appropriately staggered intervals

Execution Plan:
* Run SQL queries against Sierra
* Generate pipe delimited txt files from query results for each data set
* SFTP files to Shoutbomb
* Remove any local files
