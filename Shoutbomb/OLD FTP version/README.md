Script used to extract Data from Sierra ILS that is needed for Shoutbomb to provide their services.  

Scripts run daily, with the shoutbomb_all script running twice and the shoutbomb_holds_only script running twice at appropriately staggered intervals

Execution Plan:
* Run SQL queries against Sierra
* Generate pipe delimited txt files from query results for each data set
* FTP files to Shoutbomb
** handled via ftp_holds.bat file instead of script because shoutbomb's ftp site uses implict TLS for FTP, which the pysftp library does not natively support as an alternative am calling a winscp script to handle the file upload 
* Remove any local files
