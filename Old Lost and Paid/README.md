Script used to generate a quarterly report of items that were marked lost and paid over 90 days previous
Reports are produced as Excel files that are then uploaded to our staff intranet site for distribution, via sftp.  Staff are emailed to alert them to the new reports.

Execution Plan:
* Run query for each library
* Compile query results into an Excel File
* Upload files via sftp to staff intranet site placing it in the appropriate directory for each library
* email staff that new reports are available