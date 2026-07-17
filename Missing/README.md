Script used to generate a quarterly report of items that have been marked missing for at least 60 days
Reports are produced as Excel files that are then uploaded to our staff intranet site for distribution, via sftp.

Execution Plan:
* Run query for each library
* Compile query results into an Excel File
* Upload files via sftp to staff intranet site placing it in the appropriate directory for each library