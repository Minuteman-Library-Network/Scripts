Script used to generate a monthly collection performance report, broken out by scat code.
Reports are produced as Excel files that are then uploaded to our staff intrenet site for distribution, via sftp.

Execution Plan:
* Run query for each library
* Compile query results into an Excel File
* Upload files via sftp to staff intranet site placing it in the appropriate directory for each library