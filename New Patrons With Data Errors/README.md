Script used to generate a monthly report of new patron records the include common data entry errors such as invalid barcodes or malformed addresses.
Reports are produced as Excel files that are then uploaded to our staff intrenet site for distribution, via sftp.

Execution Plan:
* Run query for each library
* Compile query results into an Excel File
* Upload files via sftp to staff intranet site placing it in the appropriate directory for each library