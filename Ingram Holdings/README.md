Python script for sending minimal MARC exports to Ingram for use with Ipage's holdings feature.

For this feature Ingram requires .mrc or .out files of MARC records.  However, the feature only looks at the 020 and 024 fields for matching purposes.  Accordingly this script creates pseudo-MARC records that contain the minimal amount of data necessary for their needs and serves to automate the process of building those records and ftping them to Ingram.

Separate SQL queries are maintained for each library using the feature as they each may have different parameters for the holdings they wish to include.  For example one may wish to exclude items within a special collection while another may want to include titles from outstanding orders.

Execution Plan
* Run SQL query to gather holdings for a particular library
* Use results of Query to generate file of minimal MARC records using the Pymarc library for Python*
* SFTP resulting file to Ingram
* Remove local files
