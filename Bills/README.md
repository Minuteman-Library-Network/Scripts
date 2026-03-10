Run on demand script that takes a file of bills (for patrons lacking email addresses), splits the file by the library associated with each patron and reformats the text of each bill to a word doc suitable for mailing.
Files are then uploaded to Minuteman's intranet site for distribution to our members.

Script was heavily vibe coded using Claude

Execution Plan
* Parse .txt file to divide up bills by each member library
* produce .doc file for each library with a bill within the .txt file, with bills formatted for mailing
* Upload files to each library's respective directory within our intranet site
