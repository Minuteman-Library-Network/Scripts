Scripts to generate patron email notifications that are not available natively within Sierra.  Scripts are broken out into smaller patron segments in an effort to avoid tripping anti-spam measures.

Three notices are covered
* Welcome Email for patron records that were created in the past day
* Expiring Patrons, for patrons whose library cards are set to expire 30 days from now
* Expired Patrons, for patrons whose library card expiration date has passed in the prior day

Execution Plan
* Run query to identify patrons meeting the parameters for a given notification and to gather data needed to compile emails
* Loop through query results and generate an email for each patron, using data from those results to fill in information in the message body
