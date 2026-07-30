Script to find and correct instances in Sierra in which a bib record lacks an 008 field, which our catalog expects in order to populate some filters, including language.

Execution Plan:
* Run query to identify bibs in this state
* Use Sierra API to insert the needed 008 field, containing cursory information
