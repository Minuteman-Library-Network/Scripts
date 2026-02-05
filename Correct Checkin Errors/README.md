Script to find instances in Sierra in which an item is simultaneously checked out and in transit.  

This error is caused by leaving a patron account open after checking an item out, which is immediately returned at a different terminal.  In that scenario the checkin does not register, but information in the item record is updated as if it had.

Execution Plan:
* Run query to identify items in this state
* Log information to Google Sheet in case follow up is needed
* Use Sierra API to check the item in again, clearing the error
