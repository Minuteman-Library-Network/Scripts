Script used to batch update the hold pickup locations for all holds at one location to a different location.  Used on demand general when a library location is forced to close for some period of time and another location is established as an alternate location.

Execution Plan
<ul>
<li>Ask user to enter the current pickup location to be changed and the new value to change it to.  Save entered values to variables</li>
<ul><li>Verify entered codes are valid, if not ask user to try again</li></ul>
<li>Run SQL query to identify all holds matching the entered current pickup location</li>
<li>Loop through query results and for each entry use the SierraAPI to update the pickuplocation to the new value</li>
</ul>
