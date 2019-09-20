# plc-invoice

Google Sheets
* Create/verify a "Settings" tab exists
* Set value of A1 to "calendarId"
* Navigate to "PLC Rental" calendar settings -> Integrate Calendar
* Copy calendar id value into "Settings" tab cel B1
* Set value of A2 to "rate"
* Set value of B2 to the desired $/lane hour rental rate (e.g. "15")
* Copy contents of https://github.com/LarryHanson/plc-invoice/blob/master/plc-invoice.gs into clipboard
* Go to Tools -> Scrip Editor and paste the contents completely replacing anything already there.
* Hit save and name "PLC Invoice"
* Close the code window

* Create a new tab and name tab with year and month you wish to invoice 
* "<4 digit year>-<2 digit month>" (e.g. "2019-08")
* Refresh the sheets page and you should now have a "PLC" menu item next to the "Help"
* Select PLC -> Generate invoice
* If this is the first time, you will be asked to authorize to run. Hit "Continue" and give access.
* If you know and trust the developer hit "Advanced" and "Go to PLC Invoice" and then Allow
* You should now see 

TODO: format start and end time with timezone
TODO: Format lanes and hours columns left justified.
TODO: Make sure we handle error cases gracefully

