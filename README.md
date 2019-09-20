# Swim Lane Rental Invoice Generator

## Google Sheets Setup
* Create a Google Sheets file or open one that exists. Name it whatever you like.
* Create a "Settings" sheet (tab) if it doesn't already exist.
* Set value of cell A1 to "calendarId" in the "Settings" sheet (tab).
* Navigate to "PLC Rental" Google calendar Settings -> Integrate Calendar and copy the calendar id value. 
* Paste the copied calendar id into cell B1 in the "Settings" sheet (tab). 
* Set value of A2 to "rate" in the "Settings" sheet (tab).
* Set value of B2 to the desired $/lane hour rental rate (e.g. "15") in the "Settings" sheet (tab).
* Copy contents of https://github.com/LarryHanson/swim-lane-rental-invoice/blob/master/plc-invoice.gs into clipboard
* In Google Sheets navigate to Tools -> Script Editor and paste the contents completely replacing anything already in Code.gs.
* Select File -> Save and name theh application whatever you like (e.g. "Pool Invoice").
* Close the code window
* Refresh the Google Sheets web page so that it picks up your new command. 
* You should now have a "PLC" menu item next to the "Help" menu item.

## Run the invoice generator
* Create a sheet (tab) and name it with the year and month you that matches what you want to invoice in your Google calendar.
  * The year and month name must be in the form <4 digit year>-<2 digit month>. Example: 2019-08
* Select PLC -> Generate invoice
  * If this is the first time, you will be asked to authorize to run. Hit "Continue" and give access.
  * If you know and trust the developer hit "Advanced" and "Go to PLC Invoice" and then Allow
* The sheet will now populate with invoice details from the Google calendar.
