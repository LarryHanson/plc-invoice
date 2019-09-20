# Swim Lane Rental Invoice Generator

## Introduction
I created this custom Google Sheets menu item to generate detailed invoices for pool rentals at my swim club.

## Calender title format requirements
* Rental calendar entries must have the following form for their title: RENTAL - <tenant name> - <number> Lanes
  * Example: RENTAL - Guppy Swim Team - 4 Lanes
* Create one or more calendar entries with the proper title format with specified start and end times.
* When you run the invioce generator it will create an entry for each calendar entry with:
  * Tenant name (from the title)
  * Start date and time of the rental
  * End date and time of the rental
  * Number of lanes rented (from the title)
  * Number of hours rented
  * $/lane hour rate charged
  * Amount charged for the calendar entry (number of hours * $/lane hour rate)
* All invoice entries will be grouped by tenant with a total for each group for the month.

## Google Sheets Setup
* Create a Google Sheets file or open one that exists. Name it whatever you like.
* Create a "Settings" sheet (tab) if it doesn't already exist.
* Set value of cell A1 to "calendarId" in the "Settings" sheet.
* Navigate to "PLC Rental" Google calendar Settings -> Integrate Calendar and copy the calendar id value. 
* Paste the copied calendar id into cell B1 in the "Settings" sheet. 
* Set value of A2 to "rate" in the "Settings" sheet.
* Set value of B2 to the desired $/lane hour rental rate (e.g. "15") in the "Settings" sheet.
* Copy contents of https://github.com/LarryHanson/swim-lane-rental-invoice/blob/master/plc-invoice.gs into clipboard
* In Google Sheets navigate to Tools -> Script Editor and paste the contents completely replacing anything already in Code.gs.
* Select File -> Save and name theh application whatever you like (e.g. "Pool Invoice").
* Close the code window
* Refresh the Google Sheets web page so that it picks up your new command. 
* You should now have a "PLC" menu item next to the "Help" menu item.

## Run the invoice generator
* Create a sheet (tab) and name it with the year and month you that matches what you want to invoice in your Google calendar.
  * The year and month name must be in the form <4 digit year>-<2 digit month>. Example: 2019-08
* With this sheet active, select PLC -> Generate invoice
  * If this is the first time you run it, you will be asked to authorize to run. Hit "Continue" and give access.
  * Click "Advanced" and "Go to PLC Invoice" and then "Allow".
* The sheet will now populate with invoice details from the Google calendar.
* Any time you want to generate an invoice just use this menu item.
  * Note if you want to regenerate you need to clear out any old data as the script will only overwrite, not delete.

## Future Features
* Maybe include rental rate per tenant in settings sheet or some such.
* Make menu item text customizable through Seattings sheet (you can just change the string in the script for now).

## TODO
TODO: format start and end time with timezone (currently zulu time).
TODO: Format lanes and hours columns left justified.
TODO: Make sure we handle error cases gracefully.
TODO: Handle non rental calendar entries nicely.
