/**
 * @OnlyCurrentDoc Limits the script to only accessing the current sheet.
 */

/**
  * A Google App Script that reads rental event information from Google Calendar and generates an invoice in Google Sheets.
  */

// Pull in moment.js
eval(UrlFetchApp.fetch('https://cdnjs.cloudflare.com/ajax/libs/moment.js/2.24.0/moment.js').getContentText());

/**
 * A special function that runs when the spreadsheet is open, used to add a
 * custom menu to the spreadsheet.
 */
function onOpen() {
    var spreadsheet = SpreadsheetApp.getActive();
    var menuItems = [
        { name: 'Generate invoice', functionName: 'generateInvoice_' }
    ];
    spreadsheet.addMenu('Invoice', menuItems);
}

/**
 * A function that generates a rental invoice.
 */
function generateInvoice_() {
    prepareSheet_();
    var invoiceSheet = SpreadsheetApp.getActiveSheet();
    var sheetName = invoiceSheet.getName();
    Logger.log("invoiceSheet.sheetName: " + sheetName);

    var settings = getSettings_();
    Logger.log("settings.rate: " + settings.rate);
    var rentalDetails = collectRentalDetails_(sheetName, settings.calendarId);
    Logger.log(rentalDetails);

    populateInvoiceSheet_(invoiceSheet, rentalDetails, settings.rate);

    formatInvoiceSheet_(invoiceSheet);

    Logger.log("Processing completed");
}

/**
  * A function that formats the invoice sheet.
  */
function formatInvoiceSheet_(invoiceSheet) {

    var moneyFormat = "$#,##0.00;$(#,##0.00)"

    var rateColuumn = invoiceSheet.getRange("F2:F");
    rateColuumn.setNumberFormat(moneyFormat);

    var totalsColumn = invoiceSheet.getRange("G2:G");
    totalsColumn.setNumberFormat(moneyFormat);
}

/**
 * A function that populates the invoice sheet with the rental details.
 */
function populateInvoiceSheet_(invoiceSheet, rentalDetails, rate) {
    var row = 2;
    for (tenant in rentalDetails) {
        var array = rentalDetails[tenant];
        var tenantEntries = 0;
        for (i = 0; i < array.length; i++) {
            var rd = array[i];

            var range = invoiceSheet.getRange(row, 1, 1, 7);
            range.setValues([[
                rd.tenant,
                moment(rd.startTime).toISOString(true),
                moment(rd.endTime).toISOString(true),
                rd.lanes,
                "=((DATEVALUE(MID(R[0]C[-2],1,10)) + TIMEVALUE(MID(R[0]C[-2],12,8))) - (DATEVALUE(MID(R[0]C[-3],1,10)) + TIMEVALUE(MID(R[0]C[-3],12,8))))*24",
                15,
                "=R[0]C[-3]*R[0]C[-2]*R[0]C[-1]"
            ]]);

            tenantEntries = tenantEntries + 1;
            row = row + 1;
        }

        var totalRange = invoiceSheet.getRange(row, 6, 1, 2);
        totalRange.setValues([["Total:", "=sum(R[-1]C[0]:R[-" + tenantEntries + "]C[0])"]]);
        totalRange.setFontWeight('bold');

        row = row + 2;
    }
    SpreadsheetApp.flush();
}

/**
 * A function that retrieves values from the 'Settings' sheet.
 */
function getSettings_() {
    var settingsSheet = SpreadsheetApp.getActive().getSheetByName("Settings");
    var row = settingsSheet.getRange(1, 2, 2, 1);
    var rowValues = row.getValues();
    return {
        calendarId: rowValues[0][0],
        rate: rowValues[1][0]
    }
}

/**
 * A function that collects rental event details from the calendar in a map of details bucketed by tenant.
 */
function collectRentalDetails_(yearAndMonth, calendarId) {
    Logger.log("calendarId: " + calendarId);
    var match = yearAndMonth.match(/^(\d\d\d\d)-(\d\d)$/);
    if (!match[0]) {
        return "Sheet title must be in the form <4 digit year>-<2 digit month>";
    }
    var year = match[1];
    var month = match[2]
    var startTime = new Date(year, month - 1, 1, 0, 0, 0, 0);
    var endTime = new Date(year, month, 1, 0, 0, 0, 0);

    Logger.log("startTime: " + startTime);
    Logger.log("endTime: " + endTime);

    var calendar = CalendarApp.getCalendarById(calendarId);
    var events = calendar.getEvents(startTime, endTime)
    Logger.log("events.length: " + events.length);
    var rentalDetails = {};
    for (i = 0; i < events.length; i++) {
        var event = events[i];
        Logger.log(event.getTitle());
        rentalDetail = createRentalDetail_(event.getTitle(), event.getStartTime(), event.getEndTime());
        if (rentalDetail.tenant in rentalDetails) {
            rentalDetails[rentalDetail.tenant].push(rentalDetail);
        } else {
            rentalDetails[rentalDetail.tenant] = [rentalDetail];
        }
    }
    return rentalDetails;
}

/**
 * A function that creates a rental detail record.
 */
function createRentalDetail_(title, start, end) {
    var match = title.match(/RENTAL - ([^-]+) - (\d+) Lanes/);
    var rentalDetail
    if (match[0]) {
        return {
            tenant: match[1],
            startTime: start,
            endTime: end,
            lanes: match[2]
        };
    } else {
        return {
            tenant: "Event title must be in the form 'RENTAL - <tenant> - <number> Lanes': " + title,
            startTime: start,
            endTime: end,
            lanes: 0
        };
    }
}

/**
 * A function that adds headers and formatting to sheet.
 */
function prepareSheet_() {
    var sheet = SpreadsheetApp.getActiveSheet()
    var headers = [
        'Tenant',
        'Start Time',
        'End Time',
        'Lanes',
        'Hours',
        '$/Lane Hour'];

    sheet.getRange('A1:F1').setValues([headers]).setFontWeight('bold');
}
