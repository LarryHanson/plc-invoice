/**
 * @OnlyCurrentDoc Limits the script to only accessing the current sheet.
 */

/**
  * A Google App Script that reads PLC rental event information from Google Calendar and generates an invoice in Google Sheets.
  **/

/**
 * A special function that runs when the spreadsheet is open, used to add a
 * custom menu to the spreadsheet.
 */
function onOpen() {
  var spreadsheet = SpreadsheetApp.getActive();
  var menuItems = [
    {name: 'Generate invoice', functionName: 'generateInvoice_'}
  ];
  spreadsheet.addMenu('PLC', menuItems);
}

/**
 * A function that generates a rental invoice.
 **/
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
  **/
function formatInvoiceSheet_(invoiceSheet) {

 var moneyFormat = "$#,##0.00;$(#,##0.00)"
  
  var rateColuumn = invoiceSheet.getRange("F2:F");  
  rateColuumn.setNumberFormat(moneyFormat);

 var totalsColumn = invoiceSheet.getRange("G2:G");  
  totalsColumn.setNumberFormat(moneyFormat);
}

/**
 * A function that populates the invoice sheet with the rental details.
 **/
function populateInvoiceSheet_(invoiceSheet, rentalDetails, rate) {
  var row = 2;
  for(tenant in rentalDetails) {
    var array = rentalDetails[tenant];
    var tenantEntries = 0;
    for(i = 0; i < array.length; i++) {
      var rd = array[i];

      var range = invoiceSheet.getRange(row, 1, 1, 7);
      range.setValues([[
        rd.tenant, 
        rd.startTime.toISOString(), 
        rd.endTime.toISOString(), 
        rd.lanes,
        "=((DATEVALUE(MID(R[0]C[-2],1,10)) + TIMEVALUE(MID(R[0]C[-2],12,8))) - (DATEVALUE(MID(R[0]C[-3],1,10)) + TIMEVALUE(MID(R[0]C[-3],12,8))))*24", 
        15, 
        "=R[0]C[-3]*R[0]C[-2]*R[0]C[-1]"
      ]]);

      tenantEntries = tenantEntries + 1;
      row = row + 1;   
    }
    
    var totalRange = invoiceSheet.getRange(row, 6, 1, 2);
    totalRange.setValues([["Total:", "=sum(R[-1]C[0]:R[-" + tenantEntries +"]C[0])"]]);
    totalRange.setFontWeight('bold');

    row = row + 2;
  }
  SpreadsheetApp.flush();
}

/**
 * A function that retrieves values from the 'Settings' sheet.
 **/
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
 **/
function collectRentalDetails_(yearAndMonth, calendarId) {
  Logger.log("calendarId: " + calendarId);
  var match = yearAndMonth.match(/^(\d\d\d\d)-(\d\d)$/);
  if(!match[0]) {
    return "Sheet title must be in the form <4 digit year>-<2 digit month>";
  }
  var year = match[1];
  var month = match[2]
  var startTime = new Date(year, month - 1, 1, 0, 0, 0, 0);
  var endTime = new Date(year, month, 1, 0, 0, 0, 0);
  
  Logger.log("startTime: " + startTime);
  Logger.log("endTime: " + endTime);
  
  var calendar = CalendarApp.getCalendarById(calendarId);
  var events =  calendar.getEvents(startTime, endTime)
  Logger.log("events.length: " + events.length);
  var rentalDetails = {};
  for (i = 0; i < events.length; i++) {
    var event = events[i];    
    Logger.log(event.getTitle());
    rentalDetail = createRentalDetail_(event.getTitle(), event.getStartTime(), event.getEndTime());
    if(rentalDetail.tenant in rentalDetails) {
      rentalDetails[rentalDetail.tenant].push(rentalDetail);
    } else {
      rentalDetails[rentalDetail.tenant] = [rentalDetail];
    }
  }
  return rentalDetails;    
}

function createRentalDetail_(title, start, end) {
  var match = title.match(/RENTAL - ([^-]+) - (\d+) Lanes/);
  var rentalDetail
  if(match[0]) {
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

/////////////// Everything from here down is from the sample code. ///////////////

/**
 * A custom function that converts meters to miles.
 *
 * @param {Number} meters The distance in meters.
 * @return {Number} The distance in miles.
 */
function metersToMiles(meters) {
  if (typeof meters != 'number') {
    return null;
  }
  return meters / 1000 * 0.621371;
}

/**
 * A custom function that gets the driving distance between two addresses.
 *
 * @param {String} origin The starting address.
 * @param {String} destination The ending address.
 * @return {Number} The distance in meters.
 */
function drivingDistance(origin, destination) {
  var directions = getDirections_(origin, destination);
  return directions.routes[0].legs[0].distance.value;
}

/**
 * Creates a new sheet containing step-by-step directions between the two
 * addresses on the "Settings" sheet that the user selected.
 */
function generateStepByStep_() {
  var spreadsheet = SpreadsheetApp.getActive();
  var settingsSheet = spreadsheet.getSheetByName('Settings');
  settingsSheet.activate();

  // Prompt the user for a row number.
  var selectedRow = Browser.inputBox('Generate step-by-step',
      'Please enter the row number of the addresses to use' +
      ' (for example, "2"):',
      Browser.Buttons.OK_CANCEL);
  if (selectedRow == 'cancel') {
    return;
  }
  var rowNumber = Number(selectedRow);
  if (isNaN(rowNumber) || rowNumber < 2 ||
      rowNumber > settingsSheet.getLastRow()) {
    Browser.msgBox('Error',
        Utilities.formatString('Row "%s" is not valid.', selectedRow),
        Browser.Buttons.OK);
    return;
  }

  // Retrieve the addresses in that row.
  var row = settingsSheet.getRange(rowNumber, 1, 1, 2);
  var rowValues = row.getValues();
  var origin = rowValues[0][0];
  var destination = rowValues[0][1];
  if (!origin || !destination) {
    Browser.msgBox('Error', 'Row does not contain two addresses.',
        Browser.Buttons.OK);
    return;
  }

  // Get the raw directions information.
  var directions = getDirections_(origin, destination);

  // Create a new sheet and append the steps in the directions.
  var sheetName = 'Driving Directions for Row ' + rowNumber;
  var directionsSheet = spreadsheet.getSheetByName(sheetName);
  if (directionsSheet) {
    directionsSheet.clear();
    directionsSheet.activate();
  } else {
    directionsSheet =
        spreadsheet.insertSheet(sheetName, spreadsheet.getNumSheets());
  }
  var sheetTitle = Utilities.formatString('Driving Directions from %s to %s',
      origin, destination);
  var headers = [
    [sheetTitle, '', ''],
    ['Step', 'Distance (Meters)', 'Distance (Miles)']
  ];
  var newRows = [];
  for (var i = 0; i < directions.routes[0].legs[0].steps.length; i++) {
    var step = directions.routes[0].legs[0].steps[i];
    // Remove HTML tags from the instructions.
    var instructions = step.html_instructions.replace(/<br>|<div.*?>/g, '\n')
        .replace(/<.*?>/g, '');
    newRows.push([
      instructions,
      step.distance.value
    ]);
  }
  directionsSheet.getRange(1, 1, headers.length, 3).setValues(headers);
  directionsSheet.getRange(headers.length + 1, 1, newRows.length, 2)
      .setValues(newRows);
  directionsSheet.getRange(headers.length + 1, 3, newRows.length, 1)
      .setFormulaR1C1('=METERSTOMILES(R[0]C[-1])');

  // Format the new sheet.
  directionsSheet.getRange('A1:C1').merge().setBackground('#ddddee');
  directionsSheet.getRange('A1:2').setFontWeight('bold');
  directionsSheet.setColumnWidth(1, 500);
  directionsSheet.getRange('B2:C').setVerticalAlignment('top');
  directionsSheet.getRange('C2:C').setNumberFormat('0.00');
  var stepsRange = directionsSheet.getDataRange()
      .offset(2, 0, directionsSheet.getLastRow() - 2);
  setAlternatingRowBackgroundColors_(stepsRange, '#ffffff', '#eeeeee');
  directionsSheet.setFrozenRows(2);
  SpreadsheetApp.flush();
}

/**
 * Sets the background colors for alternating rows within the range.
 * @param {Range} range The range to change the background colors of.
 * @param {string} oddColor The color to apply to odd rows (relative to the
 *     start of the range).
 * @param {string} evenColor The color to apply to even rows (relative to the
 *     start of the range).
 */
function setAlternatingRowBackgroundColors_(range, oddColor, evenColor) {
  var backgrounds = [];
  for (var row = 1; row <= range.getNumRows(); row++) {
    var rowBackgrounds = [];
    for (var column = 1; column <= range.getNumColumns(); column++) {
      if (row % 2 == 0) {
        rowBackgrounds.push(evenColor);
      } else {
        rowBackgrounds.push(oddColor);
      }
    }
    backgrounds.push(rowBackgrounds);
  }
  range.setBackgrounds(backgrounds);
}

/**
 * A shared helper function used to obtain the full set of directions
 * information between two addresses. Uses the Apps Script Maps Service.
 *
 * @param {String} origin The starting address.
 * @param {String} destination The ending address.
 * @return {Object} The directions response object.
 */
function getDirections_(origin, destination) {
  var directionFinder = Maps.newDirectionFinder();
  directionFinder.setOrigin(origin);
  directionFinder.setDestination(destination);
  var directions = directionFinder.getDirections();
  if (directions.status !== 'OK') {
    throw directions.error_message;
  }
  return directions;
}
