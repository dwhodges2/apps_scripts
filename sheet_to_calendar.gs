// Add calendar events based on rows in spreadsheet.

// ID and tab of the calendar to add events to:
var calendarID = "XXXXXXXXXXXXXXXXXX@group.calendar.google.com"
var sheetName = "Sheet1"
var sleepTime = 800 // add time to prevent rate limiting of api

function onOpen() {
    // Set up script menu
    var ui = SpreadsheetApp.getUi();
    ui.createMenu('Calendar')
        .addItem('Add seat reservations', 'seatReservations')
        .addItem('Clear calendar', 'deleteEvents')
        .addToUi();
}


function make_alert(msg) {
    var ui = SpreadsheetApp.getUi();
    ui.alert(msg);
}

function seatReservations() {
    // Clear out the current events.
    deleteEvents();
    Logger.log("**** Adding new events... ****");
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
    var eventCal = CalendarApp.getCalendarById(calendarID);
    var rows = spreadsheet.getDataRange().getValues();
    // find out how many populated rows there are
    var sheetLength = getLastPopulatedRow(spreadsheet);

    Logger.log(sheetLength);

    for (x = 0; x < sheetLength; x++) {
        var shift = rows[x];
        var startTime = shift[11];
        var endTime = shift[12];
        var title = shift[10];
        var loc = shift[1];
        var desc = null;
        var otherOptions = {
            'location': loc,
            'description': desc,
            'guests': ' ',
            'sendInvites': 'False',
        }
        Logger.log(title);
        Logger.log(startTime);
        Logger.log(endTime);
        Logger.log(otherOptions);
        eventCal.createEvent(title, startTime, endTime, otherOptions);
        Utilities.sleep(sleepTime);
    }
    Logger.log("Events added: " + x);
    make_alert("Events added: " + x);
}


function deleteEvents() {
    var fromDate = new Date(2021, 0, 0, 0, 0, 0);
    var toDate = new Date(2022, 0, 0, 0, 0, 0);
    // delete from Jan 1 2021 to Jan 1 2022(for month 0 = Jan, 1 = Feb...)
    var eventCal = CalendarApp.getCalendarById(calendarID);

    var events = eventCal.getEvents(fromDate, toDate);
    Logger.log("*** Deleting old events... ***");
    for (var i = 0; i < events.length; i++) {
        var ev = events[i];
        Logger.log(ev.getTitle()); // show event name in log
        ev.deleteEvent();
        Utilities.sleep(sleepTime);
    }
}


function createTimeDrivenTriggers() {

    // Trigger every day at 01:00.
    // https://developers.google.com/apps-script/reference/script/clock-trigger-builder#atHour(Integer)
    ScriptApp.newTrigger('seatReservations')
        .timeBased()
        .atHour(1)
        .everyDays(1)
        .create();
}

// Function borrowed from here:
// https://support.google.com/docs/thread/4526642?hl=en
// Addresses problem of inability to obtain getMaxRow() when there is an arrayformula in one or more cols.
function getLastPopulatedRow(sheet) {
    var data = sheet.getDataRange().getValues();
    for (var i = data.length - 1; i > 0; i--) {
        for (var j = 0; j < data[0].length; j++) {
            if (data[i][j]) return i + 1;
        }
    }
    return 0;
}