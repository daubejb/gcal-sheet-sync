function onOpen(e) {
  SpreadsheetApp.getUi().createAddonMenu()
      .addItem('Start', 'showSidebar')
      .addToUi();
}

function onInstall(e) {
  onOpen(e);
}

/**
 * Opens a sidebar in the document containing the add-on's user interface.
 */
function showSidebar() {
  var ui = HtmlService.createTemplateFromFile('Sidebar').evaluate()
      .setSandboxMode(HtmlService.SandboxMode.IFRAME)
      .setTitle('Daube Communications Synchronization');
  SpreadsheetApp.getUi().showSidebar(ui);
}

/****GLOBAL STUFF****/
var ss = SpreadsheetApp.getActiveSpreadsheet();
var sheet = ss.getSheetByName('Events');
var daubeCommCal = getCalendar();




/****GET CALENDAR EVENTS FROM GOOGLE****/

function refreshEvents() {
  var events = getEventsFromGoogle();
  var eventsFormatted = getCalendarEventDetails(events);
  var eventsFormattedAndNew = filterNewEvents(eventsFormatted);
  Logger.log(eventsFormattedAndNew);
  putNewEventsOnSheet(eventsFormattedAndNew);
}

function getEventsFromGoogle() {
  var now = new Date();
  var then = new Date(now.getTime() + (3.154e+10));

  var events = daubeCommCal.getEvents(now, then);

  return events;
}

function getCalendarEventDetails(events) {
  var eventsFormatted = [];
  for (var i=0;i<events.length;i++) {
    var eventFormatted = {};
    var startTime = events[i].getStartTime();
    var endTime = events[i].getEndTime();
    var description = events[i].getDescription();
    var title = events[i].getTitle();
    var id = events[i].getId();
    eventFormatted = {
      "startTime": startTime,
      "endTime": endTime,
      "title": title,
      "description": description,
      "id": id
    };
    eventsFormatted.push(eventFormatted);  
  }
  return eventsFormatted;
}

function filterNewEvents(eN) {
  var newEvents = [];
  var EvIds = getEventIdsFromSheet();
  var stringIds = EvIds.toString();
  Logger.log(stringIds);
  for (var i=0;i<eN.length;i++) {
    var con = stringIds.indexOf(eN[i].id) 
    Logger.log(con);
    if (con == -1) {
      newEvents.push(eN[i]);
    } else {
      Logger.log(eN[i].id);
    }
  }
  return newEvents;
}
  
function tester() {

  var teststring = 'A,B,C';

  var y = teststring.indexOf("C");
  var n = teststring.indexOf("P");
  Logger.log('yes ' + y);
  Logger.log('no ' + n);

}
  


function getEventIdsFromSheet() {
  
  var lastRow = sheet.getLastRow();
  var sheetEvents = sheet.getRange(2, 1, lastRow, 1);
  var EventIds = sheetEvents.getValues();
  Logger.log('Event Ids: ' + EventIds);
  return EventIds
}
/**** PUBLISH SHEETS EVENTS TO GOOGLE CALENDAR ****/

function putNewEventsOnSheet(eF) {
  for (var i=0;i<eF.length;i++) {
  sheet.appendRow([
    eF[i].id,
    eF[i].startTime,
    eF[i].endTime,
    "Published",
    eF[i].title,
    eF[i].description
  ]);
  }
}


function publishNewEvents() {
  var event = getCalendar().createEvent('Test Push to Gcal',
                                        new Date('September 15, 2017 12:00:00 UTC'),
                                        new Date('September 15, 2017 13:00:00 UTC'),
                                        {location: 'The Moon',
                                        desription: 'Some other stuff'});
  Logger.log('Event ID: ' + event.getId());
                                                                                                                         
}
/**********************/
/** Helper Functions **/

function getCalendar() {
  return CalendarApp.getCalendarById('redhat.com_rcl08qun4gu0gdj9ac1mr1pni4@group.calendar.google.com');
}


// Helper function that puts external JS / CSS into the HTML file.
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename)
      .getContent();
}
