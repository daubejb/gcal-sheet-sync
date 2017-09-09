
/************************STANDARD STUFF********************************/
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

/*************************GLOBAL STUFF*********************************/
var ss = SpreadsheetApp.getActiveSpreadsheet();
var sheet = ss.getSheetByName('Events');
var daubeCommCal = getCalendar();




/*******************GET CALENDAR EVENTS FROM GOOGLE********************/

//gets calendar events for the next year and inserts new ones in the sheet
function refreshEvents() {
  var events = getEventsFromGoogle();
  var eventsFormatted = getCalendarEventDetails(events);
  var eventsFormattedAndNew = filterNewEvents(eventsFormatted);
  Logger.log(eventsFormattedAndNew);
  putNewEventsOnSheet(eventsFormattedAndNew);
}

//get calendar events for today plus one year
function getEventsFromGoogle() {
  var now = new Date();
  var then = new Date(now.getTime() + (3.154e+10));

  var events = daubeCommCal.getEvents(now, then);

  return events;
}

//parse events into the desired fields
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

//get all existing event ids from the sheet
function getEventIdsFromSheet() {
  
  var lastRow = sheet.getLastRow();
  var sheetEvents = sheet.getRange(2, 1, lastRow, 1);
  var EventIds = sheetEvents.getValues();
  Logger.log('Event Ids: ' + EventIds);
  return EventIds
}

//filter events to just new events
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

// append the new events to the sheet
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

/**** PUBLISH SHEETS EVENTS TO GOOGLE CALENDAR ****/

function publishNewEvents() {
  var data = getEventsFromSheet();
  var eventsToPublish = filterNonPublishedEvents(data);
  Logger.log('event count ' + eventsToPublish.length);
  for (i=0; i<eventsToPublish.length; i++) {
    var event = createEventOnCalendar(eventsToPublish[i]);
    var id = event.getId();
    var name = event.getTitle();
    var row = eventsToPublish[i].rowNum;
    var idCell = sheet.getRange('A' + row);
    idCell.setValue(id);
    var status = sheet.getRange('D' + row);
    status.setValue('Published');
    
    Logger.log('id ' + id + '  name ' + name + '  row ' + row);
    
  }
}

function getEventsFromSheet() {
  var lastRow = sheet.getLastRow() - 1;
  
  var sheetEvents = sheet.getRange(2, 1, lastRow, 6);
  var data = sheetEvents.getValues();
  return data;
}

function filterNonPublishedEvents(data) {
  var eventsToPublish = [];
  var lastRow = sheet.getLastRow() - 1;
  for (var i=0; i<lastRow; i++) {
    var status = data[i][3];
    if (status === 'Published') { continue; };
    
    var eventToPub = {}
    var id = data[i][0];
    var startDate = data[i][1];
    var endDate = data[i][2];
    var topic = data[i][4];
    var details = data[i][5];
    var rowNum = i + 2;
    eventToPub = {
      "id": id,
      "startDate": startDate,
      "endDate": endDate,
      "topic": topic,
      "details": details,
      "rowNum": rowNum
    };
    eventsToPublish.push(eventToPub);
  }
  return eventsToPublish;  
}
  
function createEventOnCalendar(ev) {
  var event = getCalendar().createEvent(
    ev.topic,
    new Date(ev.startDate),
    new Date(ev.endDate),
    {description: ev.details}
  );   
  return event;                                                                                                                     
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
