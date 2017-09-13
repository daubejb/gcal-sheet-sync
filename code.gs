/*************************GLOBAL STUFF*********************************/
var ss = SpreadsheetApp.getActiveSpreadsheet();
var sheet = ss.getSheetByName('Events');
var daubeCommCal = getCalendar();
var filterSheet = ss.getSheetByName('Filter');

/*************************SETUP STUFF**********************************/
function onOpen(e) {
  SpreadsheetApp.getUi().createAddonMenu()
      .addItem('Start', 'showSidebar')
      .addToUi();
      showSidebar();
  
  SpreadsheetApp.getUi()
      .createMenu('Custom Menu')
      .addItem('Show alert', 'showAlert')
      .addToUi();
}

function onInstall(e) {
  onOpen(e);
}

function showAlert(type, prompt, message) {

  var ui = SpreadsheetApp.getUi(); // Same variations.
  if (type == 'delete') {
    var result = ui.alert(
      prompt,
      message,
      ui.ButtonSet.YES_NO);
    // Process the user's response.
    if (result == ui.Button.YES) {
      var userResponse = "yes";
    } else {
      var userResponse = "no";
    }
    return userResponse;
  }
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

/*******************GET CALENDAR EVENTS FROM GOOGLE********************/

//gets calendar events for the next year and inserts new ones in the sheet
function refreshEvents() {
  var events = getEventsFromGoogle();
  var eventsFormatted = getCalendarEventDetails(events);
  var eventsFormattedAndNew = filterNewEvents(eventsFormatted);
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
    var description = events[i].getDescription();
    var title = events[i].getTitle();
    var id = events[i].getId();
    eventFormatted = {
      "startTime": startTime,
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
  return EventIds
}

//filter events to just new events
function filterNewEvents(eN) {
  var newEvents = [];
  var EvIds = getEventIdsFromSheet();
  var stringIds = EvIds.toString();
  for (var i=0;i<eN.length;i++) {
    var con = stringIds.indexOf(eN[i].id) 
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
    "Published",
    eF[i].title,
    eF[i].description
  ]);
  }
}

/**** PUBLISH SHEETS EVENTS TO GOOGLE CALENDAR ****/

function publishNewEvent(pubRowNum) {
  var a1 = 'A' + pubRowNum + ':' + 'E' + pubRowNum;
  var data = sheet.getRange(a1).getValues();
  Logger.log(data);
  var eventsToPublish = [];
  var eventToPub = {}
    var id = data[0][0];
    var startDate = data[0][1];
    var topic = data[0][3];
    var details = data[0][4];
    var rowNum = pubRowNum;
    eventToPub = {
      "id": id,
      "startDate": startDate,
      "topic": topic,
      "details": details,
      "rowNum": rowNum
    };
    eventsToPublish.push(eventToPub);
    var event = createEventOnCalendar(eventsToPublish[0]);
    var id = event.getId();
    var row = eventsToPublish[0].rowNum;
    var idCell = sheet.getRange('A' + row);
    idCell.setValue(id);
    var status = sheet.getRange('C' + row);
    status.setValue('Published');
}

function publishNewEvents() {
  var data = getEventsFromSheet();
  Logger.log(data);
  var eventsToPublish = filterNonPublishedEvents(data);
  for (i=0; i<eventsToPublish.length; i++) {
    var event = createEventOnCalendar(eventsToPublish[i]);
    var id = event.getId();
    var row = eventsToPublish[i].rowNum;
    var idCell = sheet.getRange('A' + row);
    idCell.setValue(id);
    var status = sheet.getRange('C' + row);
    status.setValue('Published');    
  }
}

function getEventsFromSheet() {
  var lastRow = sheet.getLastRow() - 1;
  
  var sheetEvents = sheet.getRange(2, 1, lastRow, 5);
  var data = sheetEvents.getValues();
  return data;
}

function filterNonPublishedEvents(data) {
  var eventsToPublish = [];
  var lastRow = sheet.getLastRow() - 1;
  for (var i=0; i<lastRow; i++) {
    var status = data[i][2];
    if (status === 'Published') { continue; };
    
    var eventToPub = {}
    var id = data[i][0];
    var startDate = data[i][1];
    var topic = data[i][3];
    var details = data[i][4];
    var rowNum = i + 2;
    eventToPub = {
      "id": id,
      "startDate": startDate,
      "topic": topic,
      "details": details,
      "rowNum": rowNum
    };
    eventsToPublish.push(eventToPub);
  }
  return eventsToPublish;  
}
  
function createEventOnCalendar(ev) {
  var event = getCalendar().createAllDayEvent(
    ev.topic,
    new Date(ev.startDate),
    {description: ev.details}
  );   
  return event;                                                                                                                     
}

/**** DELETE EVENT FROM CALENDAR AND SHEET ****/

function deleteEvent(delRowNum) {
  var deleteId = sheet.getRange(delRowNum, 1).getValue();
  Logger.log('row number: ' + delRowNum + 'id: ' + deleteId);
  var cal = getCalendar();
  var event = cal.getEventSeriesById(deleteId);
  event.deleteEventSeries();
  sheet.deleteRow(delRowNum);
  filterSheet.deleteRow(delRowNum);
}


/**** GENERAL SHEET EDIT DETECTION ****/

function onEdit(e){
  var range = e.range;
  var editColumn = range.getColumn();
  var editSheet = range.getSheet();
  var editSheetName = editSheet.getName();
  var anImportantChange = anImportantFieldChanged(editColumn, editSheetName);
  if (anImportantChange) {
    var editRow = range.getRow();
    var updateStatusCell = sheet.getRange('C' + editRow);
    updateStatusCell.setValue('Update calendar');
  }
}

function anImportantFieldChanged(col, editSheetName) {
  if ((col == '2.0' || c == '4.0' || col == '5.0') && (editSheetName == 'Events')) {
    return true;  
  } else {
    return false;  
  }
}
  



/**********************/
/** Helper Functions **/

function getCalendar() {
  return CalendarApp.getCalendarById(
    'redhat.com_rcl08qun4gu0gdj9ac1mr1pni4@group.calendar.google.com');
}

function displayToast(m) {
  ss.toast("An error occured:" + m);
}
