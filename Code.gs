/**
 * Record timesheet for each project via event in Hangouts Chat and store to Spreedsheet
 * Author: octoberstorm <storm.october@gmail.com>
 *
 * Supported actions: Checkin, Checkout, Get my today timesheet, Get all project members' timesheet
 * Commands: checkin, checkout, get, getall
 *
 */

// link to local timesheet file
const TIMESHEET_FILE = '<https://docs.google.com/spreadsheets/d/<SHEET_ID>/edit>';
// link to log file
const LOG_FILE = 'https://docs.google.com/spreadsheets/d/<SHEET_ID>/edit';
const CHECKIN_COLUMN = 'C';
const CHECKOUT_COLUMN = 'D';
const LOG_ENABLED = true;

/**
 * Receive message from Hangouts
 * @param {Object} event the event object from Hangouts Chat
*/
function onMessage(event) {
    var t = new Date(event.message.createTime.seconds * 1000);

    var name = "";
  
    if (event.space.type == "DM") {
      name = "You";
    } else {
      name = event.user.displayName;
    }
    var message = name + " said \"" + event.message.text + "\"";
    
    // for event logging
    var logSS = SpreadsheetApp.openByUrl(LOG_FILE);
    SpreadsheetApp.setActiveSpreadsheet(logSS);
    var logSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Logs');
    logSheet.appendRow([event.space.displayName, name, event.message.text, t]);

    var ss = SpreadsheetApp.openByUrl(TIMESHEET_FILE);
    SpreadsheetApp.setActiveSpreadsheet(ss);
    var project = event.space.displayName;

    // main working sheet of the project
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(project);

    var year = t.getFullYear();

    var month = (t.getMonth() + 1).toString();
    if (month.length == 1) {
        month = "0" + month;
    }

    var day = t.getDate().toString();
    if (day.length == 1) {
        day = "0" + day;
    }

    var hours = t.getHours().toString();
    if (hours.length == 1) {
        hours = "0" + hours;
    }

    var mins = t.getMinutes().toString();
    if (mins.length == 1) {
        mins = "0" + mins;
    }

    var date = "'" + day + "-" + month + "-" + year;
    var time = "'" + hours + ":" + mins;

    var m = event.message.text;
  
    // get all members timesheet
    if (m.includes('getall')) {
        return { text: getAllMembersTimesheet(sheet, date) };
    }
  
    // get my timesheet
    if (m.includes('get')) {
      return { text: '<' + event.user.name + '> ' + getMyTimesheet(sheet, date, name) };
    }  

    var timesheetType = '';
    var updateCol = '';
    if (m.includes('checkin') || m.includes('Checkin') || m.includes('CheckIn')) {
      timesheetType = 'Checkin';
      updateCol = CHECKIN_COLUMN;
    } else if (m.includes('checkout') || m.includes('Checkout') || m.includes('CheckOut'))  {
      timesheetType = 'Checkout';
      updateCol = CHECKOUT_COLUMN;
    }
    
    var row = findTimesheetRow(sheet, date, name);
    
    if (row == -1) {
      // new record
      if (timesheetType == 'Checkin') {
        sheet.appendRow([date, name, time, '', event.space.displayName]);
      } else if (timesheetType == 'Checkout') {
        sheet.appendRow([date, name, '', time, event.space.displayName]);
      }
    } else {
      // update existing record
        sheet.getRange(updateCol + row).setValue(time);
    }

    return { "text": '<' + event.user.name + '> ' + timesheetType + ' logged for ' + name + ' at ' + t };
  }
  
  /**
   * Responds to an ADDED_TO_SPACE event in Hangouts Chat.
   *
   * @param {Object} event the event object from Hangouts Chat
   */
  function onAddToSpace(event) {
    var message = "";
  
    if (event.space.type == "DM") {
      message = "Thank you for adding me to a DM, " + event.user.displayName + "!";
    } else {
      message = "Thank you for adding me to " + event.space.displayName;
    }
  
    if (event.message) {
      // Bot added through @mention.
      message = message + " and you said: \"" + event.message.text + "\"";
    }
  
    return { "text": message };
  }
  
  /**
   * Responds to a REMOVED_FROM_SPACE event in Hangouts Chat.
   *
   * @param {Object} event the event object from Hangouts Chat
   */
  function onRemoveFromSpace(event) {
    console.info("Bot removed from ", event.space.name);
  }
  
  /**
   * Find timesheet row of a member in given date
   * 
   * @param {SheetObject} sheet 
   * @param {string} date 
   * @param {string} name 
   */
  function findTimesheetRow(sheet, date, name) {
    var dataRange = sheet.getDataRange();
    var values = dataRange.getValues();
  
    for (var i = 0; i < values.length; i++) {
      if (date.includes(values[i][0])) {
        if (values[i][1] == name) {
          return i + 1;
        }
      }
    }
    return -1;
  }

function getMyTimesheet(sheet, date, name) {
  var row = findTimesheetRow(sheet, date, name);
  
  if (row == -1) {
    return 'No data yet';
  }
  
  Logger.log('found row:', row);
  var dataRange = sheet.getDataRange();
  var values = dataRange.getValues();
  
  var data = values[row-1];
  
  return data[1] + ' - ' + data[0] + ' - IN: ' + data[2] + ' - OUT: ' + data[3];
}

function getAllMembersTimesheet(sheet, date) {
  var dataRange = sheet.getDataRange();
  var values = dataRange.getValues();
  
  var data = [["DATE - NAME - TIME IN - TIMEOUT - PROJECT"]];
  
  for (var i = 0; i < values.length; i++) {
    if (date.includes(values[i][0])) {
      data.push(values[i].join(' - '));
    }
  }
  
  Logger.log(data);
  
  return data.join('\n');
}