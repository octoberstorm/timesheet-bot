/**
 * Sync local timesheet data to official sheet
 * Author: octoberstorm <storm.october@gmail.com>
 */

const LOCAL_SHEET_URL = 'https://docs.google.com/spreadsheets/d/<LOCAL_SHEET_ID>/edit';
const OFFICIAL_SHEET_URL = 'https://docs.google.com/spreadsheets/d/<OFFICIAL_SHEET_ID>/edit';
const OFFICIAL_SHEET_NAME = getCurrentCheckinMonth(); // e.g. '08.2021';
const LOCAL_MEMBER_MAPPING_SHEET_NAME = 'MemberMapping';
const LOCAL_PROJECT_LIST_SHEET_NAME = 'ProjectList';

var projectList = [];
var timesheetColumns;
var updatedMembers = [];
var notFoundMembers = [];

// main function
function syncTimesheet() {
  var localSS = SpreadsheetApp.openByUrl(LOCAL_SHEET_URL);
  var officialSS = SpreadsheetApp.openByUrl(OFFICIAL_SHEET_URL);
  
  // get local data
  SpreadsheetApp.setActiveSpreadsheet(localSS);

  projectList = getProjectList();

  var nameList = getNameList();
  
  var date = getTodayDate();
  // var date = '24-12-2021';

  var lastDayOfTimesheetMonth = false;
  var hours = new Date().getHours();
  if (date.split("-")[0] === '23' && hours > 12) {
    lastDayOfTimesheetMonth = true;
  }

  var localData = findLocalData(nameList, date);
  
  // Now sending data to official sheet
  SpreadsheetApp.setActiveSpreadsheet(officialSS);
  var officialSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(OFFICIAL_SHEET_NAME);
    
  var dataRange = officialSheet.getDataRange();
  var values = dataRange.getValues();

  timesheetColumns = getTimesheetCols(values);

  var updatedCount = 0;
  
  for (var i in localData) {
    var item = localData[i];
    
    var name = item.name;
    
    if (name == null) {
      continue;
    }

    var timeIn = item.in;
    var timeOut = item.out;

    if (timeIn != '' && timeOut === '' && lastDayOfTimesheetMonth) {
      timeOut = '18:45';
    }
    
    var row = findTimesheetRowNumber(values, name);
    
    if (row == -1) {
      Logger.log('not found member in official sheet: ' + name);
      notFoundMembers.push(name);
      continue;
    }
    
    var cols = findTimesheetColums(date);

    officialSheet.getRange(row, cols.timeIn).setValue(timeIn);
    officialSheet.getRange(row, cols.timeOut).setValue(timeOut);
    
    updatedCount++;
    updatedMembers.push(name);
  }
  
  const body = `<b>Updated: ${updatedCount} member's timesheets for ${date}</b><br><br>
    - Projects: [${projectList}] <br><br>
    - Not found members: ${JSON.stringify(notFoundMembers)}<br><br>
    - Updated Members: ${JSON.stringify(updatedMembers)} <br> <br>
    - Timesheet Columns: ${JSON.stringify(timesheetColumns)}`;

  MailApp.sendEmail({
    to: "logmail@email.example",
    subject: `[TimesheetBotSync] ${date}, ${updatedCount} members updated`,
    htmlBody: body
  })
}

function findLocalData(nameList, date) {
  var data = [];

  for (var i in projectList) {
    var projectName = projectList[i];
    var projectSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(projectName);
    var dataRange = projectSheet.getDataRange();
    var values = dataRange.getValues();
    
    for (var i = 1; i < values.length; i++) {
      if (date.includes(values[i][0])) {
        var officialName = findOfficialName(nameList, values[i][1]);
        
        data.push({name: officialName, in: values[i][2], out: values[i][3]});
      }
    }
  }

  return data;
}

function findOfficialName(nameList, gsuiteName) {
  for (var i in nameList) {
    if (nameList[i].g == gsuiteName) {
      return nameList[i].o;
    }
  }

  notFoundMembers.push(gsuiteName);
}

// get official name from Gsuite name
// return [{officialName - o, gsuiteName - g}
function getNameList() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(LOCAL_MEMBER_MAPPING_SHEET_NAME);
  
  var dataRange = sheet.getDataRange();
  var values = dataRange.getValues();
  
  var names = [];
  
  for (var i = 0; i < values.length; i++) {
    names.push({g: values[i][0], o: values[i][1]});
  }
  
  return names;
}

function findTimesheetRowNumber(values, name) {
  for (var i = 0; i < values.length; i++) {
    if (values[i][2] == name) {
      return i + 1;
    }
  }
  return -1;
}

function findTimesheetColums(date) {
  return timesheetColumns[date];
}

function getTodayDate() {
  var date = new Date().toISOString().slice(0, 10);

  return date.split('-').reverse().join('-');
}

function getTimesheetCols(values) {
  var columns = {};

  for (var i = 0; i < values[1].length; i++) {
    var val = values[1][i];
    var isDate = typeof val.getMonth === 'function';

    if (isDate) {
      val = getFormattedDate(val);
      var colData = {timeIn: i-2, timeOut: i-1};
      columns[val] = colData;
    }
  }
Logger.log('timesheet Columns (before)', columns);
  var dateArr = getTodayDate().split('-');
  var curMonth = OFFICIAL_SHEET_NAME.split('.')[0];
  var curYear = dateArr[2];

  var lastDate = '23-' + curMonth + '-' + curYear;
  var prevLastDate = '22-' + curMonth + '-' + curYear;
  if (!columns[lastDate]) {
    var t = columns[prevLastDate];
    columns[lastDate] = {timeIn: t.timeIn +3, timeOut: t.timeOut + 3 }
  }

  return columns;
}

function getFormattedDate(date) {
  var str = date.toISOString().slice(0, 10);

  return str.split('-').reverse().join('-');
}

// GET current checkin month
// sample: '08.2021'
// current month: 24th month1 ~ 23rd month2
function getCurrentCheckinMonth() {
  var d = new Date();
  var curYear = d.getFullYear();
  var curMonth = d.getMonth() + 1;
  var curDate = d.getDate();

  if (curDate > 23 && curMonth < 12) curMonth++;

  if (curMonth == 12 && curDate > 23) {
    curMonth = 1;
    curYear++;
  }

  var str = String("0" + curMonth).slice(-2) + '.' + curYear;

  return str;
}

// Get project list from local sheet
function getProjectList() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(LOCAL_PROJECT_LIST_SHEET_NAME);
  
  var dataRange = sheet.getDataRange();
  var values = dataRange.getValues();
  
  var projects = [];
  
  for (var i = 0; i < values.length; i++) {
    projects.push(values[i][0]);
  }
  
  return projects;
}