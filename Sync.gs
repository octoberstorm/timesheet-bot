/**
 * Sync local timesheet data to official sheet
 * Author: octoberstorm <storm.october@gmail.com>
 */

const PROJECTS = ['ProjectCool', 'ProjectHot'];
const LOCAL_SHEET_URL = 'https://docs.google.com/spreadsheets/d/<SHEET_ID>/edit';
const OFFICIAL_SHEET_URL = 'https://docs.google.com/spreadsheets/d/<SHEET_ID>/edit';
const OFFICIAL_SHEET_NAME = 'TS_April';
const LOCAL_MEMBER_MAPPING_SHEET_NAME = 'MemberMapping';

// TODO auto get cols
const TIMESHEET_COLUMNS = {
  '10-03-2020': {in: 'O', out: 'P'},
  '11-03-2020': {in: 'R', out: 'S'},
  '12-03-2020': {in: 'U', out: 'V'},
  '13-03-2020': {in: 'X', out: 'Y'},
  '14-03-2020': {in: 'AA', out: 'AB'},
  '15-03-2020': {in: 'AD', out: 'AE'},
  '16-03-2020': {in: 'AG', out: 'AH'},
  '17-03-2020': {in: 'AJ', out: 'AK'},
  '18-03-2020': {in: 'AM', out: 'AN'},
  '19-03-2020': {in: 'AP', out: 'AQ'},
  '20-03-2020': {in: 'AS', out: 'AT'},
  '23-03-2020': {in: 'BB', out: 'BC'},
  '24-03-2020': {in: 'BE', out: 'BF'},
  '25-03-2020': {in: 'BH', out: 'BI'},
  '26-03-2020': {in: 'J', out: 'K'},
  '27-03-2020': {in: 'M', out: 'N'},
  '28-03-2020': {in: 'P', out: 'Q'},
  '29-03-2020': {in: 'S', out: 'T'},
  '30-03-2020': {in: 'V', out: 'W'},
  '31-03-2020': {in: 'Y', out: 'Z'},
  '01-04-2020': {in: 'AB', out: 'AC'},
  '02-04-2020': {in: 'AE', out: 'AF'},
  '03-04-2020': {in: 'AH', out: 'AI'},
  '04-04-2020': {in: 'AK', out: 'AL'},
  '05-04-2020': {in: 'AN', out: 'AO'},
  '06-04-2020': {in: 'AQ', out: 'AR'},
  '07-04-2020': {in: 'AT', out: 'AU'},
  '08-04-2020': {in: 'AW', out: 'AX'},
  '09-04-2020': {in: 'AZ', out: 'BA'},
  '10-04-2020': {in: 'BC', out: 'BD'},
  '11-04-2020': {in: 'BF', out: 'BG'},
  '12-04-2020': {in: 'BI', out: 'BJ'},
  '13-04-2020': {in: 'BL', out: 'BM'},
  '14-04-2020': {in: 'BO', out: 'BP'},
  '15-04-2020': {in: 'BR', out: 'BS'},
  '16-04-2020': {in: 'BQ', out: 'BP'},
  '17-04-2020': {in: 'BU', out: 'BV'},
  '18-04-2020': {in: 'BX', out: 'BY'},
  '19-04-2020': {in: 'CA', out: 'CB'},
  '20-04-2020': {in: 'CD', out: 'CE'},
};

// main function
function syncTimesheet() {
  var localSS = SpreadsheetApp.openByUrl(LOCAL_SHEET_URL);
  var officialSS = SpreadsheetApp.openByUrl(OFFICIAL_SHEET_URL);
  
  // get local data
  var project = PROJECTS[0];
  SpreadsheetApp.setActiveSpreadsheet(localSS);

  var nameList = getNameList();
  
  var date = getTodayDate();
  // var date = '26-03-2020';
  
  var localData = findLocalData(nameList, date);
  
  Logger.log('localData:', localData);
  
  // Now sending data to official sheet
  SpreadsheetApp.setActiveSpreadsheet(officialSS);      
  var officialSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(OFFICIAL_SHEET_NAME);  
    
  var dataRange = officialSheet.getDataRange();
  var values = dataRange.getValues();
  
  var updatedCount = 0;
  
  for (var i in localData) {
    var item = localData[i];
    
    var name = item.name;
    
    if (name == null) {
      continue;
    }
    
    var timeIn = item.in;
    var timeOut = item.out;
    
    var row = findTimesheetRowNumber(values, name);
    
    if (row == -1) {
      Logger.log('not found member in official sheet: ' + name);
      continue;
    }
    
    var cols = findTimesheetColums(date);
    officialSheet.getRange(cols.in + row).setValue(timeIn);
    officialSheet.getRange(cols.out + row).setValue(timeOut);
    
    updatedCount++;
  }
  
  Logger.log('Updated ' + updatedCount + ' member timesheets for ' + date);
}

function findLocalData(nameList, date) {
  var data = [];

  for (var i in PROJECTS) {
    var projectName = PROJECTS[i];
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
  return TIMESHEET_COLUMNS[date];
}

function getTodayDate() {
  var date = new Date().toISOString().slice(0, 10);

  return date.split('-').reverse().join('-');
}