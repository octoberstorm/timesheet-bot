const { google } = require('googleapis');
const { promisify } = require('util');

// id of spreadsheet file
const sheetID = process.env.TIMESHEET_FILE_ID;

const findTimesheetRow = (sheetValues, date, name) => {
    for (var i = 0; i < sheetValues.length; i++) {
      if (date.includes(sheetValues[i][0])) {
        if (sheetValues[i][1] == name) {
          return i + 1;
        }
      }
    }
    return -1;
  }

const getDateTime = () => {
  var t = generateDatetime();

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

    return {date, time};
}

const addNewRow = (sheetsAPI, range, values) => {
  const addRow = promisify(sheetsAPI.spreadsheets.values.append.bind(sheetsAPI.spreadsheets));

  addRow({
      spreadsheetId: sheetID,
      range,
      valueInputOption: 'USER_ENTERED',
      insertDataOption: 'INSERT_ROWS',
      resource: {
        "majorDimension": "ROWS",
        "values": values
      }
    }).then(resp => console.log('resp', resp))
    .catch(err => console.log('err:', err));
}

  const updateRow = (sheetsAPI, range, values) => {
    const update = promisify(sheetsAPI.spreadsheets.values.update.bind(sheetsAPI.spreadsheets));

    update({
      spreadsheetId: sheetID,
      range,
      valueInputOption: 'USER_ENTERED',
      resource: {
        values
      },
    }).then(resp => console.log('resp', resp))
    .catch(err => console.log('err:', err));
}

exports.checkin = async (req, res) => {
  const auth = await google.auth.getClient({
    scopes: [
      "https://www.googleapis.com/auth/spreadsheets"
    ]
  });

  const {date, time} = getDateTime();
  const {name, project, op} = req.body;

  let values = [];
  if (op == 'checkin') {
    values = [[date, name, time, '', project]];
  } else { // checkout
    values = [[time]];
  }

  const sheetsAPI = google.sheets({version: 'v4', auth});

  const findRow = promisify(sheetsAPI.spreadsheets.values.get.bind(sheetsAPI.spreadsheets));

  findRow({
      spreadsheetId: sheetID,
      range: project + '!A:F',
    }).then((resp) => {
      const sheetValues = resp.data.values;
      const row = findTimesheetRow(sheetValues, date, name);

      let range = '';
      if (op == 'checkin') {
        range = `${project}!A:C`;
      } else { // checkout
        range = `${project}!D${row}:D${row}`;
      }

      if (row == -1) {
        addNewRow(sheetsAPI, range, values);
      } else {
        // update existing record
        if (op == 'checkout') {
          updateRow(sheetsAPI, range, values);
        }
      }

    }).catch(err => {
      console.log(err);
    });
  
    const ok = 'ok';
    res.status(200).send({ ok });
}

const generateDatetime = () => {
  let date = new Date(); // 2018-07-24:17:26:00 (Look like GMT+0)
  const myTimeZone = 7; // my timeZone 

  date.setTime( date.getTime() + myTimeZone * 60 * 60 * 1000 );
  return date;
}