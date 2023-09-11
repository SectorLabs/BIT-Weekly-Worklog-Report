///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
function getMonday(d) {
  d = new Date(d);
  var day = d.getDay(),
      diff = d.getDate() - day + (day == 0 ? -6 : 1); // adjust when day is Sunday
  return new Date(d.setDate(diff));
}
function BITPK_weekclickuptime() {
  var apiKeys = [
    'pk_***************************************',
    'pk_***************************************',
    'pk_***************************************',
    'pk_***************************************'
  ];

  String.prototype.addQuery = function (obj) {
    return this + Object.keys(obj).reduce(function (p, e, i) {
      return p + (i == 0 ? "?" : "&") +
        (Array.isArray(obj[e]) ? obj[e].reduce(function (str, f, j) {
          return str + e + "=" + encodeURIComponent(f) + (j != obj[e].length - 1 ? "&" : "")
        }, "") : e + "=" + encodeURIComponent(obj[e]));
    }, "");
  };

  var start_date = getMonday(new Date());
  start_date.setHours(9, 0, 0, 0);

  var end_date = new Date(start_date); // Create a new Date object based on start_date
  end_date.setDate(end_date.getDate() + 4); // Add 4 days to the start_date (Monday) to get Friday
  end_date.setHours(23, 0, 0, 0);

  // Convert start_date and end_date to timestamps
  var start_timestamp = start_date.getTime();
  var end_timestamp = end_date.getTime();

  const query = Object.entries({
    start_date: start_date,
    end_date: end_date,
    assignee: '0',
    include_task_tags: 'true',
    include_location_names: 'true',
    space_id: '0',
    folder_id: '0',
    list_id: '0',
    task_id: '0',
    custom_task_ids: 'true',
    team_id: '3692463',
  }).map(row => row.join('='))
    .join('&');

  const teamId = '3692463';
  const baseUrl = `https://api.clickup.com/api/v2/team/${teamId}/time_entries`;
  const url = baseUrl.addQuery(query);

  var allDataArray = [];

  for (var k = 0; k < apiKeys.length; k++) {
    const response = UrlFetchApp.fetch(url,
      {
        method: 'GET',
        headers: {
          'Content-Type': 'application/json',
          Authorization: apiKeys[k]
        }
      }
    );
    try {
      var data = JSON.parse(response.getContentText());
      var dataArray = [];
      var dataLength = data.data.length;
      for (var i = 0; i < dataLength; i++) {
        var timeData = data.data[i];
        var start = timeData.start;
        var end = timeData.end;
        var task = timeData.task.name;
        var user = timeData.user.username;
        var duration = timeData.duration;
        var hours = Math.floor(duration / (1000 * 60 * 60));
        var minutes = Math.floor((duration % (1000 * 60 * 60)) / (1000 * 60));
        var task_url = timeData.task_url;

        if (start >= start_date && end <= end_date) {
          dataArray.push([task, hours, minutes, timeData.id, timeData.user.username]);
        }
      }
      allDataArray = allDataArray.concat(dataArray);
    } catch (e) {
      Logger.log(e);
    }
  }

  var sheet = SpreadsheetApp.getActive().getSheetByName('WeeklyTeamClickUp');
  sheet.clear();
  var numRows = allDataArray.length;
  if (numRows > 0) {
    sheet.getRange(1, 1, numRows, 5).setValues(allDataArray);
} else {
Logger.log("No data found for today");
}
}

//////////////////////////////////////////////////////////////////////////////////////////////////////////////
function getLastMonday() {
  const now = new Date();
  const dayOfWeek = now.getDay();
  const lastMonday = new Date(now.getFullYear(), now.getMonth(), now.getDate(), 0, 0, 0);

  if (dayOfWeek === 5) { // Friday
    lastMonday.setDate(lastMonday.getDate() - 4); // Subtract 4 days from Friday to get last Monday
  } else {
    const diff = lastMonday.getDate() - dayOfWeek + (dayOfWeek === 0 ? -6 : 1); // adjust when day is Sunday
    lastMonday.setDate(diff);
  }

  return lastMonday;
}
///////////////////////////////////
function BIT_WeeklySDPRequests() {
  const sdpaapo = SpreadsheetApp.getActiveSpreadsheet();
  const apisheet = SpreadsheetApp.openById("1h9MATYDtYtsMQgziMuffiC3W6sC8ebkcuuxGyopgDQA").getSheetByName('MAIN').getDataRange().getValues();

  const url = "https://helpdesk.empgservices.com/app/itdesk/api/v3/requests";

  const lastMonday = getLastMonday();
  const lastMondayStartTime = lastMonday.getTime();

  const params = {
  "list_info": {
    "fields_required": [
      "requester",
      "subject",
      "technician",
      "site"
    ],
    "search_criteria": [
      {
        "field": "last_updated_time",
        "condition": "greater than",
        "values": [
          {
            "value": lastMondayStartTime,
            "display_value": "Today"
          }
        ],
        "logical_operator": "AND"
      },
      {
        "field": "technician.email_id",
        "condition": "is",
        "values": [
          "shabbir.haider@dubizzle.com",
          "ayub.sarfraz@dubizzle.com",
          "zain.rauf@dubizzle.com",
          "helpdesk@empgservices.com"
        ],
        "logical_operator": "AND"
      }
    ],
    "row_count": 100
  }
};

  const paramsJSON = JSON.stringify(params);
  const encodedpURL = encodeURIComponent(paramsJSON);

  const options = {
    "method": "get",
    "headers": {
      "contentType": "application/x-www-form-urlencoded", // <--  HERE
      "Accept": "application/vnd.manageengine.sdp.v3+json",
      "Authorization": "Zoho-oauthtoken " + apisheet[1][0]
    },
    "muteHttpExceptions": true
  };

  const response = UrlFetchApp.fetch(`${url}?input_data=${encodedpURL}`, options);
  const data = JSON.parse(response.getContentText()).requests;

  // Write the data to the sheet
  const sheet = sdpaapo.getSheetByName("Weekly SDP");
  sheet.clearContents(); // clear existing data
  if (data.length > 0) {
  sheet.getRange(1, 1, data.length, 4).setValues(data.map(request => [request.requester.name, request.subject, request.technician.email_id, request.id]));
}
}
//////////////////////////////////////////////////////////////////////////////////
function BIT_WeeklyWorklog() {
  const sdpaapo = SpreadsheetApp.getActiveSpreadsheet();
  const apisheet = SpreadsheetApp.openById("1h9MATYDtYtsMQgziMuffiC3W6sC8ebkcuuxGyopgDQA").getSheetByName('MAIN').getDataRange().getValues();

  const baseUrl = "https://helpdesk.empgservices.com/app/itdesk/api/v3/requests";
  const inputSheet = sdpaapo.getSheetByName("Weekly SDP");
  const outputSheet = sdpaapo.getSheetByName("WeeklyTeamWorklog");
  outputSheet.clearContents();

  const lastMonday = getLastMonday();
  const lastMondayStartTime = lastMonday.getTime();

  const lastRow = inputSheet.getLastRow();

  const options = {
    "method": "get",
    "headers": {
      "contentType": "application/x-www-form-urlencoded",
      "Accept": "application/vnd.manageengine.sdp.v3+json",
      "Authorization": "Zoho-oauthtoken " + apisheet[1][0]
    },
    "muteHttpExceptions": true
  };

  for (let row = 1; row <= lastRow; row++) {
    const requestId = inputSheet.getRange(row, 4).getValue();
    if (!requestId) {
      continue; // Skip rows without a request ID
    }

    const url = `${baseUrl}/${requestId}/worklogs`;

    const params = {
      "list_info": {
        "fields_required": [
          "id",
          "description",
          "created_by",
          "technician",
          "time_spent",
          "recorded_time",
          "start_time"
        ],
        "row_count": 100,
        "search_criteria": [
          {
            "field": "start_time",
            "condition": "greater than",
            "values": [
              {
                "value": lastMondayStartTime
              }
            ],
            "logical_operator": "AND"
          }
        ]
      }
    };
    const paramsJSON = JSON.stringify(params);
    const encodedpURL = encodeURIComponent(paramsJSON);
    const response = UrlFetchApp.fetch(`${url}?input_data=${encodedpURL}`, options);
    const responseContentText = response.getContentText();

    let data;
    try {
      data = JSON.parse(responseContentText).worklogs;
    } catch (e) {
      console.error("Error parsing JSON data:", e);
    }

    if (data && data.length > 0) {
      const mappedData = data.map(request => {
        const {hours, minutes} = request.time_spent;
        return [request.description,hours, minutes, request.id,request.created_by ? request.created_by.email_id : ""];
      });
      outputSheet.getRange(outputSheet.getLastRow() + 1, 1, mappedData.length, 5).setValues(mappedData);
    }
  }
}

/////////////////////////Weekly Team Report///////////////////////////////////////////////////////////////
function wekBusyDurations(calendarIds) {
  var today = new Date();
  var dayOfWeek = today.getDay();
  var daysSinceMonday = (dayOfWeek + 6) % 7;
  var startDate = new Date(today.getFullYear(), today.getMonth(), today.getDate() - daysSinceMonday);
  var daysUntilFriday = (5 - dayOfWeek + 7) % 7;
  var endDate = new Date(today.getFullYear(), today.getMonth(), today.getDate() + daysUntilFriday);

  var busyDurations = {};
  for (var i = 0; i < calendarIds.length; i++) {
    var calendar = CalendarApp.getCalendarById(calendarIds[i]);
    var events = calendar.getEvents(startDate, endDate, { 'showDeleted': false, 'maxResults': 1000, 'singleEvents': true, 'orderBy': 'startTime' });
    var busyDuration = 0;
    for (var j = 0; j < events.length; j++) {
      if (events[j].isAllDayEvent()) continue;
      var start = events[j].getStartTime();
      var end = events[j].getEndTime();
      var duration = end.getTime() - start.getTime();
      busyDuration += duration;
    }
    var totalHours = Math.floor(busyDuration / (1000 * 60 * 60));
    var totalMinutes = Math.floor(busyDuration % (1000 * 60 * 60) / (1000 * 60));
    busyDurations[calendarIds[i]] = { hours: totalHours, minutes: totalMinutes };
  }

  return busyDurations;
}

function sendWeeklyReport() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const worklogSheet = ss.getSheetByName("WeeklyTeamWorklog");
  const clickUpSheet = ss.getSheetByName("WeeklyTeamClickUp");
  
  const userEmails = {
    'Omer Zia': {email: 'helpdesk@empgservices.com', calendarId: 'omer.zia@dubizzle.com'},
    'Shabbir Haider': {email: 'shabbir.haider@dubizzle.com', calendarId: 'shabbir.haider@dubizzle.com'},
    'Ayub Sarfraz': {email: 'ayub.sarfraz@dubizzle.com', calendarId: 'ayub.sarfraz@dubizzle.com'},
    'Zain Rauf': {email: 'zain.rauf@dubizzle.com', calendarId: 'zain.rauf@dubizzle.com'}
  };

  const busyDurations = wekBusyDurations(Object.values(userEmails).map(user => user.calendarId));

  const tables = [];

  for (const [name, user] of Object.entries(userEmails)) {
    const table = WeeklyTable(name, user.email, user.calendarId, busyDurations, worklogSheet, clickUpSheet);
    tables.push(table);
  }

  WeeklyEmail(tables);
}

function WeeklyTable(name, email, calendarId, busyDurations, worklogSheet, clickUpSheet) {
  let table = `<h3>${name}</h3><table border="1" cellpadding="5" cellspacing="0"><tr><th>Source</th><th>Tasks/Worklog</th><th>Time Spend in Hours</th><th>Time Spend in Mins</th></tr>`;

  const worklogData = worklogSheet.getDataRange().getValues();
  const clickUpData = clickUpSheet.getDataRange().getValues();

  const worklogFilteredData = worklogData.filter(row => row[4] === email).map(row => ({ source: 'Worklog', data: row }));
  const clickUpFilteredData = clickUpData.filter(row => row[4] === name).map(row => ({ source: 'ClickUp', data: row }));

  const filteredData = [...worklogFilteredData, ...clickUpFilteredData];

  let totalHours = 0;
  let totalMinutes = 0;

  for (const rowObj of filteredData) {
    const row = rowObj.data;
    table += `<tr><td>${rowObj.source}</td><td>${row[0]}</td><td>${row[1]}</td><td>${row[2]}</td></tr>`;
    totalHours += row[1];
    totalMinutes += row[2];
  }

  // Calculate total time spent in hours and minutes
  totalHours += Math.floor(totalMinutes / 60);
  totalMinutes = totalMinutes % 60;

  table += `<tr><td>&#128193;</td><td><strong>Total SDP+ ClickUp Time Spent</strong></td><td><strong>${totalHours}</strong></td><td><strong>${totalMinutes}</strong></td></tr>`;
  table += `<tr><td>&#128197;</td><td><strong>Calendar Meeting Durations</strong></td><td><strong>${busyDurations[calendarId].hours}</strong></td><td><strong>${busyDurations[calendarId].minutes}</strong></td></tr>`;
  
  // Add Total SDP+ClickUp Time Spent and Calendar Meeting Durations
  let grandTotalHours = totalHours + busyDurations[calendarId].hours;
  let grandTotalMinutes = totalMinutes + busyDurations[calendarId].minutes;
  grandTotalHours += Math.floor(grandTotalMinutes / 60);
  grandTotalMinutes = grandTotalMinutes % 60;

  table += `<tr><td>&#128336;</td><td><strong>Total Time (${Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'MMMM dd, yyyy')})</strong></td><td><strong>${grandTotalHours}</strong></td><td><strong>${grandTotalMinutes}</strong></td></tr>`;

  table += '</table>';

  return table;
}

function WeeklyTable(name, email, calendarId, busyDurations, worklogSheet, clickUpSheet) {
  const tableStyle = `
    border-collapse: collapse;
    font-family: Arial, sans-serif;
    width: 100%;
    max-width: 600px;
    margin: 0 auto;
  `;

  const thStyle = `
    background-color: #f2f2f2;
    font-weight: bold;
    padding: 5px;
    border: 1px solid #d9d9d9;
    text-align: left;
  `;

  const tdStyle = `
    padding: 5px;
    border: 1px solid #d9d9d9;
    text-align: left;
  `;

  let table = `<h3>${name}</h3><table style="${tableStyle}"><tr><th style="${thStyle}">Source</th><th style="${thStyle}">Tasks/Worklog</th><th style="${thStyle}">Time Spend in Hours</th><th style="${thStyle}">Time Spend in Mins</th></tr>`;

  const worklogData = worklogSheet.getDataRange().getValues();
  const clickUpData = clickUpSheet.getDataRange().getValues();

  const worklogFilteredData = worklogData.filter(row => row[4] === email).map(row => ({ source: 'Worklog', data: row }));
  const clickUpFilteredData = clickUpData.filter(row => row[4] === name).map(row => ({ source: 'ClickUp', data: row }));

  const filteredData = [...worklogFilteredData, ...clickUpFilteredData];

  let totalHours = 0;
  let totalMinutes = 0;

  for (const rowObj of filteredData) {
    const row = rowObj.data;
    table += `<tr><td style="${tdStyle}">${rowObj.source}</td><td style="${tdStyle}">${row[0]}</td><td style="${tdStyle}">${row[1]}</td><td style="${tdStyle}">${row[2]}</td></tr>`;
    totalHours += row[1];
    totalMinutes += row[2];
  }

  // Calculate total time spent in hours and minutes
  totalHours += Math.floor(totalMinutes / 60);
  totalMinutes = totalMinutes % 60;

  table += `<tr><td style="${tdStyle}">&#128193;</td><td style="${tdStyle}"><strong>Total SDP+ ClickUp Time Spent</strong></td><td style="${tdStyle}"><strong>${totalHours}</strong></td><td style="${tdStyle}"><strong>${totalMinutes}</strong></td></tr>`;
  table += `<tr><td style="${tdStyle}">&#128197;</td><td style="${tdStyle}"><strong>Calendar Meeting Durations</strong></td><td style="${tdStyle}"><strong>${busyDurations[calendarId].hours}</strong></td><td style="${tdStyle}"><strong>${busyDurations[calendarId].minutes}</strong></td></tr>`;
  
  // Add Total SDP+ClickUp Time Spent and Calendar Meeting Durations
  let grandTotalHours = totalHours + busyDurations[calendarId].hours;
  let grandTotalMinutes = totalMinutes + busyDurations[calendarId].minutes;
  grandTotalHours += Math.floor(grandTotalMinutes / 60);
  grandTotalMinutes = grandTotalMinutes % 60;

  table += `<tr><td style="${tdStyle}">&#128336;</td><td style="${tdStyle}"><strong>Total Time (${Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'MMMM dd, yyyy')})</strong></td><td style="${tdStyle}"><strong>${grandTotalHours}</strong></td><td style="${tdStyle}"><strong>${grandTotalMinutes}</strong></td></tr>`;
  table += '</table><hr>';
  return table;
}

function WeeklyEmail(tables) {
  const recipient = 'adeel@dubizzle.com,omer.zia@dubizzle.com,shabbir.haider@dubizzle.com,ayub.sarfraz@dubizzle.com,zain.rauf@dubizzle.com'; // Replace with the recipient's email address
  const now = new Date();
  const timeZone = 'Asia/Karachi';
  const dateFormatter = new Intl.DateTimeFormat('en-US', { timeZone, year: 'numeric', month: 'long', day: '2-digit' });
  const formattedDate = dateFormatter.format(now);
  const monthName = now.toLocaleString('default', { month: 'long' });
  const weekNumber = Math.ceil(now.getDate() / 7);
  const subject = `Week ${weekNumber} of ${monthName}, ${now.getFullYear()} - BIT PK Time Spent `;


  let htmlBody = '<html><body>';
  htmlBody += `<h2 style="text-align: center; margin-top: 20px;">${subject}</h2>`;
  for (const table of tables) {
    htmlBody += `<div style="margin-top: 20px;">${table}</div>`;
  }
  htmlBody += '</body></html>';

  const pdfBlob = Utilities.newBlob(htmlBody, 'text/html', 'report.html').getAs('application/pdf').setName(`BIT PK Weekly Report - ${formattedDate}.pdf`);
  const options = {
    htmlBody: htmlBody,
    attachments: [pdfBlob],
    name: 'BIT Automation'
  };

  MailApp.sendEmail(recipient, subject, '', options);
}

function getWeekOfMonth(date) {
  const firstDayOfMonth = new Date(date.getFullYear(), date.getMonth(), 1);
  const dayOfWeek = firstDayOfMonth.getDay();
  const leadingDays = dayOfWeek == 0 ? 6 : dayOfWeek - 1;
  const remainingDays = new Date(date.getFullYear(), date.getMonth() + 1, 0).getDate() - (7 - leadingDays);
  const weekNumber = Math.ceil((date.getDate() - leadingDays) / 7);
  if (weekNumber < 1) {
    return getWeekOfMonth(new Date(date.getFullYear(), date.getMonth(), remainingDays));
  }
  if (weekNumber > 5) {
    return getWeekOfMonth(new Date(date.getFullYear(), date.getMonth() + 1, 1));
  }
  return weekNumber;
}
