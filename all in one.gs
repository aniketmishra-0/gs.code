function sendAbsenteeAlert() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Daily Attendance");
  if (!sheet) {
    Logger.log("Sheet 'Daily Attendance' not found.");
    return;
  }

  var headerRow = 2;
  var attendanceStartCol = 5; // Column E=5
  var numAttendanceCols = sheet.getRange("E2:JC2").getNumColumns();
  var headers = sheet.getRange(headerRow, attendanceStartCol, 1, numAttendanceCols).getValues()[0];

  var firstDataRow = 5;
  var lastDataRow = sheet.getLastRow();

  var attendanceData = sheet.getRange(firstDataRow, attendanceStartCol, lastDataRow - firstDataRow + 1, numAttendanceCols).getValues();
  var names = sheet.getRange(firstDataRow, 4, lastDataRow - firstDataRow + 1, 1).getValues(); // Column D=4
  var grades = sheet.getRange(firstDataRow, 2, lastDataRow - firstDataRow + 1, 1).getValues(); // Column B=2

  var today = new Date();
  today.setHours(0, 0, 0, 0);

  var dateCols = [];
  var dateValues = [];

  for (var i = headers.length - 1; i >= 0; i--) {
    var hd = headers[i];
    var parsedDate = null;

    if (hd instanceof Date) {
      parsedDate = new Date(hd.getFullYear(), hd.getMonth(), hd.getDate());
    } else if (typeof hd === "string" && hd.trim() !== "") {
      parsedDate = tryParseDateString(hd.trim(), today.getFullYear());
    }

    if (parsedDate && !isNaN(parsedDate.getTime())) {
      var diffDays = (today - parsedDate) / (1000 * 3600 * 24);
      if (diffDays >= 0 && diffDays <= 2) {
        dateCols.unshift(i);
        dateValues.unshift(parsedDate);
        if (dateCols.length === 3) break;
      }
    }
  }

  if (dateCols.length < 3) {
    Logger.log("Not enough recent date columns found for the check.");
    return;
  }

  var absentees = [];
  for (var r = 0; r < attendanceData.length; r++) {
    var studentName = names[r][0];
    var grade = grades[r][0];
    if (!studentName || !grade) continue;

    var absentAll3 = true;
    var absentDates = [];

    for (var c = 0; c < 3; c++) {
      var status = attendanceData[r][dateCols[c]];
      if (String(status).toUpperCase() === "A") {
        absentDates.push(formatDateSimple(dateValues[c]));
      } else {
        absentAll3 = false;
        break;
      }
    }
    if (absentAll3) {
      absentees.push({name: studentName, dates: absentDates.join(", "), grade: grade});
    }
  }

  var recipient = "aniket.mishra@pw.live"; // Change to your email address
  var subject = "Absentee Alert: 3 Consecutive Days";

  if (absentees.length === 0) {
    var htmlBody = getAllPresentRichTemplate();
    MailApp.sendEmail({
      to: recipient,
      subject: subject,
      htmlBody: htmlBody
    });
    Logger.log("No absentees, sent perfect attendance email.");
    return;
  }

  var rowsHtml = absentees.map(row => 
    "<tr><td>" + row.grade + "</td><td>" + row.name + "</td><td>" + row.dates + "</td></tr>"
  ).join("");

  var htmlBody = `
<!DOCTYPE html>
<html>
<head>
  <style>
    body {
      font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
      background-color: #f9fafb;
      margin: 0;
      padding: 25px;
    }
    .wrapper {
      max-width: 700px;
      margin: auto;
      background: #ffffff;
      border-radius: 12px;
      box-shadow: 0 6px 22px rgba(0, 0, 0, 0.08);
      padding: 24px 32px;
    }
    h2 {
      text-align: center;
      color: #30a14e;
      margin-bottom: 24px;
      font-weight: 700;
    }
    table {
      width: 100%;
      border-collapse: collapse;
    }
    thead tr {
      background: #30a14e;
      color: #ffffff;
      font-weight: 600;
    }
    th, td {
      padding: 14px 16px;
      border: 1px solid #e3e6ee;
      text-align: left;
      font-size: 15px;
    }
    tbody tr:nth-child(odd) {
      background-color: #f5fafd;
    }
    tbody tr:hover {
      background-color: #d4f1dc;
      transition: background-color 0.3s ease;
    }
    .footer {
      margin-top: 32px;
      font-size: 13px;
      color: #888;
      text-align: center;
    }
  </style>
</head>
<body>
  <div class="wrapper">
    <h2>Students Absent for 3 Consecutive Days</h2>
    <table>
      <thead>
        <tr>
          <th>Grade</th>
          <th>Student Name</th>
          <th>Absent Dates</th>
        </tr>
      </thead>
      <tbody>
        ${rowsHtml}
      </tbody>
    </table>
  </div>
</body>
</html>`;

  MailApp.sendEmail({
    to: recipient,
    subject: subject,
    htmlBody: htmlBody
  });

  Logger.log("Absentee alert email sent to " + recipient);
}

function formatDateSimple(dateObj) {
  return Utilities.formatDate(dateObj, Session.getScriptTimeZone(), "dd-MMM-yyyy");
}

function tryParseDateString(str, year) {
  var months = {
    "Jan":0,"Feb":1,"Mar":2,"Apr":3,"May":4,"Jun":5,"Jul":6,
    "Aug":7,"Sep":8,"Oct":9,"Nov":10,"Dec":11,"Sept":8
  };
  var dmy = str.match(/^(\d{1,2})[- ]([A-Za-z]{3,4})$/);
  var mdy = str.match(/^([A-Za-z]{3,4})[- ](\d{1,2})$/);
  if (dmy) {
    var day = parseInt(dmy[1], 10);
    var mon = months[dmy[2].substring(0,3)];
    if (!isNaN(day) && mon !== undefined) return new Date(year, mon, day);
  } else if (mdy) {
    var mon = months[mdy[1].substring(0,3)];
    var day = parseInt(mdy[2], 10);
    if (!isNaN(day) && mon !== undefined) return new Date(year, mon, day);
  }
  return null;
}

function getAllPresentRichTemplate() {
  return `
<!DOCTYPE html>
<html>
<head>
  <meta charset="UTF-8" />
  <style>
    body {
      font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
      background-color: #f9fafb;
      margin: 0;
      padding: 40px 20px;
      color: #2e7d32;
      text-align: center;
    }
    .container {
      background: #e8f5e9;
      border-radius: 15px;
      max-width: 480px;
      margin: auto;
      padding: 30px 20px;
      box-shadow: 0 5px 18px rgba(30, 110, 50, 0.15);
      border: 2px solid #a5d6a7;
    }
    .check-icon {
      width: 64px;
      height: 64px;
      margin: 0 auto 20px auto;
      background-color: #4caf50;
      border-radius: 50%;
      display: flex;
      align-items: center;
      justify-content: center;
      box-shadow: 0 3px 12px rgba(60, 140, 60, 0.15);
    }
    .check-icon svg {
      fill: white;
      width: 35px;
      height: 35px;
    }
    h1 {
      font-weight: 700;
      font-size: 1.7em;
      margin-bottom: 10px;
      letter-spacing: 0.03em;
    }
    p {
      font-size: 1.1em;
      margin-top: 0;
      color: #387c3c;
    }
  </style>
</head>
<body>
  <div class="container">
    <div class="check-icon" aria-hidden="true">
      <svg viewBox="0 0 24 24">
        <path d="M9.7 16.3l-4-4c-.39-.39-1.02-.39-1.41 0-.39.39-.39 1.02 0 1.41l4.7 4.7c.39.39 1.02.39 1.41 0l9-9c.39-.39.39-1.02 0-1.41-.39-.39-1.02-.39-1.41 0l-8.3 8.3z"></path>
      </svg>
    </div>
    <h1>Perfect Attendance!</h1>
    <p>All students are present for the last <strong>3 consecutive days</strong>.</p>
  </div>
</body>
</html>
`;
}
