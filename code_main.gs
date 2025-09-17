const CONFIG = {
  sheetName: "Daily Attendance",
  headerRow: 2,
  attendanceStartCol: 5, // Column E = 5
  attendanceDateRange: "E2:JC2",
  firstDataRow: 5,
  nameCol: 4,   // Column D
  gradeCol: 2,  // Column B
  absentMarker: "A",
  absentDaysCount: 3,
  recipientEmails: "aniket.mishra@pw.live", // Comma separated emails
  emailSubject: "Absentee Alert: 3 Consecutive Days"
};

function sendAbsenteeAlert() {
  Logger.log("Starting absentee alert processing.");

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.sheetName);
  if (!sheet) {
    Logger.log(`Sheet '${CONFIG.sheetName}' not found. Exiting script.`);
    return;
  }
  Logger.log(`Sheet '${CONFIG.sheetName}' found.`);

  var numAttendanceCols = sheet.getRange(CONFIG.attendanceDateRange).getNumColumns();
  Logger.log(`Found ${numAttendanceCols} attendance date columns.`);

  var lastDataRow = sheet.getLastRow();
  var headers = sheet.getRange(CONFIG.headerRow, CONFIG.attendanceStartCol, 1, numAttendanceCols).getValues()[0];

  var dateCols = [];
  var dateValues = [];
  var today = new Date();
  today.setHours(0, 0, 0, 0);

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
      if (diffDays >= 0 && diffDays < CONFIG.absentDaysCount) {
        dateCols.unshift(i);
        dateValues.unshift(parsedDate);
        if (dateCols.length === CONFIG.absentDaysCount) break;
      }
    }
  }
  Logger.log(`Using date columns: ${dateCols} for absence check.`);
  Logger.log(`Corresponding dates: ${dateValues.map(d => d.toDateString()).join(", ")}`);

  if (dateCols.length < CONFIG.absentDaysCount) {
    Logger.log("Not enough recent date columns found for the check. Script stopped.");
    return;
  }

  var attendanceData = sheet.getRange(CONFIG.firstDataRow, CONFIG.attendanceStartCol, lastDataRow - CONFIG.firstDataRow + 1, numAttendanceCols).getValues();
  var names = sheet.getRange(CONFIG.firstDataRow, CONFIG.nameCol, lastDataRow - CONFIG.firstDataRow + 1, 1).getValues();
  var grades = sheet.getRange(CONFIG.firstDataRow, CONFIG.gradeCol, lastDataRow - CONFIG.firstDataRow + 1, 1).getValues();

  Logger.log(`Processing attendance data for ${names.length} students.`);

  var absentees = [];
  var totalValidStudents = 0;

  for (var r = 0; r < attendanceData.length; r++) {
    var studentName = names[r][0];
    var grade = grades[r][0];

    if (!studentName || !studentName.toString().trim()) continue;
    if (!grade) continue;

    // Clean name: only letters and spaces
    var cleanedName = studentName.toString().replace(/[^A-Za-z ]+/g, "").trim();
    if (!cleanedName) continue;

    totalValidStudents++;

    var absentAllDays = true;
    var absentDates = [];

    for (var c = 0; c < dateCols.length; c++) {
      var status = attendanceData[r][dateCols[c]];
      if (String(status).toUpperCase() === CONFIG.absentMarker) {
        absentDates.push(formatDateSimple(dateValues[c]));
      } else {
        absentAllDays = false;
        break;
      }
    }

    if (absentAllDays) absentees.push({name: cleanedName, dates: absentDates.join(", "), grade: grade});
  }

  Logger.log(`Total valid students counted: ${totalValidStudents}`);
  Logger.log(`Total absentees found: ${absentees.length}`);
  if(absentees.length){
    Logger.log(`Absentees: ${absentees.map(a => a.name).join(", ")}`);
  }

  var recipient = CONFIG.recipientEmails;
  var subject = CONFIG.emailSubject;

  if (absentees.length === 0) {
    Logger.log("No absentees found, sending 'all present' email.");
    var htmlBody = getAllPresentRichTemplate(CONFIG.absentDaysCount);
    MailApp.sendEmail({to: recipient, subject: subject, htmlBody: htmlBody});
    Logger.log("Perfect attendance email sent.");
    return;
  }

  var rowsHtml = absentees.map(row => 
    `<tr><td>${row.grade}</td><td>${row.name}</td><td>${row.dates}</td></tr>`
  ).join("");

  var htmlBody = generateAbsenteesHtml(
    rowsHtml,
    CONFIG.absentDaysCount,
    totalValidStudents,
    absentees.length
  );

  MailApp.sendEmail({to: recipient, subject: subject, htmlBody: htmlBody});
  Logger.log("Absentee alert email sent successfully.");
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

function generateAbsenteesHtml(rowsHtml, daysCount, totalStudents, totalAbsentees) {
  return `
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
      margin-bottom: 16px;
      font-weight: 700;
    }
    .summary {
      font-size: 15px;
      margin-bottom: 16px;
      color: #222;
      text-align: center;
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
    <h2>Students Absent for ${daysCount} Consecutive Days</h2>
    <div class="summary">
      <b>Total Students:</b> ${totalStudents}
      &nbsp;|&nbsp;
      <b>Total Absentees:</b> ${totalAbsentees}
    </div>
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
}

function getAllPresentRichTemplate(daysCount) {
  return `
<!DOCTYPE html>
<html>
<head>
  <meta charset="UTF-8" />
  <style>
    body {
      font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
      background-color: #f7fafc;
      margin: 40px auto;
      padding: 0 20px;
      color: #2e7d32;
      text-align: center;
    }
    .alert-container {
      background: #ffffff;
      max-width: 480px;
      margin: 0 auto;
      padding: 30px 20px;
      border-radius: 8px;
      box-shadow: 0 4px 18px rgba(30,100,70,0.07), 0 1.5px 5px rgba(70,70,70,0.09);
    }
    .icon-circle {
      background: #4caf50;
      width: 64px;
      height: 64px;
      margin: 0 auto 20px auto;
      border-radius: 50%;
      display: flex;
      align-items: center;
      justify-content: center;
      box-shadow: 0 2px 8px rgba(60,140,60,0.10);
    }
    .icon-circle svg {
      width: 32px;
      height: 32px;
      fill: #fff;
    }
    .alert-title {
      font-size: 1.40em;
      font-weight: 600;
      color: #2e7d32;
      margin-bottom: 12px;
      margin-top: 10px;
      letter-spacing: 0.01em;
    }
    .alert-message {
      font-size: 1.06em;
      margin-bottom: 0;
    }
  </style>
</head>
<body>
  <div class="alert-container">
    <div class="icon-circle" aria-hidden="true">
      <svg viewBox="0 0 24 24"><path d="M9.7 16.3l-4-4c-.39-.39-1.02-.39-1.41 0-.39.39-.39 1.02 0 1.41l4.7 4.7c.39.39 1.02.39 1.41 0l9-9c.39-.39.39-1.02 0-1.41-.39-.39-1.02-.39-1.41 0l-8.3 8.3z"/></svg>
    </div>
    <div class="alert-title">Perfect Attendance!</div>
    <div class="alert-message">
      All students are present for the last <b>${daysCount} consecutive days</b>.
    </div>
  </div>
</body>
</html>
`;
}
