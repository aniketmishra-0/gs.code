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
  var names = sheet.getRange(firstDataRow, 4, lastDataRow - firstDataRow + 1, 1).getValues(); // D column = 4
  var grades = sheet.getRange(firstDataRow, 2, lastDataRow - firstDataRow + 1, 1).getValues(); // B column = 2

  var today = new Date();
  today.setHours(0, 0, 0, 0);

  var dateCols = [];
  var dateValues = [];

  // Find last 3 dates columns dynamically
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
    Logger.log("Not enough recent date columns found for check.");
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

  var recipient = "aniket.mishra@pw.live";  // Put your email here
  var subject = "Absentee Alert: 3 Consecutive Days";

  if (absentees.length === 0) {
    var htmlBody = "<p>All students are present for the last 3 consecutive days.</p>";
    MailApp.sendEmail({
      to: recipient,
      subject: subject,
      htmlBody: htmlBody
    });
    Logger.log("No absentees, 'all present' email sent.");
    return;
  }

  var rowsHtml = absentees.map(row => 
    "<tr><td>" + row.grade + "</td><td>" + row.name + "</td><td>" + row.dates + "</td></tr>"
  ).join("");

  var template = HtmlService.createTemplateFromFile('AbsenteeEmailTemplate');
  template.rows = rowsHtml;
  var htmlBody = template.evaluate().getContent();

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
  var months = {"Jan":0,"Feb":1,"Mar":2,"Apr":3,"May":4,"Jun":5,"Jul":6,"Aug":7,"Sep":8,"Oct":9,"Nov":10,"Dec":11,"Sept":8};
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
