const CONFIG = {
  sheetName: "Daily Attendance",
  headerRow: 2,
  attendanceStartCol: 5,
  attendanceDateRange: "E2:JC2",
  firstDataRow: 5,
  nameCol: 4,
  gradeCol: 2,
  absentMarker: "A",
  absentDaysCount: 3,
  recipientEmails: "aniket.mishra@pw.live",
  emailSubject: "Absentee Alert: 3 Consecutive Days"
};

function sendAbsenteeAlert() {
  Logger.log("Starting absentee alert processing.");

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.sheetName);
  if (!sheet) {
    Logger.log(`Sheet '${CONFIG.sheetName}' not found. Exiting.`);
    return;
  }
  Logger.log(`Sheet '${CONFIG.sheetName}' found.`);

  var numCols = sheet.getRange(CONFIG.attendanceDateRange).getNumColumns();
  var headerRowValues = sheet.getRange(CONFIG.headerRow, CONFIG.attendanceStartCol, 1, numCols).getValues()[0];
  var lastRow = sheet.getLastRow();

  var dateCols = [];
  var dateValues = [];
  var today = new Date();
  today.setHours(0, 0, 0, 0);

  for (var i = headerRowValues.length - 1; i >= 0; i--) {
    var val = headerRowValues[i];
    var dateVal = null;
    if (val instanceof Date) {
      dateVal = new Date(val.getFullYear(), val.getMonth(), val.getDate());
    } else if (typeof val === "string" && val.trim() !== "") {
      dateVal = tryParseDateString(val.trim(), today.getFullYear());
    }
    if (dateVal && !isNaN(dateVal.getTime())) {
      var diffDay = (today - dateVal) / (24 * 3600 * 1000);
      if (diffDay >= 0 && diffDay < CONFIG.absentDaysCount) {
        dateCols.unshift(i);
        dateValues.unshift(dateVal);
        if (dateCols.length === CONFIG.absentDaysCount) break;
      }
    }
  }

  if (dateCols.length < CONFIG.absentDaysCount) {
    Logger.log("Not enough recent attendance dates.");
    return;
  }

  var attendanceData = sheet
    .getRange(CONFIG.firstDataRow, CONFIG.attendanceStartCol, lastRow - CONFIG.firstDataRow + 1, numCols)
    .getValues();
  var names = sheet.getRange(CONFIG.firstDataRow, CONFIG.nameCol, lastRow - CONFIG.firstDataRow + 1, 1).getValues();
  var grades = sheet.getRange(CONFIG.firstDataRow, CONFIG.gradeCol, lastRow - CONFIG.firstDataRow + 1, 1).getValues();

  var totalStudents = 0;
  var absentees = [];
  var gradeCounts = {};
  var gradePresent = {};
  var gradePresentPerDay = {};

  for (var r = 0; r < names.length; r++) {
    var name = names[r][0];
    var grade = grades[r][0];
    if (!name || !name.toString().trim()) continue;
    if (!grade) continue;
    var cleanName = name.toString().replace(/[^A-Za-z ]+/g, "").trim();
    if (!cleanName) continue;
    totalStudents++;
    gradeCounts[grade] = (gradeCounts[grade] || 0) + 1;
    if (!gradePresent[grade]) gradePresent[grade] = 0;
    if (!gradePresentPerDay[grade]) gradePresentPerDay[grade] = new Array(dateCols.length).fill(0);

    var isAbsentAllDays = true;
    var absentDatesList = [];

    for (var c = 0; c < dateCols.length; c++) {
      var status = attendanceData[r][dateCols[c]];
      var isAbsent = String(status).toUpperCase() === CONFIG.absentMarker;
      if (!isAbsent) {
        gradePresent[grade]++;
        gradePresentPerDay[grade][c]++;
      } else {
        absentDatesList.push(formatDateSimple(dateValues[c]));
      }
      if (isAbsentAllDays && isAbsent == false) isAbsentAllDays = false;
    }
    if (isAbsentAllDays) {
      absentees.push({ grade: grade, name: cleanName, dates: absentDatesList.join(", ") });
    }
  }

  var totalPresent = Object.values(gradePresent).reduce((a, b) => a + b, 0);
  var totalPossible = totalStudents * CONFIG.absentDaysCount;
  var totalPercentage = totalPossible === 0 ? 0 : Math.round((totalPresent * 100) / totalPossible);

  // Prepare data for email body
  var htmlRows = absentees
    .map((r) => `<tr><td>${r.grade}</td><td>${r.name}</td><td>${r.dates}</td></tr>`)
    .join("");

  var htmlBody = generateHtml(
    totalStudents,
    absentees.length,
    gradeCounts,
    gradePresentPerDay,
    dateValues,
    totalPresent,
    totalPercentage,
    htmlRows
  );

  MailApp.sendEmail({ to: CONFIG.recipientEmails, subject: CONFIG.emailSubject, htmlBody: htmlBody });
  Logger.log("Email sent.");
}

function generateHtml(
  totalStudents,
  totalAbsentees,
  gradeCounts,
  gradePresentPerDay,
  dateValues,
  totalPresent,
  totalPercentage,
  htmlRows
) {
  const gradeSummary = Object.keys(gradeCounts)
    .map((k) => `<span style="margin-right:10px"><b>${k}:</b> ${gradeCounts[k]}</span>`)
    .join("");

  const dateHeaders = dateValues.map((d) => `<th>${formatDateSimple(d)}</th>`).join("");

  const gradeRows = Object.keys(gradeCounts)
    .map((grade) => {
      const presentCounts = gradePresentPerDay[grade].map((c) => `<td>${c}</td>`).join("");
      return `<tr><td>${grade}</td>${presentCounts}</tr>`;
    })
    .join("");

  return `
<!DOCTYPE html>
<html>
<head>
  <style>
    body {font-family: Arial, sans-serif; background: #f9fafb; padding: 20px;}
    .container { max-width: 900px; margin: auto; background: white; padding: 20px; border-radius:10px; box-shadow: 0 2px 10px rgba(0,0,0,0.1);}
    h1 {color: #27632a; text-align: center;}
    .summary { text-align:center; margin-bottom:10px; font-size:16px; }
    .grade-summary {text-align:center; margin-bottom:30px; font-weight:bold; color: #27632a;}
    table {width: 100%; border-collapse: collapse; margin-bottom: 10px;}
    th, td {border:1px solid #ddd; padding: 16px; text-align: center;}
    th {background-color: #2e7d32; color: white;}
    tr:nth-child(even) {background-color: #f5f5f5;}
    tr:hover {background-color: #d4edda;}
    .alert-title {color: red; font-weight: bold; font-size: 20px; margin-bottom: 15px; text-align: center;}
  </style>
</head>
<body>
  <div class="container">
    <h1>Bazipur Attendance Day Wise</h1>
    <div class="summary" style="font-size: 20px;">
      Total Students: <b>${totalStudents}</b>
    </div>
    <div class="grade-summary" style="font-size:14px;">
      Grade-wise totals: ${gradeSummary}
    </div>
    <table>
      <thead>
        <tr>
          <th>Grade</th>
          ${dateHeaders}
        </tr>
      </thead>
      <tbody>
        ${gradeRows}
        <tr style="font-weight:bold;">
          <td>Total Present per Day</td>
          ${dateValues.map((_, i) => {
            const sumDay = Object.keys(gradeCounts).reduce((val, g) => val + gradePresentPerDay[g][i], 0);
            return `<td>${sumDay}</td>`;
          }).join("")}
        </tr>
        <tr style="font-weight:bold;">
          <td>Percentage per Day</td>
          ${dateValues.map((_, i) => {
            const sumDay = Object.keys(gradeCounts).reduce((val, g) => val + gradePresentPerDay[g][i], 0);
            const percent = totalStudents === 0 ? 0 : Math.round((sumDay * 100) / totalStudents);
            return `<td>${percent}%</td>`;
          }).join("")}
        </tr>
      </tbody>
    </table>
    <div class="alert-title">Absentee Alert: 3 Consecutive Days</div>
    <div style="text-align:center; font-weight:bold; margin-bottom:15px; font-size:20px; color: red;">Total Absentees: ${totalAbsentees}</div>
    <table>
      <thead>
        <tr>
          <th>Grade</th>
          <th>Student Name</th>
          <th>Absent Dates</th>
        </tr>
      </thead>
      <tbody>
        ${htmlRows}
      </tbody>
    </table>
  </div>
</body>
</html>`;
}

function formatDateSimple(dateObj) {
  return Utilities.formatDate(dateObj, Session.getScriptTimeZone(), "dd-MMM-yyyy");
}

function tryParseDateString(str, year) {
  const months = {
    Jan: 0,
    Feb: 1,
    Mar: 2,
    Apr: 3,
    May: 4,
    Jun: 5,
    Jul: 6,
    Aug: 7,
    Sep: 8,
    Oct: 9,
    Nov: 10,
    Dec: 11,
    Sept: 8
  };
  let dmy = str.match(/^(\d{1,2})[- ]([A-Za-z]{3,4})$/);
  let mdy = str.match(/^([A-Za-z]{3,4})[- ](\d{1,2})$/);
  if (dmy) {
    let day = parseInt(dmy[1], 10);
    let mon = months[dmy[2].substring(0, 3)];
    if (!isNaN(day) && mon !== undefined) return new Date(year, mon, day);
  } else if (mdy) {
    let mon = months[mdy[1].substring(0, 3)];
    let day = parseInt(mdy[2], 10);
    if (!isNaN(day) && mon !== undefined) return new Date(year, mon, day);
  }
  return null;
}
