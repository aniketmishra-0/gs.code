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
  today.setHours(0,0,0,0);

  for (var i = headerRowValues.length - 1; i >= 0; i--) {
    var val = headerRowValues[i];
    var dateVal = null;
    if (val instanceof Date) {
      dateVal = new Date(val.getFullYear(), val.getMonth(), val.getDate());
    } else if (typeof val === "string" && val.trim() !== "") {
      dateVal = tryParseDateString(val.trim(), today.getFullYear());
    }
    if (dateVal && !isNaN(dateVal.getTime())) {
      var diffDay = (today - dateVal) / (24*3600*1000);
      if (diffDay >= 0 && diffDay <= CONFIG.absentDaysCount) {
        dateCols.unshift(i);
        dateValues.unshift(dateVal);
      }
    }
  }

  // Remove today if present at the end
  if (dateValues.length > 0 && dateValues[dateValues.length - 1].getTime() === today.getTime()) {
    dateCols.pop();
    dateValues.pop();
  }

  // Keep last N days only
  while (dateCols.length > CONFIG.absentDaysCount) {
    dateCols.shift();
    dateValues.shift();
  }

  if (dateCols.length < CONFIG.absentDaysCount) {
    Logger.log("Not enough recent attendance dates after excluding today.");
    return;
  }

  var subject = generateDynamicSubject(dateValues);

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

  var absenteesPerDay = new Array(dateCols.length).fill(0);

  for (var r=0; r<names.length; r++) {
    var name = names[r][0];
    var grade = grades[r][0];
    if (!name || !name.toString().trim()) continue;
    if (!grade) continue;
    var cleanName = name.toString().replace(/[^A-Za-z ]+/g,"").trim();
    if (!cleanName) continue;
    totalStudents++;
    gradeCounts[grade] = (gradeCounts[grade]||0)+1;
    if (!gradePresent[grade]) gradePresent[grade] = 0;
    if (!gradePresentPerDay[grade]) gradePresentPerDay[grade] = new Array(dateCols.length).fill(0);

    var isAbsentAllDays = true;
    var absentDatesList = [];

    for (var c=0; c<dateCols.length; c++) {
      var status = attendanceData[r][dateCols[c]];
      var isAbsent = String(status).toUpperCase() === CONFIG.absentMarker;
      if (!isAbsent) {
        gradePresent[grade]++;
        gradePresentPerDay[grade][c]++;
      } else {
        absentDatesList.push(formatDateSimple(dateValues[c]));
        absenteesPerDay[c]++;
      }
      if (isAbsentAllDays && !isAbsent) isAbsentAllDays = false;
    }
    if (isAbsentAllDays) {
      absentees.push({grade: grade, name: cleanName, dates: absentDatesList.join(", ")});
    }
  }

  var totalPresent = Object.values(gradePresent).reduce((a,b)=>a+b,0);
  var totalPossible = totalStudents * CONFIG.absentDaysCount;
  var totalPercentage = totalPossible === 0 ? 0 : Math.round((totalPresent * 100) / totalPossible);

  var htmlRows = absentees.map(r => `<tr><td>${r.grade}</td><td>${r.name}</td><td>${r.dates}</td></tr>`).join("");

  var htmlBody = generateHtml(
    totalStudents,
    absentees.length,
    gradeCounts,
    gradePresentPerDay,
    dateValues,
    totalPresent,
    totalPercentage,
    htmlRows,
    absenteesPerDay
  );

  MailApp.sendEmail({to: CONFIG.recipientEmails, subject: subject, htmlBody: htmlBody});
  Logger.log("Email sent.");
}

function generateDynamicSubject(dateValues) {
  if (!dateValues || dateValues.length === 0) return "Bazipur School Attendance Summary";
  var tz = Session.getScriptTimeZone();
  var startDate = Utilities.formatDate(dateValues[0], tz, "d");
  var endDate = Utilities.formatDate(dateValues[dateValues.length-1], tz, "d");
  var monthYear = Utilities.formatDate(dateValues[0], tz, "MMM yyyy");
  return (startDate === endDate)
    ? `Bazipur School Attendance Summary: ${startDate} ${monthYear}`
    : `Bazipur School Attendance Summary: ${startDate}–${endDate} ${monthYear}`;
}

function generateHtml(
  totalStudents,
  totalAbsentees,
  gradeCounts,
  gradePresentPerDay,
  dateValues,
  totalPresent,
  totalPercentage,
  htmlRows,
  absenteesPerDay
) {
  const gradeSummary = Object.keys(gradeCounts).map(k => `<span style="margin-right:10px"><b>${k}:</b> ${gradeCounts[k]}</span>`).join("");
  const dateHeaders = dateValues.map(d => `<th>${formatDateSimple(d)}</th>`).join("");
  
  const gradeRows = Object.keys(gradeCounts).map(grade => {
    const presentCounts = gradePresentPerDay[grade].map(c => `<td class="present">${c}</td>`).join("");
    return `<tr><td>${grade}</td>${presentCounts}</tr>`;
  }).join("");

  const totalPresentCells = dateValues.map((_,i) => {
    const sumDay = Object.keys(gradeCounts).reduce((val,g)=>val+gradePresentPerDay[g][i],0);
    return `<td class="present"><b>${sumDay}</b></td>`;
  }).join("");

  const percentageCells = dateValues.map((_,i) => {
    const sumDay = Object.keys(gradeCounts).reduce((val,g)=>val+gradePresentPerDay[g][i], 0);
    const percent = totalStudents === 0 ? 0 : Math.round((sumDay * 100) / totalStudents);
    return `<td class="present"><b>${percent}%</b></td>`;
  }).join("");

  const dayWiseAbsentInfo = dateValues.map((d,i) => `<span style="display:inline-block; margin-right:15px; font-weight:bold; color:#d32f2f;">${formatDateSimple(d)}: ${absenteesPerDay[i] || 0} Absent</span>`).join("");

  return `
<!DOCTYPE html>
<html>
<head>
  <style>
    body {
      font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
      background: #f3f6f8;
      padding: 25px;
      color: #37474f;
      margin: 0;
    }
    .container {
      max-width: 920px;
      margin: auto;
      background: #ffffff;
      padding: 30px 40px;
      border-radius: 14px;
      box-shadow: 0 10px 30px rgba(0, 105, 92, 0.15);
    }
    h1 {
      color: #004d40;
      text-align: center;
      font-weight: 700;
      letter-spacing: 1.2px;
      margin-top: 0;
    }
    .summary, .grade-summary {
      text-align: center;
      margin-bottom: 20px;
      font-weight: 600;
      color: #00796b;
    }
    table {
      width: 100%;
      border-collapse: collapse;
      margin-bottom: 18px;
      box-shadow: 0 3px 6px rgba(0,0,0,0.05);
      table-layout: fixed;
      word-wrap: break-word;
    }
    th, td {
      border: 1px solid #b2dfdb;
      padding: 14px 12px;
      text-align: center;
      font-size: 16px;
    }
    th {
      background-color: #00796b;
      color: #fff;
      font-weight: 700;
      text-transform: uppercase;
      letter-spacing: 1px;
    }
    tbody tr:nth-child(even) {
      background-color: #e0f2f1;
    }
    tbody tr:hover {
      background-color: #b2dfdb;
      box-shadow: 0 2px 8px rgba(0, 121, 107, 0.3);
      transition: all 0.25s ease;
    }
    td.present {
      color: #388e3c;
      font-weight: 600;
    }
    td.absent {
      color: #d32f2f;
      font-weight: 700;
    }
    .alert-title {
      font-size: 22px;
      font-weight: 800;
      text-align: center;
      color: #d32f2f;
      margin-bottom: 15px;
    }
    .alert-container {
      text-align: center;
      max-width: 480px;
      margin: 30px auto 5px;
      background: linear-gradient(135deg, #b2dfdb 0%, #004d40 100%);
      border-radius: 45px;
      box-shadow: 0 8px 20px rgba(0,77,64,0.2);
      padding: 40px 30px 30px 30px;
      position: relative;
      color: #e0f7fa;
      font-weight: 700;
    }
    .icon-circle {
      background-color: #fbc02d;
      width: 80px;
      height: 80px;
      border-radius: 50%;
      margin: 0 auto;
      box-shadow: 0 2px 12px #b38600;
      display: flex;
      align-items: center;
      justify-content: center;
      position: absolute;
      top: -40px;
      left: 50%;
      transform: translateX(-50%);
    }
    .alert-title-strong {
      margin-top: 60px;
      font-size: 2.4rem;
      letter-spacing: 1.2px;
      text-shadow: 0 3px 9px #fbc02d;
    }
    .alert-message {
      margin-top: 12px;
      font-size: 1.3rem;
      color: #ffeb3b;
    }
    .alert-message-sub {
      margin-top: 14px;
      font-size: 1.1rem;
      color: #fff9c4;
    }
    /* Responsive for mobile */
    @media only screen and (max-width: 600px) {
      body {
        padding: 10px !important;
        font-size: 14px !important;
      }
      .container {
        padding: 20px 15px !important;
        width: 100% !important;
        box-shadow: none !important;
        border-radius: 8px !important;
      }
      table, th, td {
        font-size: 13px !important;
      }
      th, td {
        padding: 8px 6px !important;
      }
      h1 {
        font-size: 24px !important;
      }
      .alert-container {
        padding: 25px 20px 20px 20px !important;
        border-radius: 30px !important;
      }
      .alert-title-strong {
        font-size: 1.8rem !important;
      }
      .alert-message {
        font-size: 1.1rem !important;
      }
    }
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
          ${totalPresentCells}
        </tr>
        <tr style="font-weight:bold;">
          <td>Percentage per Day</td>
          ${percentageCells}
        </tr>
      </tbody>
    </table>
    ${
      absenteesPerDay.every(count => count === 0)
        ? `
          <div class="alert-container">
            <div class="icon-circle">
              <svg viewBox="0 0 24 24" fill="#fff" width="44" height="44" xmlns="http://www.w3.org/2000/svg">
                <path d="M9.7 16.3l-4-4c-.39-.39-1.02-.39-1.41 0-.39.39-.39 1.02 0 1.41l4.7 4.7c.39.39 1.02.39 1.41 0l9-9c.39-.39.39-1.02 0-1.41-.39-.39-1.02-.39-1.41 0l-8.3 8.3z"/>
              </svg>
            </div>
            <div class="alert-title-strong">Perfect Attendance!</div>
            <div class="alert-message">
              All students have perfect attendance over the last 3 days!
            </div>
            <div class="alert-message-sub">
              🌟 Congratulations to the entire class! 🌟
            </div>
            <div style="margin-top: 15px; font-weight:bold; color:#d32f2f;">
              ${dayWiseAbsentInfo}
            </div>
          </div>
        `
        : `
          <div class="alert-title">Absentee Alert: 3 Consecutive Days</div>
          <div style="text-align:center; font-weight:bold; margin-bottom:15px; font-size:20px; color: #d32f2f;">
            Total Absentees: ${totalAbsentees}
          </div>
          <div style="margin-top: 10px; font-weight:bold; color:#d32f2f; text-align:center;">
            ${dayWiseAbsentInfo}
          </div>
        `
    }
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
    Jan: 0, Feb: 1, Mar: 2, Apr: 3, May: 4, Jun: 5,
    Jul: 6, Aug: 7, Sep: 8, Oct: 9, Nov: 10, Dec: 11, Sept: 8
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
