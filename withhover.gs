const CONFIG = {
  sheetName: "Daily Attendance",
  headerRow: 2,
  attendanceStartCol: 5,
  attendanceStartRange: "E2:JC2",
  firstDataRow: 5,
  nameCol: 4,
  gradeCol: 2,
  absentMarker: "A",
  absentsDaysCount: 3,
  recipientEmails: "aniket.mishra@pw.live",
  emailSubject: "Absentee Alert: 3 Consecutive Days"
};

function sendAbsenteeAlert() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.sheetName);
  if (!sheet) {
    Logger.log(`Sheet '${CONFIG.sheetName}' not found. Exiting.`);
    return;
  }

  var headers = sheet.getRange(CONFIG.headerRow, CONFIG.attendanceStartCol, 1, sheet.getLastColumn() - CONFIG.attendanceStartCol + 1).getValues()[0];
  var today = new Date();
  today.setHours(0, 0, 0, 0);

  var latestDateColIndexes = [];
  var latestDateValues = [];

  for (var i = headers.length - 1; i >= 0; i--) {
    var dateValue = null;
    if (headers[i] instanceof Date) {
      dateValue = headers[i];
    } else if (typeof headers[i] === 'string' && headers[i].trim() !== '') {
      dateValue = tryParseDateString(headers[i].trim(), today.getFullYear());
    }
    if (dateValue && !isNaN(dateValue.getTime())) {
      const diffDays = (today - dateValue) / (1000 * 3600 * 24);
      if (diffDays >= 0 && diffDays <= CONFIG.absentsDaysCount) {
        latestDateColIndexes.unshift(i);
        latestDateValues.unshift(dateValue);
      }
    }
  }

  if (latestDateValues.length > 0 && latestDateValues[latestDateValues.length - 1].getTime() === today.getTime()) {
    latestDateColIndexes.pop();
    latestDateValues.pop();
  }

  while (latestDateColIndexes.length > CONFIG.absentsDaysCount) {
    latestDateColIndexes.shift();
    latestDateValues.shift();
  }

  if (latestDateColIndexes.length < CONFIG.absentsDaysCount) {
    Logger.log("Not enough recent dates found.");
    return;
  }

  var lastRow = sheet.getLastRow();

  var attendanceData = sheet.getRange(CONFIG.firstDataRow, CONFIG.attendanceStartCol, lastRow - CONFIG.firstDataRow + 1, headers.length).getValues();
  var names = sheet.getRange(CONFIG.firstDataRow, CONFIG.nameCol, lastRow - CONFIG.firstDataRow + 1, 1).getValues();
  var grades = sheet.getRange(CONFIG.firstDataRow, CONFIG.gradeCol, lastRow - CONFIG.firstDataRow + 1, 1).getValues();

  var totalStudents = 0;
  var absentees = [];
  var gradeCounts = {};
  var gradePresent = {};
  var gradePresentPerDay = {};

  var absentsPerDay = new Array(latestDateColIndexes.length).fill(0);

  for (var i = 0; i < names.length; i++) {
    var studentName = names[i][0];
    var grade = grades[i][0];
    if (!studentName || !grade) continue;
    const cleanName = studentName.toString().replace(/[^a-zA-Z ]/g, '').trim();
    if (!cleanName) continue;

    totalStudents++;
    gradeCounts[grade] = (gradeCounts[grade] || 0) + 1;
    gradePresent[grade] = gradePresent[grade] || 0;
    gradePresentPerDay[grade] = gradePresentPerDay[grade] || Array(latestDateColIndexes.length).fill(0);

    let absentAllDays = true;
    let absentDatesList = [];

    for (var j = 0; j < latestDateColIndexes.length; j++) {
      let colIndex = latestDateColIndexes[j];
      let status = attendanceData[i][colIndex];
      let isAbsent = String(status).toUpperCase() === CONFIG.absentMarker;
      if (!isAbsent) {
        gradePresent[grade]++;
        gradePresentPerDay[grade][j]++;
      } else {
        absentDatesList.push(formatDateSimple(latestDateValues[j]));
        absentsPerDay[j]++;
      }
      if (!isAbsent) {
        absentAllDays = false;
      }
    }

    if (absentAllDays) {
      absentees.push({grade: grade, name: cleanName, dates: absentDatesList.join(', ')});
    }
  }

  var totalPresent = Object.values(gradePresent).reduce((a,b) => a + b, 0);
  var maxPresent = totalStudents * CONFIG.absentsDaysCount;
  var overallPercentage = maxPresent > 0 ? Math.round((totalPresent / maxPresent) * 100) : 0;

  var htmlRows = absentees.map(a => `<tr><td>${a.grade}</td><td>${a.name}</td><td>${a.dates}</td></tr>`).join('');

  var htmlBody = generateHtml(totalStudents, absentees.length, gradeCounts, gradePresentPerDay, latestDateValues, totalPresent, overallPercentage, htmlRows, absentsPerDay);

  var emailSubject = `Absentee Alert: ${CONFIG.absentsDaysCount} Consecutive Days`;

  MailApp.sendEmail({
      to: CONFIG.recipientEmails,
      subject: emailSubject,
      htmlBody: htmlBody
  });

  Logger.log("Email sent.");
}

function generateHtml(totalStudents, totalAbsentees, gradeCounts, gradePresentPerDay, dateValues, totalPresent, overallPercentage, htmlRows, absentsPerDay) {
  const gradeSummaryCard = `
  <div style="
    display: flex;
    gap: 48px;
    justify-content: center;
    background: #f0faf1;
    border-radius: 20px;
    padding: 10px 15px;
    margin-bottom: 40px;
    box-shadow: 0 12px 24px rgba(0, 128, 0, 0.15);
    max-width: 640px;
    margin-left: auto;
    margin-right: auto;
    font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
  ">
    <div style="background:#d4ebd8; border-radius: 14px; padding: 25px 40px; min-width: 130px; text-align: center; box-shadow: 0 3px 8px rgba(0, 128, 0, 0.2);">
      <div style="font-size: 18px; font-weight: 700; color: #256829; margin-bottom: 8px;">Total Students</div>
      <div style="font-size: 40px; font-weight: 900; color: #178029;">${totalStudents}</div>
    </div>
    <div style="background:#cae6be; border-radius: 14px; padding: 25px 40px; min-width: 130px; text-align: center; box-shadow: 0 3px 8px rgba(0, 114, 0, 0.15);">
      <div style="font-size: 18px; font-weight: 700; color: #3d5822;">Grade 1</div>
      <div style="font-size: 38px; font-weight: 900; color: #2f6a0a;">${gradeCounts['Grade 1'] || 0}</div>
    </div>
    <div style="background:#d1e5fb; border-radius: 14px; padding: 25px 40px; min-width: 130px; text-align: center; box-shadow: 0 3px 8px rgba(21, 101, 209, 0.2);">
      <div style="font-size: 18px; font-weight: 700; color: #2363b9;">Grade 2</div>
      <div style="font-size: 38px; font-weight: 900; color: #1a53a1;">${gradeCounts['Grade 2'] || 0}</div>
    </div>
  </div>
`;

  const dateHeaders = dateValues.map(d => `<th>${formatDateSimple(d)}</th>`).join('');
  const gradeRows = Object.keys(gradeCounts).map(g => {
    const row = gradePresentPerDay[g].map(p => `<td class="present">${p}</td>`).join('');
    return `<tr><td>${g}</td>${row}</tr>`;
  }).join('');
  const totalPresentRow = dateValues.map((_,i) => {
    const sum = Object.keys(gradeCounts).reduce((acc, g) => acc + gradePresentPerDay[g][i], 0);
    return `<td class="present"><b>${sum}</b></td>`;
  }).join('');
  const percentageRow = dateValues.map((_,i) => {
    const sum = Object.keys(gradeCounts).reduce((acc, g) => acc + gradePresentPerDay[g][i], 0);
    const pct = totalStudents ? Math.round((sum / totalStudents)*100) : 0;
    return `<td class="present"><b>${pct}%</b></td>`;
  }).join('');
  const absenteeInfo = dateValues.map((d, i) => `<span class="absent-pill">${formatDateSimple(d)}: ${absentsPerDay[i]} Absent</span>`).join('');

  return `<!DOCTYPE html>
  <html>
  <head>
  <style>
    body { font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; background: #f6f9f8; margin:0; color: #444;}
    .container { max-width: 800px; margin:auto; background:#fff; padding: 40px 30px; border-radius: 10px; box-shadow: 0 2px 8px rgba(0,0,0,0.1);}
    h1 {text-align:center; color: #2e7d32; margin-bottom: 32px; font-weight:700;}
    table {width: 100%; border-collapse: collapse; margin-bottom: 30px;}
    th,td { border: 1px solid #ddd; padding: 14px; text-align:center;}
    th {background: #2e7d32; color:#fff; text-transform: uppercase;}
    tbody tr:nth-child(even) {background: #f3f6f5;}
    tbody tr:hover {background: #c7ebc5; transition: 0.3s; font-weight:bold;}
    .present {color: #2e7d32; font-weight: bold;}
    .absent {color: #a83232; font-weight: bold;}
    .summary {display: flex; justify-content: center; gap: 20px; margin-bottom: 32px;}
    .summary > div {background: #dcf5d8; border-radius: 12px; padding: 24px 40px; min-width: 100px; box-shadow: 0 2px 6px rgba(46,125,50,0.2); text-align:center;}
    .summary h3 { margin: 0 0 10px 0; color: #1f5420; font-size: 18px; font-weight: 600;}
    .summary span { font-size: 36px; font-weight: 800; color: #2e7d32;}
    .absent-pill { display: inline-block; background: #ffce00; color:#444; border-radius: 16px; padding: 8px 12px; margin-right:10px; font-weight:600; box-shadow: 0 1px 3px rgba(0,0,0,0.1);}
    .alert-box { background:#fff9e5; border-radius: 12px; border: 1px solid #ffd742; padding: 20px; box-shadow: 0 1px 6px rgba(255,215,0,0.3); margin-bottom: 32px; display:flex; align-items:center;}
    .alert-box .icon {font-size: 28px; color: #b58300; margin-right: 15px;}
    .alert-box .text { font-weight: 700; font-size: 20px; color: #b58300;}
    .badge {background: #f97154; color:#fff; padding: 14px 18px; border-radius: 14px; font-weight: 700; margin-right: 20px; box-shadow: 0 3px 12px rgba(255,85,55,0.5);}
    .footer { font-size: 14px; color: #666; padding-top: 20px; border-top: 1px solid #eee; text-align: center; font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;}
    .footer a {color: #2e7d32; text-decoration: none;}
    @media screen and (max-width: 600px) {
      .summary {flex-direction: column; align-items: center;}
      .summary > div {width: 80%; margin-bottom: 20px;}
      .badge {margin-bottom: 12px; margin-right: 12px;}
    }
  </style>
  </head>
  <body>
  <div class="container">
    <h1>Bazipur Attendance Day Wise</h1>
    ${gradeSummaryCard}
    <table>
      <thead>
        <tr>
          <th>Grade</th>
          ${dateHeaders}
        </tr>
      </thead>
      <tbody>
        ${gradeRows}
        <tr>
          <td><b>Total Present per Day</b></td>
          ${totalPresentRow}
        </tr>
        <tr>
          <td><b>Percentage per Day</b></td>
          ${percentageRow}
        </tr>
      </tbody>
    </table>
    <div class="alert-box">
      <span class="icon">&#9888;</span>
      <span class="text">Absentee Alert: ${CONFIG.absentsDaysCount} Consecutive Days</span>
      <span class="badge">Total Absentees: ${totalAbsentees}</span>
      ${absenteeInfo}
    </div>
    <table>
      <thead>
        <tr>
          <th>Grade</th>
          <th>Student Name</th>
          <th>Absent Dates</th>
        </tr>
      </thead>
      <tbody>${htmlRows}</tbody>
    </table>
    <div class="footer">
      &copy; 2025 Bazipur School. All rights reserved. <br/>
      Contact: <a href="mailto:info@bazipurschool.edu">info@bazipurschool.edu</a> | +91 12345 67890
    </div>
  </div>
  </body>
  </html>`;
}

function formatDateSimple(dateObj) {
  return Utilities.formatDate(dateObj, Session.getTimeZone(), "dd-MMM-yyyy");
}

function tryParseDateString(str, year) {
  const months = {
    Jan: 0, Feb: 1, Mar: 2, Apr: 3, May: 4,
    Jun: 5, Jul: 6, Aug: 7, Sep: 8, Oct: 9,
    Nov: 10, Dec: 11, Sept: 8
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
