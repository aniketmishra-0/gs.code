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
      if (diffDay >= 0 && diffDay <= CONFIG.absentDaysCount) {
        dateCols.unshift(i);
        dateValues.unshift(dateVal);
      }
    }
  }

  if (dateValues.length > 0 && dateValues[dateValues.length - 1].getTime() === today.getTime()) {
    dateCols.pop();
    dateValues.pop();
  }

  while (dateCols.length > CONFIG.absentDaysCount) {
    dateCols.shift();
    dateValues.shift();
  }

  if (dateCols.length < CONFIG.absentDaysCount) {
    Logger.log("Not enough recent attendance dates after excluding today.");
    return;
  }

  var subject = generateDynamicSubject(dateValues);

  var attendanceData = sheet.getRange(CONFIG.firstDataRow, CONFIG.attendanceStartCol, lastRow - CONFIG.firstDataRow + 1, numCols).getValues();
  var names = sheet.getRange(CONFIG.firstDataRow, CONFIG.nameCol, lastRow - CONFIG.firstDataRow + 1, 1).getValues();
  var grades = sheet.getRange(CONFIG.firstDataRow, CONFIG.gradeCol, lastRow - CONFIG.firstDataRow + 1, 1).getValues();

  var totalStudents = 0;
  var absentees = [];
  var gradeCounts = {};
  var gradePresent = {};
  var gradePresentPerDay = {};

  var absenteesPerDay = new Array(dateCols.length).fill(0);

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
        absenteesPerDay[c]++;
      }
      if (isAbsentAllDays && !isAbsent) isAbsentAllDays = false;
    }
    if (isAbsentAllDays) {
      absentees.push({ grade: grade, name: cleanName, dates: absentDatesList.join(", ") });
    }
  }

  var totalPresent = Object.values(gradePresent).reduce((a, b) => a + b, 0);
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

  MailApp.sendEmail({ to: CONFIG.recipientEmails, subject: subject, htmlBody: htmlBody });
  Logger.log("Email sent.");
}

function generateDynamicSubject(dateValues) {
  if (!dateValues || dateValues.length === 0) return "Bazipur School Attendance Summary";
  var tz = Session.getScriptTimeZone();
  var startDate = Utilities.formatDate(dateValues[0], tz, "d");
  var endDate = Utilities.formatDate(dateValues[dateValues.length - 1], tz, "d");
  var monthYear = Utilities.formatDate(dateValues[0], tz, "MMM yyyy");
  return startDate === endDate
    ? `Bazipur School Attendance Summary: ${startDate} ${monthYear}`
    : `Bazipur School Attendance Summary: ${startDate}–${endDate} ${monthYear}`;
}

// ✅ Updated UI: Centered, professional, modern look
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
  const gradeSummaryCard = `
    <table role="presentation" style="width:100%; margin-bottom:25px; border-collapse:collapse; text-align:center;">
      <tr>
        <td style="width:33.3%; padding:10px; vertical-align:top;">
          <div style="background:#ffffff; border-radius:12px; padding:20px; text-align:center; border:1px solid #e5e7eb; box-shadow:0 2px 6px rgba(0,0,0,0.1);">
            <div style="font-size:14px; font-weight:600; color:#6b7280; margin-bottom:6px;">Total Students</div>
            <div style="font-size:28px; font-weight:800; color:#16a34a;">${totalStudents}</div>
          </div>
        </td>
        <td style="width:33.3%; padding:10px; vertical-align:top;">
          <div style="background:#ffffff; border-radius:12px; padding:20px; text-align:center; border:1px solid #e5e7eb; box-shadow:0 2px 6px rgba(0,0,0,0.1);">
            <div style="font-size:14px; font-weight:600; color:#6b7280; margin-bottom:6px;">Grade 1</div>
            <div style="font-size:26px; font-weight:700; color:#0284c7;">${gradeCounts["Grade 1"] || 0}</div>
          </div>
        </td>
        <td style="width:33.3%; padding:10px; vertical-align:top;">
          <div style="background:#ffffff; border-radius:12px; padding:20px; text-align:center; border:1px solid #e5e7eb; box-shadow:0 2px 6px rgba(0,0,0,0.1);">
            <div style="font-size:14px; font-weight:600; color:#6b7280; margin-bottom:6px;">Grade 2</div>
            <div style="font-size:26px; font-weight:700; color:#7c3aed;">${gradeCounts["Grade 2"] || 0}</div>
          </div>
        </td>
      </tr>
    </table>
  `;

  const dateHeaders = dateValues.map((d) => `<th>${formatDateSimple(d)}</th>`).join("");
  const gradeRows = Object.keys(gradeCounts)
    .map((grade) => {
      const presentCounts = gradePresentPerDay[grade].map((c) => `<td>${c}</td>`).join("");
      return `<tr><td>${grade}</td>${presentCounts}</tr>`;
    })
    .join("");

  const totalPresentCells = dateValues
    .map((_, i) => {
      const sumDay = Object.keys(gradeCounts).reduce((val, g) => val + gradePresentPerDay[g][i], 0);
      return `<td><b>${sumDay}</b></td>`;
    })
    .join("");

  const percentageCells = dateValues
    .map((_, i) => {
      const sumDay = Object.keys(gradeCounts).reduce((val, g) => val + gradePresentPerDay[g][i], 0);
      const percent = totalStudents === 0 ? 0 : Math.round((sumDay * 100) / totalStudents);
      return `<td><b>${percent}%</b></td>`;
    })
    .join("");

  const dayWiseAbsentInfo = dateValues
    .map(
      (d, i) =>
        `<span style="background:#facc15; color:#1f2937; padding:6px 12px; border-radius:9999px; font-size:13px; font-weight:600; margin:2px; display:inline-block;">${formatDateSimple(
          d
        )}: ${absenteesPerDay[i] || 0} Absent</span>`
    )
    .join(" ");

  return `
  <div style="max-width:900px; margin:auto; background:#f9fafb; padding:30px; border-radius:14px; font-family:'Segoe UI', sans-serif; color:#111827; border:1px solid #e5e7eb;">
    
    <h2 style="text-align:center; margin-top:0; margin-bottom:20px; font-size:24px; color:#1f2937;">📊 Bazipur Attendance Report</h2>
    
    ${gradeSummaryCard}

    <table style="width:100%; border-collapse:collapse; margin-bottom:25px; font-size:14px; text-align:center;">
      <thead>
        <tr style="background:#e5e7eb; text-transform:uppercase; font-size:13px; letter-spacing:0.5px;">
          <th style="border:1px solid #d1d5db; padding:10px;">Grade</th>
          ${dateHeaders}
        </tr>
      </thead>
      <tbody>
        ${gradeRows}
        <tr style="background:#f9fafb; font-weight:600;">
          <td style="padding:10px; border:1px solid #d1d5db;">Total Present</td>
          ${totalPresentCells}
        </tr>
        <tr style="background:#f9fafb; font-weight:600;">
          <td style="padding:10px; border:1px solid #d1d5db;">% Present</td>
          ${percentageCells}
        </tr>
      </tbody>
    </table>

    ${
      absenteesPerDay.every(count => count === 0)
        ? `<div style="background:#dcfce7; border:1px solid #86efac; padding:18px; border-radius:12px; text-align:center; font-weight:600; color:#166534; box-shadow:0 2px 6px rgba(0,0,0,0.1);">🎉 Perfect attendance in last ${CONFIG.absentDaysCount} days!</div>`
        : `<div style="background:#fef9c3; border:1px solid #fde047; padding:18px; border-radius:12px; margin-bottom:20px; text-align:center; box-shadow:0 2px 6px rgba(0,0,0,0.05);">
            <strong style="color:#92400e; font-size:15px;">⚠️ Absentee Alert: ${CONFIG.absentDaysCount} Consecutive Days</strong><br/>
            <div style="margin-top:12px;">${dayWiseAbsentInfo}</div>
            <div style="margin-top:10px; font-weight:600; color:#b91c1c;">Total Absentees: ${totalAbsentees}</div>
          </div>`
    }

    <table style="width:100%; border-collapse:collapse; margin-top:10px; font-size:14px; text-align:center;">
      <thead>
        <tr style="background:#e5e7eb;">
          <th style="border:1px solid #d1d5db; padding:10px;">Grade</th>
          <th style="border:1px solid #d1d5db; padding:10px;">Student Name</th>
          <th style="border:1px solid #d1d5db; padding:10px;">Absent Dates</th>
        </tr>
      </thead>
      <tbody>
        ${htmlRows || `<tr><td colspan="3" style="text-align:center; padding:12px;">No consecutive absentees 🎉</td></tr>`}
      </tbody>
    </table>

    <div style="text-align:center; margin-top:30px; font-size:12px; color:#6b7280;">
      <div>© 2025 Bazipur School • All rights reserved</div>
      <div>Contact: <a href="mailto:info@bazipurschool.edu" style="color:#2563eb; text-decoration:none;">info@bazipurschool.edu</a></div>
    </div>
  </div>`;
}

function formatDateSimple(dateObj) {
  return Utilities.formatDate(dateObj, Session.getScriptTimeZone(), "dd-MMM-yyyy");
}

function tryParseDateString(str, year) {
  const months = {
    Jan: 0, Feb: 1, Mar: 2, Apr: 3, May: 4, Jun: 5,
    Jul: 6, Aug: 7, Sep: 8, Oct: 9, Nov: 10, Dec: 11,
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
