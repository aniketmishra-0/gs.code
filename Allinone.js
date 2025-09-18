// ✅ FIXED: Removed the extra space from the masterSheetName
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
  emailSubject: "Absentee Alert: 3 Consecutive Days",

  // --- SETTINGS FOR MASTER SHEET ---
  masterSheetId: "1NY7ByF0xrE7jaxAPhJsI2I2Xc6WV6Nul7kPLoocKqPE", 
  masterSheetName: "Garhi Bazidpur ",      // The name of the tab within that master spreadsheet
  masterNameCol: 2,                       // The column number for Student Name (A=1)
  masterMobileCol: 11                     // The column number for Parent Mobile (J=10)
};


// ✅ FIXED: Restored the correct function that uses the masterSheetId
function sendAbsenteeAlert() {
  Logger.log("Starting absentee alert processing.");

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.sheetName);
  if (!sheet) {
    Logger.log(`Sheet '${CONFIG.sheetName}' not found. Exiting.`);
    return;
  }
  Logger.log(`Sheet '${CONFIG.sheetName}' found.`);

  // --- NEW LOGIC: Fetch data from the master sheet using its ID ---
  const parentData = {};
  let masterSpreadsheet;

  // If a specific ID is provided, open that spreadsheet.
  if (CONFIG.masterSheetId) {
    try {
      masterSpreadsheet = SpreadsheetApp.openById(CONFIG.masterSheetId);
    } catch (e) {
      Logger.log(`ERROR: Could not open master spreadsheet with ID "${CONFIG.masterSheetId}". Please check the ID and script permissions. Error: ${e}`);
      // If opening by ID fails, we stop to avoid sending incomplete data.
      return; 
    }
  } else {
    // If no ID is provided, use the currently active spreadsheet.
    masterSpreadsheet = ss;
  }

  const masterSheet = masterSpreadsheet.getSheetByName(CONFIG.masterSheetName);
  if (!masterSheet) {
    Logger.log(`WARNING: Master sheet tab '${CONFIG.masterSheetName}' not found in the specified spreadsheet. Cannot fetch parent numbers.`);
  } else {
    // Read all master data at once for efficiency
    const masterRange = masterSheet.getRange(2, 1, masterSheet.getLastRow() - 1, CONFIG.masterMobileCol);
    const masterValues = masterRange.getValues();
    masterValues.forEach(row => {
      const name = row[CONFIG.masterNameCol - 1];
      const mobile = row[CONFIG.masterMobileCol - 1];
      if (name && mobile) {
        // Create a lookup map: { "Student Name": "Mobile Number" }
        parentData[name.toString().trim()] = mobile.toString().trim();
      }
    });
    Logger.log(`Loaded ${Object.keys(parentData).length} records from the master sheet.`);
  }
  // --- END NEW LOGIC ---

  const numCols = sheet.getRange(CONFIG.attendanceDateRange).getNumColumns();
  const headerRowValues = sheet.getRange(CONFIG.headerRow, CONFIG.attendanceStartCol, 1, numCols).getValues()[0];
  const lastRow = sheet.getLastRow();

  const dateCols = [];
  const dateValues = [];
  const today = new Date();
  today.setHours(0, 0, 0, 0);

  for (let i = headerRowValues.length - 1; i >= 0; i--) {
    const val = headerRowValues[i];
    let dateVal = null;
    if (val instanceof Date) {
      dateVal = new Date(val.getFullYear(), val.getMonth(), val.getDate());
    } else if (typeof val === "string" && val.trim() !== "") {
      dateVal = tryParseDateString(val.trim(), today.getFullYear());
    }
    if (dateVal && !isNaN(dateVal.getTime())) {
      const diffDay = (today - dateVal) / (24 * 3600 * 1000);
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

  const subject = generateDynamicSubject(dateValues);
  const attendanceData = sheet.getRange(CONFIG.firstDataRow, CONFIG.attendanceStartCol, lastRow - CONFIG.firstDataRow + 1, numCols).getValues();
  const names = sheet.getRange(CONFIG.firstDataRow, CONFIG.nameCol, lastRow - CONFIG.firstDataRow + 1, 1).getValues();
  const grades = sheet.getRange(CONFIG.firstDataRow, CONFIG.gradeCol, lastRow - CONFIG.firstDataRow + 1, 1).getValues();

  let totalStudents = 0;
  const absentees = [];
  const gradeCounts = {};
  const gradePresentPerDay = {};
  const gradeWiseAbsenteesPerDay = dateValues.map(() => ({}));

  for (let r = 0; r < names.length; r++) {
    const name = names[r][0];
    const grade = grades[r][0];
    if (!name || !name.toString().trim() || !grade) continue;
    
    const cleanName = name.toString().replace(/[^A-Za-z ]+/g, "").trim();
    if (!cleanName) continue;
    
    totalStudents++;
    gradeCounts[grade] = (gradeCounts[grade] || 0) + 1;
    if (!gradePresentPerDay[grade]) gradePresentPerDay[grade] = new Array(dateCols.length).fill(0);

    let isAbsentAllDays = true;
    const absentDatesList = [];

    for (let c = 0; c < dateCols.length; c++) {
      const status = attendanceData[r][dateCols[c]];
      const isAbsent = String(status).toUpperCase() === CONFIG.absentMarker;
      
      if (!isAbsent) {
        gradePresentPerDay[grade][c]++;
        isAbsentAllDays = false;
      } else {
        absentDatesList.push(formatDateSimple(dateValues[c]));
        gradeWiseAbsenteesPerDay[c][grade] = (gradeWiseAbsenteesPerDay[c][grade] || 0) + 1;
      }
    }
    
    if (isAbsentAllDays) {
      const mobileNo = parentData[cleanName] || "Not Found";
      absentees.push({
        grade: grade,
        name: cleanName,
        dates: absentDatesList.join(", "),
        mobile: mobileNo
      });
    }
  }
  
  const htmlRows = absentees.map(r => 
    `<tr>
       <td style="border:1px solid #e5e7eb; padding:8px; text-align:center;">${r.grade}</td>
       <td style="border:1px solid #e5e7eb; padding:8px; text-align:center;">${r.name}</td>
       <td style="border:1px solid #e5e7eb; padding:8px; text-align:center;">${r.mobile}</td>
       <td style="border:1px solid #e5e7eb; padding:8px; text-align:center;">${r.dates}</td>
     </tr>`
  ).join("");

  const htmlBody = generateHtml(
    totalStudents,
    absentees.length, 
    gradeCounts,
    gradePresentPerDay,
    dateValues,
    htmlRows,
    gradeWiseAbsenteesPerDay
  );

  MailApp.sendEmail({ to: CONFIG.recipientEmails, subject: subject, htmlBody: htmlBody });
  Logger.log("Email sent.");
}

function generateDynamicSubject(dateValues) {
  if (!dateValues || dateValues.length === 0) return "Bazipur School Attendance Summary";
  const tz = Session.getScriptTimeZone();
  const startDate = Utilities.formatDate(dateValues[0], tz, "d");
  const endDate = Utilities.formatDate(dateValues[dateValues.length - 1], tz, "d");
  const monthYear = Utilities.formatDate(dateValues[0], tz, "MMM yyyy");
  return startDate === endDate
    ? `Bazipur School Attendance Summary: ${startDate} ${monthYear}`
    : `Bazipur School Attendance Summary: ${startDate}–${endDate} ${monthYear}`;
}


function generateHtml(
  totalStudents,
  totalAbsentees,
  gradeCounts,
  gradePresentPerDay,
  dateValues,
  htmlRows,
  gradeWiseAbsenteesPerDay
) {
  const sortedGrades = Object.keys(gradeCounts).sort();
  let gradeCardsHtml = '';

  sortedGrades.forEach(grade => {
    gradeCardsHtml += `
      <td class="summary-card-cell" style="padding: 0 8px; vertical-align: top;">
        <div style="background:#f9fafb; border-radius:12px; padding:20px; text-align:center; border:1px solid #e5e7eb; height: 100%;">
          <div style="font-size:15px; font-weight:600; color:#374151; margin-bottom:6px;">${grade}</div>
          <div style="font-size:26px; font-weight:700; color:#4f46e5;">${gradeCounts[grade] || 0}</div>
        </div>
      </td>
    `;
  });

  const gradeSummaryCard = `
    <table class="summary-card-table" role="presentation" style="width:100%; margin-bottom:25px; border-spacing:0; border-collapse: separate;">
      <tr>
        <td class="summary-card-cell" style="padding: 0 8px; vertical-align: top;">
          <div style="background:#f0fdf4; border-radius:12px; padding:20px; text-align:center; border:1px solid #bbf7d0; height: 100%;">
            <div style="font-size:15px; font-weight:600; color:#166534; margin-bottom:6px;">Total Students</div>
            <div style="font-size:28px; font-weight:800; color:#16a34a;">${totalStudents}</div>
          </div>
        </td>
        ${gradeCardsHtml}
      </tr>
    </table>
  `;

  const dateHeaders = dateValues.map((d) => `<th style="border:1px solid #e5e7eb; padding:10px; text-align:center;">${formatDateSimple(d)}</th>`).join("");
  const gradeRows = sortedGrades.map((grade) => {
    const presentCounts = (gradePresentPerDay[grade] || []).map((c) => `<td style="border:1px solid #e5e7eb; padding:10px; text-align:center;">${c}</td>`).join("");
    return `<tr><td style="border:1px solid #e5e7eb; padding:10px; font-weight:600; text-align:center;">${grade}</td>${presentCounts}</tr>`;
  }).join("");

  const totalPresentCells = dateValues.map((_, i) => {
    const sumDay = sortedGrades.reduce((val, g) => val + ((gradePresentPerDay[g] || [])[i] || 0), 0);
    return `<td style="border:1px solid #e5e7eb; padding:10px; text-align:center;"><b>${sumDay}</b></td>`;
  }).join("");

  const percentageCells = dateValues.map((_, i) => {
    const sumDay = sortedGrades.reduce((val, g) => val + ((gradePresentPerDay[g] || [])[i] || 0), 0);
    const percent = totalStudents === 0 ? 0 : Math.round((sumDay * 100) / totalStudents);
    return `<td style="border:1px solid #e5e7eb; padding:10px; text-align:center;"><b>${percent}%</b></td>`;
  }).join("");

  const dayWiseAbsentInfo = dateValues.map((dateObj, dayIndex) => {
      const dailyCounts = gradeWiseAbsenteesPerDay[dayIndex];
      const gradesWithAbsentees = Object.keys(dailyCounts).sort();
      
      let totalDailyAbsentees = 0;
      const gradeDetails = gradesWithAbsentees.map(grade => {
        const count = dailyCounts[grade];
        totalDailyAbsentees += count;
        const shortGrade = grade.replace('Grade ', 'G'); 
        return `${shortGrade}: ${count}`;
      }).join(', ');

      const formattedDate = Utilities.formatDate(dateObj, Session.getScriptTimeZone(), "dd-MMM");
      
      return `<span style="display:inline-block; background:#fefce8; color:#713f12; padding:5px 12px; border-radius:16px; font-size:13px; font-weight:600; border:1px solid #fde047; margin: 4px;">
                ${formattedDate}: <b>${totalDailyAbsentees}</b> ${gradeDetails ? `(${gradeDetails})` : ''}
              </span>`;
    }).join(" ");

  return `
  <!DOCTYPE html>
  <html lang="en">
  <head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Bazipur Attendance Report</title>
    <style>
      @media screen and (max-width: 600px) {
        .container { padding: 15px !important; width: 100% !important; box-sizing: border-box; }
        .summary-card-cell { display: block !important; width: 100% !important; padding: 8px 0 !important; box-sizing: border-box; }
        .main-table, .absentee-table { font-size: 12px !important; }
        .main-table th, .main-table td, .absentee-table th, .absentee-table td { padding: 6px 4px !important; }
        h2 { font-size: 20px !important; }
      }
    </style>
  </head>
  <body>
    <div class="container" style="max-width:900px; margin:auto; background:#ffffff; padding:30px; border-radius:12px; font-family:'Segoe UI', sans-serif; color:#111827; border:1px solid #e5e7eb;">
      <h2 style="text-align:center; margin-top:0; margin-bottom:20px; font-size:22px; color:#111827;">Bazipur Attendance Report</h2>
      
      ${gradeSummaryCard}

      <table class="main-table" style="width:100%; border-collapse:collapse; margin-bottom:20px; font-size:14px;">
        <thead>
          <tr style="background:#f3f4f6; text-transform:uppercase; font-size:13px; letter-spacing:0.5px; color: #374151;">
            <th style="border:1px solid #e5e7eb; padding:10px; text-align:center;">Grade</th>
            ${dateHeaders}
          </tr>
        </thead>
        <tbody>
          ${gradeRows}
          <tr style="background:#f9fafb; font-weight:600;">
            <td style="padding:10px; border:1px solid #e5e7eb; text-align:center;">Total Present</td>
            ${totalPresentCells}
          </tr>
          <tr style="background:#f9fafb; font-weight:600;">
            <td style="padding:10px; border:1px solid #e5e7eb; text-align:center;">% Present</td>
            ${percentageCells}
          </tr>
        </tbody>
      </table>

      ${
        totalAbsentees === 0
          ? `<div style="background:#dcfce7; border:1px solid #86efac; padding:15px; border-radius:10px; text-align:center; font-weight:600; color:#166534;">🎉 No students were absent for ${CONFIG.absentDaysCount} consecutive days!</div>`
          : `<div style="background:#fffbeb; border:1px solid #fde047; padding:15px; border-radius:10px; margin-bottom:20px; text-align:center;">
              <strong style="color:#b45309; font-size:16px;">⚠️ Absentee Alert: ${CONFIG.absentDaysCount} Consecutive Days</strong><br/>
              <div style="margin-top:12px; font-size:14px; font-weight:600; color:#b91c1c;">Total Students Absent: ${totalAbsentees}</div>
              <div style="margin-top:12px; text-align:center;">${dayWiseAbsentInfo}</div>
            </div>
            <table class="absentee-table" style="width:100%; border-collapse:collapse; margin-top:10px; font-size:14px;">
              <thead>
                <tr style="background:#f3f4f6; color:#374151;">
                  <th style="border:1px solid #e5e7eb; padding:8px; text-align:center;">Grade</th>
                  <th style="border:1px solid #e5e7eb; padding:8px; text-align:center;">Student Name</th>
                  <th style="border:1px solid #e5e7eb; padding:8px; text-align:center;">Parent Mobile No.</th>
                  <th style="border:1px solid #e5e7eb; padding:8px; text-align:center;">Absent Dates</th>
                </tr>
              </thead>
              <tbody>
                ${htmlRows}
              </tbody>
            </table>`
      }

      <div style="text-align:center; margin-top:25px; font-size:12px; color:#6b7280;">
        <div>© 2025 Bazipur School • All rights reserved</div>
        <div>Contact: <a href="mailto:info@bazipurschool.edu" style="color:#2563eb; text-decoration:none;">info@bazipurschool.edu</a></div>
      </div>
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
    Jul: 6, Aug: 7, Sep: 8, Oct: 9, Nov: 10, Dec: 11,
    Sept: 8
  };
  const dmy = str.match(/^(\d{1,2})[- ]([A-Za-z]{3,4})$/);
  const mdy = str.match(/^([A-Za-z]{3,4})[- ](\d{1,2})$/);
  if (dmy) {
    const day = parseInt(dmy[1], 10);
    const mon = months[dmy[2].substring(0, 3)];
    if (!isNaN(day) && mon !== undefined) return new Date(year, mon, day);
  } else if (mdy) {
    const mon = months[mdy[1].substring(0, 3)];
    const day = parseInt(mdy[2], 10);
    if (!isNaN(day) && mon !== undefined) return new Date(year, mon, day);
  }
  return null;
}
