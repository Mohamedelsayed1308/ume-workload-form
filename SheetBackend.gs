// ============================================================
// EMPLOYEE WORKLOAD TRACKING - Google Apps Script Backend
// ============================================================
// SETUP INSTRUCTIONS:
// 1. Open Google Sheets → Extensions → Apps Script
// 2. Delete existing code, paste this entire file
// 3. Click Deploy → New Deployment
// 4. Type: Web App
// 5. Execute as: Me | Access: Anyone
// 6. Click Deploy → Copy the Web App URL
// 7. Open index.html in your repo, replace:
//    'YOUR_GOOGLE_APPS_SCRIPT_WEB_APP_URL_HERE'
//    with the URL you copied
// 8. Commit the change → your form is now fully connected!
// ============================================================

var SHEET_NAME_SUBMISSIONS = 'Submissions';
var SHEET_NAME_DAILY       = 'Daily Tasks';
var SHEET_NAME_WEEKLY      = 'Weekly Tasks';
var SHEET_NAME_MONTHLY     = 'Monthly Tasks';

// ── Entry point for POST requests from the web form ──────
function doPost(e) {
  try {
    var params = e.parameter || {};
    var ss = SpreadsheetApp.getActiveSpreadsheet();

    // Initialize sheets if first run
    ensureSheets(ss);

    var timestamp      = params.timestamp      || new Date().toISOString();
    var submissionId   = params.submissionId   || 'WL-' + Date.now();
    var fullName       = params.fullName       || '';
    var jobTitle       = params.jobTitle       || '';
    var department     = params.department     || '';
    var teamRole       = params.teamRole       || '';
    var directManager  = params.directManager  || '';
    var startDate      = params.startDate      || '';
    var email          = params.email          || '';
    var reportPeriod   = params.reportPeriod   || '';

    var dailyTaskCount    = params.dailyTaskCount    || 0;
    var dailyTotalHours   = params.dailyTotalHours   || 0;
    var dailyTotalPct     = params.dailyTotalPct     || 0;
    var weeklyTaskCount   = params.weeklyTaskCount   || 0;
    var weeklyTotalHours  = params.weeklyTotalHours  || 0;
    var weeklyTotalPct    = params.weeklyTotalPct    || 0;
    var monthlyTaskCount  = params.monthlyTaskCount  || 0;
    var monthlyTotalHours = params.monthlyTotalHours || 0;
    var monthlyTotalPct   = params.monthlyTotalPct   || 0;

    var overallRating  = params.overallRating  || '';
    var wlbRating      = params.wlbRating      || '';
    var challenges     = params.challenges     || '';
    var suggestions    = params.suggestions    || '';
    var additionalNotes= params.additionalNotes|| '';

    var dailyTasksJSON   = params.dailyTasks   || '[]';
    var weeklyTasksJSON  = params.weeklyTasks  || '[]';
    var monthlyTasksJSON = params.monthlyTasks || '[]';

    // ── Write to Submissions (summary) sheet ──────────────
    var subSheet = ss.getSheetByName(SHEET_NAME_SUBMISSIONS);
    subSheet.appendRow([
      timestamp, submissionId, fullName, jobTitle, department, teamRole,
      directManager, startDate, email, reportPeriod,
      dailyTaskCount, dailyTotalHours, dailyTotalPct + '%',
      weeklyTaskCount, weeklyTotalHours, weeklyTotalPct + '%',
      monthlyTaskCount, monthlyTotalHours, monthlyTotalPct + '%',
      overallRating, wlbRating, challenges, suggestions, additionalNotes
    ]);

    // ── Write individual task rows ─────────────────────────
    writeTasks(ss, SHEET_NAME_DAILY,   submissionId, fullName, dailyTasksJSON);
    writeTasks(ss, SHEET_NAME_WEEKLY,  submissionId, fullName, weeklyTasksJSON);
    writeTasks(ss, SHEET_NAME_MONTHLY, submissionId, fullName, monthlyTasksJSON);

    return ContentService
      .createTextOutput(JSON.stringify({ status: 'success', id: submissionId }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: 'error', message: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ── Write task rows to the appropriate sheet ─────────────
function writeTasks(ss, sheetName, submissionId, employeeName, tasksJSON) {
  try {
    var tasks = JSON.parse(tasksJSON);
    if (!tasks || tasks.length === 0) return;
    var sheet = ss.getSheetByName(sheetName);
    tasks.forEach(function(task) {
      sheet.appendRow([
        new Date().toISOString(), submissionId, employeeName,
        task.name || '', task.cat || '', task.hrs || 0,
        task.prio || '', task.pct || 0, task.notes || ''
      ]);
    });
  } catch (err) {
    Logger.log('writeTasks error for ' + sheetName + ': ' + err);
  }
}

// ── Create and format sheets on first use ─────────────────
function ensureSheets(ss) {
  setupSheet(ss, SHEET_NAME_SUBMISSIONS, [
    'Timestamp', 'Submission ID', 'Full Name', 'Job Title', 'Department',
    'Team Role', 'Direct Manager', 'Start Date', 'Email', 'Report Period',
    'Daily Tasks', 'Daily Hours', 'Daily Workload %',
    'Weekly Tasks', 'Weekly Hours', 'Weekly Workload %',
    'Monthly Tasks', 'Monthly Hours', 'Monthly Workload %',
    'Workload Rating', 'Work-Life Balance', 'Challenges', 'Suggestions', 'Additional Notes'
  ], '#1a3c5e');

  setupSheet(ss, SHEET_NAME_DAILY, [
    'Timestamp', 'Submission ID', 'Employee Name',
    'Task Name', 'Category', 'Est. Hours', 'Priority', 'Workload %', 'Notes'
  ], '#1e40af');

  setupSheet(ss, SHEET_NAME_WEEKLY, [
    'Timestamp', 'Submission ID', 'Employee Name',
    'Task Name', 'Category', 'Est. Hours', 'Priority', 'Workload %', 'Notes'
  ], '#166534');

  setupSheet(ss, SHEET_NAME_MONTHLY, [
    'Timestamp', 'Submission ID', 'Employee Name',
    'Task Name', 'Category', 'Est. Hours', 'Priority', 'Workload %', 'Notes'
  ], '#5b21b6');
}

function setupSheet(ss, name, headers, hexColor) {
  var sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
  }
  // Only add headers if sheet is empty
  if (sheet.getLastRow() === 0) {
    var headerRange = sheet.getRange(1, 1, 1, headers.length);
    headerRange.setValues([headers]);
    headerRange.setBackground(hexColor);
    headerRange.setFontColor('#ffffff');
    headerRange.setFontWeight('bold');
    headerRange.setFontSize(10);
    sheet.setFrozenRows(1);
    // Auto-resize columns
    for (var i = 1; i <= headers.length; i++) {
      sheet.setColumnWidth(i, 140);
    }
    Logger.log('Sheet created: ' + name);
  }
}

// ── Test function — run this to verify setup ──────────────
function testSetup() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ensureSheets(ss);
  Logger.log('All sheets initialized successfully!');
  Logger.log('Spreadsheet URL: ' + ss.getUrl());
}
