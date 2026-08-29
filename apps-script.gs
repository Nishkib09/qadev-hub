// QA Audit Script — Per-Agent Folder + Auto Email + Diagnostics
// Copy this entire file into your Apps Script editor.
// Then: Deploy → Manage Deployments → edit pencil → New version → Deploy → Authorize

var MAX_DASHBOARD_AUDIT_ROWS = 5000;
var MAX_EMAIL_STATUS_LOOKBACK_ROWS = 10000;
var MAX_TRACKER_AUDIT_ROWS_PER_TYPE = 500;
var MAX_TRACKER_LOG_ROWS = 2000;
var AUDIT_VERSION_PROPERTY = "QA_AUDIT_VERSION";

function doPost(e) {
  try {
    var d = JSON.parse(e.postData.contents);
    var doc = SpreadsheetApp.getActiveSpreadsheet();
    var fileUrls = [];
    var fileUrlStr = "No files attached";
    var attachments = [];
    var agentFolderUrl = "";
    var auditId = d.auditId ? String(d.auditId) : "";
    var duplicateAudit = false;

    var diag = 'Payload: files=' + (d.files ? d.files.length : 0) +
               ', agentName="' + (d.agentName || 'EMPTY') + '"' +
               ', emailOnly=' + (d.emailOnly ? 'yes' : 'no');

    // --- Resolve Agent Folder in Drive (early) ---
    var agentFolder;
    if (d.agentName) {
      try {
        var rootIter = DriveApp.getFoldersByName("QA Recordings");
        var rootFolder = rootIter.hasNext() ? rootIter.next() : DriveApp.createFolder("QA Recordings");

        var agentIter = rootFolder.getFoldersByName(d.agentName);
        if (agentIter.hasNext()) {
          agentFolder = agentIter.next();
        } else {
          agentFolder = rootFolder.createFolder(d.agentName);
        }

        if (d.agentEmail) {
          try { agentFolder.addViewer(d.agentEmail); } catch(e) {}
        }
        agentFolderUrl = agentFolder.getUrl();
      } catch (folderErr) {
        diag += ' | FolderResolve=FAIL:' + folderErr.toString();
      }
    }

    // --- Decode files for email attachments ---
    if (d.files && d.files.length > 0) {
      var today = new Date();
      var dateStr = today.getFullYear() + "-"
        + String(today.getMonth() + 1).padStart(2, "0") + "-"
        + String(today.getDate()).padStart(2, "0");
      var safeName = (d.agentName || "Agent").replace(/[^a-zA-Z0-9]/g, "_");
      var safeType = (d.auditType || "Audit").replace(/[^a-zA-Z0-9]/g, "_");

      for (var j = 0; j < d.files.length; j++) {
        try {
          var f = d.files[j];
          var ext = f.fileName.split(".").pop();
          var markedName = safeName + "_" + safeType + "_" + dateStr + "_" + (j+1) + "." + ext;

          var parts = f.fileData.split(",");
          var b64 = parts.length > 1 ? parts[1] : parts[0];
          var bytes = Utilities.base64Decode(b64);
          var blob = Utilities.newBlob(bytes, f.fileMimeType || "audio/mpeg", markedName);
          attachments.push(blob);
        } catch (fileErr) {
          diag += ' | FileDecode' + (j+1) + '=FAIL:' + fileErr.toString();
        }
      }
    }

    if (d.emailOnly) {
      // --- Update Status in Google Sheet (no new row) ---
      try {
        var sheetName = d.auditType;
        var sheet = doc.getSheetByName(sheetName);
        if (sheet) {
          // 1. Ensure new columns/headers are present in the sheet
          if (d.headers && d.headers.length > 0) {
            ensureHeaders(sheet, d.headers);
          }

          var lastRow = sheet.getLastRow();
          var lastColumn = sheet.getLastColumn();
          var headers = lastColumn > 0 ? sheet.getRange(1, 1, 1, lastColumn).getValues()[0] : [];
          var dataStartRow = Math.max(2, lastRow - MAX_EMAIL_STATUS_LOOKBACK_ROWS + 1);
          var dataRowCount = Math.max(0, lastRow - dataStartRow + 1);
          var values = dataRowCount > 0
            ? sheet.getRange(dataStartRow, 1, dataRowCount, lastColumn).getValues()
            : [];
          var displayValues = dataRowCount > 0
            ? sheet.getRange(dataStartRow, 1, dataRowCount, lastColumn).getDisplayValues()
            : [];

          var foundRowIndex = -1;
          var sheetTimezone = doc.getSpreadsheetTimeZone();

          for (var r = values.length - 1; r >= 0; r--) {
            var cellVal = values[r][0];
            var displayVal = displayValues[r][0];
            var isMatched = false;

            // Prefer the permanent Audit ID. Timestamp matching remains only for
            // legacy rows created before Audit IDs were introduced.
            var auditIdColIndex = headers.indexOf("Audit ID");
            if (auditId && auditIdColIndex !== -1 && String(values[r][auditIdColIndex]) === auditId) {
              isMatched = true;
            }

            // Check exact matches first
            if (!isMatched && (cellVal == d.timestamp || displayVal == d.timestamp)) {
              isMatched = true;
            } else if (!isMatched) {
              // Try timezone formatting match
              if (cellVal) {
                try {
                  var cellDate = (cellVal instanceof Date) ? cellVal : new Date(cellVal);
                  if (!isNaN(cellDate.getTime())) {
                    var cellStr = Utilities.formatDate(cellDate, sheetTimezone, "yyyy-MM-dd HH:mm:ss");
                    if (cellStr == d.timestamp) {
                      isMatched = true;
                    } else {
                      var cellStrScript = Utilities.formatDate(cellDate, Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");
                      if (cellStrScript == d.timestamp) {
                        isMatched = true;
                      } else {
                        var cellStrUTC = Utilities.formatDate(cellDate, "UTC", "yyyy-MM-dd HH:mm:ss");
                        if (cellStrUTC == d.timestamp) {
                          isMatched = true;
                        }
                      }
                    }
                  }
                } catch (dateErr) {
                  // Ignore parsing errors
                }
              }

              // Try component-based match (ignores timezone shifts)
              if (!isMatched && cellVal && d.timestamp) {
                try {
                  var cellDate = (cellVal instanceof Date) ? cellVal : new Date(cellVal);
                  var parts = d.timestamp.split(/[- :]/);
                  if (parts.length >= 6) {
                    var y = parseInt(parts[0], 10);
                    var m = parseInt(parts[1], 10) - 1;
                    var day = parseInt(parts[2], 10);
                    var hr = parseInt(parts[3], 10);
                    var min = parseInt(parts[4], 10);
                    var sec = parseInt(parts[5], 10);

                    var targetDateLocal = new Date(y, m, day, hr, min, sec);
                    var cellDateLocal = new Date(
                      cellDate.getFullYear(),
                      cellDate.getMonth(),
                      cellDate.getDate(),
                      cellDate.getHours(),
                      cellDate.getMinutes(),
                      cellDate.getSeconds()
                    );

                    var diffMs = Math.abs(targetDateLocal.getTime() - cellDateLocal.getTime());
                    if (diffMs <= 5000) { // 5-second tolerance
                      isMatched = true;
                    }
                  }
                } catch (fallbackErr) {
                  // Ignore
                }
              }
            }

            // Verify that the agent name matches as a safety check
            if (isMatched) {
              var rowStr = JSON.stringify(values[r]);
              if (!d.agentName || rowStr.indexOf(d.agentName) !== -1) {
                foundRowIndex = dataStartRow + r;
                break;
              }
            }
          }

          if (foundRowIndex !== -1) {
            var statusCol = headers.indexOf("Email Sent Status") + 1;
            var timeCol = headers.indexOf("Email Sent Timestamp") + 1;

            if (statusCol > 0) {
              sheet.getRange(foundRowIndex, statusCol).setValue("Sent");
            }
            if (timeCol > 0) {
              sheet.getRange(foundRowIndex, timeCol).setValue(new Date().toLocaleString());
            }
            diag += ' | SheetUpdate=OK (row ' + foundRowIndex + ')';
          } else {
            diag += ' | SheetUpdate=ROW_NOT_FOUND (ts=' + d.timestamp + ')';
          }
        } else {
          diag += ' | SheetUpdate=SHEET_NOT_FOUND (' + sheetName + ')';
        }
      } catch (sheetUpdateErr) {
        diag += ' | SheetUpdate=FAIL:' + sheetUpdateErr.toString();
      }

    } else {
      // --- Normal Submit: Write to Google Sheet & Upload to Drive ---
      var sheetName = d.auditType;
      var sheet = doc.getSheetByName(sheetName);
      if (!sheet) {
        sheet = doc.insertSheet(sheetName);
      }
      auditId = auditId || Utilities.getUuid();
      d.headers = d.headers || [];
      d.row = d.row || [];
      var auditIdIndex = d.headers.indexOf("Audit ID");
      if (auditIdIndex === -1) {
        d.headers.push("Audit ID");
        d.row.push(auditId);
      } else {
        while (d.row.length <= auditIdIndex) d.row.push("");
        d.row[auditIdIndex] = auditId;
      }
      var sheetHeaders = ensureHeaders(sheet, d.headers);
      duplicateAudit = findRecentAuditIdRow_(sheet, auditId) !== -1;
      if (duplicateAudit) {
        diag += ' | DuplicateAudit=SKIPPED (' + auditId + ')';
      }

      // Upload files to Google Drive (if any and folder is resolved)
      if (!duplicateAudit && attachments.length > 0 && agentFolder) {
        for (var j = 0; j < attachments.length; j++) {
          try {
            var blob = attachments[j];
            var file = agentFolder.createFile(blob);
            fileUrls.push(file.getUrl());
            diag += ' | File' + (j+1) + '=DriveOK';
          } catch (uploadErr) {
            diag += ' | File' + (j+1) + '=DriveFAIL:' + uploadErr.toString();
          }
        }
      }

      fileUrlStr = fileUrls.length > 0 ? fileUrls.join("\n") : "No files attached";
      diag += ' | Drive URLs: ' + fileUrls.length;

      // --- Write Drive links to sheet ---
      if (!duplicateAudit) {
        for (var i = 0; i < d.row.length; i++) {
          if (d.row[i] === "[UPLOADING...]") d.row[i] = fileUrlStr;
        }
        var appendLock = LockService.getScriptLock();
        appendLock.waitLock(30000);
        try {
          duplicateAudit = findRecentAuditIdRow_(sheet, auditId) !== -1;
          if (!duplicateAudit) {
            sheet.appendRow(alignRowToHeaders_(sheetHeaders, d.headers, d.row));
            markAuditDataChanged_();
          } else {
            diag += ' | ConcurrentDuplicate=SKIPPED (' + auditId + ')';
          }
        } finally {
          appendLock.releaseLock();
        }
      }
    }

    // --- Send Email ---
    if (!duplicateAudit && d.emailSettings && d.emailSettings.to && d.emailSettings.send !== false) {
      var body = d.emailSettings.body;

      if (agentFolderUrl) {
        var folderLinkHtml = '<a href="' + agentFolderUrl + '">All Audited Calls Folder</a>';
        body = body.replace(/\[Call Recording Link\]/gi, folderLinkHtml);
      } else {
        body = body.replace(/\[Call Recording Link\]/gi, 'No folder link available (no recordings uploaded).');
      }

      var cc = d.emailSettings.cc || '';
      if (cc.indexOf('teresia.nyokabi@food4education.org') === -1) {
        cc += (cc ? ',' : '') + 'teresia.nyokabi@food4education.org';
      }

      var mailOptions = {
        to: d.emailSettings.to,
        subject: d.emailSettings.subject,
        htmlBody: body,
        cc: cc
      };

      if (attachments && attachments.length > 0) {
        mailOptions.attachments = attachments;
      }

      MailApp.sendEmail(mailOptions);

      diag += ' | Email sent';
    }

    return ContentService
      .createTextOutput(JSON.stringify({ status: "success", auditId: auditId, debug: diag }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: "error", msg: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function doGet(e) {
  // If there's an action query parameter, handle API calls
  if (e && e.parameter && e.parameter.action) {
    if (e.parameter.action === "getTrackerData") {
      return ContentService
        .createTextOutput(JSON.stringify(getTrackerData()))
        .setMimeType(ContentService.MimeType.JSON);
    }
    if (e.parameter.action === "getAuditsData") {
      return ContentService
        .createTextOutput(JSON.stringify(getAuditsData(e.parameter.type, e.parameter.limit)))
        .setMimeType(ContentService.MimeType.JSON);
    }
    if (e.parameter.action === "getAuditVersion") {
      return ContentService
        .createTextOutput(JSON.stringify({ version: getAuditDataVersion_() }))
        .setMimeType(ContentService.MimeType.JSON);
    }
  }
  
  try {
    // Dynamic Loader: Fetch HTML directly from GitHub raw link (updates instantly on push!)
    var url = "https://raw.githubusercontent.com/nishkib09/qadev-hub/main/coaching_training_tracker.html";
    var htmlContent = UrlFetchApp.fetch(url).getContentText();
    return HtmlService.createHtmlOutput(htmlContent)
        .setTitle('Tap2eat | Q&T Coaching & Training Hub')
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
        .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  } catch(err) {
    // Fallback: Serve the local copy in case of network issues or missing branch
    Logger.log("doGet dynamic load failed, falling back: " + err.toString());
    return HtmlService.createTemplateFromFile('coaching_training_tracker')
        .evaluate()
        .setTitle('Tap2eat | Q&T Coaching & Training Hub (Fallback)')
        .setSandboxMode(HtmlService.SandboxMode.IFRAME)
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
        .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  }
}

function getAuditsData(sheetName, requestedLimit) {
  if (!sheetName) return [];
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) return [];
  var lastRow = sheet.getLastRow();
  var lastColumn = sheet.getLastColumn();
  if (lastRow < 1 || lastColumn < 1) return [];

  var parsedLimit = parseInt(requestedLimit, 10);
  var limit = isNaN(parsedLimit) ? MAX_DASHBOARD_AUDIT_ROWS : parsedLimit;
  limit = Math.max(1, Math.min(limit, MAX_DASHBOARD_AUDIT_ROWS));

  var headers = sheet.getRange(1, 1, 1, lastColumn).getDisplayValues()[0];
  if (lastRow === 1) return [headers];

  var rowCount = Math.min(limit, lastRow - 1);
  var startRow = lastRow - rowCount + 1;
  var rows = sheet.getRange(startRow, 1, rowCount, lastColumn).getDisplayValues();
  return [headers].concat(rows);
}

function markAuditDataChanged_() {
  PropertiesService.getScriptProperties().setProperty(
    AUDIT_VERSION_PROPERTY,
    String(Date.now()) + "-" + Utilities.getUuid()
  );
}

function getAuditDataVersion_() {
  return PropertiesService.getScriptProperties().getProperty(AUDIT_VERSION_PROPERTY) || "0";
}

function findRecentAuditIdRow_(sheet, auditId) {
  if (!sheet || !auditId || sheet.getLastRow() < 2) return -1;
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var auditIdCol = headers.indexOf("Audit ID") + 1;
  if (auditIdCol === 0) return -1;
  var lastRow = sheet.getLastRow();
  var startRow = Math.max(2, lastRow - 9999);
  var values = sheet.getRange(startRow, auditIdCol, lastRow - startRow + 1, 1).getDisplayValues();
  for (var i = values.length - 1; i >= 0; i--) {
    if (String(values[i][0]) === auditId) return startRow + i;
  }
  return -1;
}

// One-time migration for legacy rows. Run this manually after deploying the
// script; it fills only blank IDs and never changes an existing Audit ID.
function backfillAuditIds() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetNames = [
    "Inbound Scorecard",
    "Outbound Call Scorecard (Campaign-Focused)",
    "Consumer Satisfaction Survey Calls",
    "SMS Quality Audit",
    "Email Audit Scorecard"
  ];
  var chunkSize = 5000;
  var updated = 0;

  sheetNames.forEach(function(sheetName) {
    var sheet = ss.getSheetByName(sheetName);
    if (!sheet || sheet.getLastRow() < 2) return;

    var lastCol = sheet.getLastColumn();
    var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
    var auditIdCol = headers.indexOf("Audit ID") + 1;
    if (auditIdCol === 0) {
      auditIdCol = lastCol + 1;
      sheet.getRange(1, auditIdCol).setValue("Audit ID");
    }

    var lastRow = sheet.getLastRow();
    for (var startRow = 2; startRow <= lastRow; startRow += chunkSize) {
      var rowCount = Math.min(chunkSize, lastRow - startRow + 1);
      var range = sheet.getRange(startRow, auditIdCol, rowCount, 1);
      var ids = range.getValues();
      var changed = false;
      for (var i = 0; i < ids.length; i++) {
        if (!ids[i][0]) {
          ids[i][0] = Utilities.getUuid();
          updated++;
          changed = true;
        }
      }
      if (changed) range.setValues(ids);
    }
  });

  if (updated > 0) markAuditDataChanged_();
  return { updated: updated };
}

// Database Sheet names
var SHEET_COACHING = "CoachingLog";
var SHEET_TRAINING = "TrainingLog";
var SHEET_PLAN = "ExecutionPlan";
var SHEET_SETTINGS = "TrackerSettings";
var COACHING_LOG_HEADERS = [
  "ID", "Date", "Agent Name", "Coach Name", "Type",
  "Focus Area", "Pre-Coaching Score (%)", "Post-Coaching Score (%)",
  "Observations", "Action Items", "Follow-up Date", "Status",
  "Created At", "Updated At", "Agent Email", "Source Audit ID",
  "Source Audit Type", "Source Audit Date", "Session Time", "Follow-up Notes"
];
var TRAINING_LOG_HEADERS = [
  "ID", "Date", "Training Title", "Trainer Name", "Type",
  "Topic/Module", "Duration (Hrs)", "Attendees Count", "Pre-Assessment Avg (%)",
  "Post-Assessment Avg (%)", "Delivery Method", "Status", "Notes",
  "Created At", "Updated At", "Start Time", "Attendee Emails"
];

function ensureAppendOnlyHeaders_(sheet, requiredHeaders) {
  if (!sheet || !requiredHeaders || requiredHeaders.length === 0) return;
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, requiredHeaders.length).setValues([requiredHeaders]);
    return;
  }
  var lastCol = Math.max(sheet.getLastColumn(), 1);
  var currentHeaders = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  var missing = requiredHeaders.filter(function(header) {
    return currentHeaders.indexOf(header) === -1;
  });
  if (missing.length > 0) {
    sheet.getRange(1, lastCol + 1, 1, missing.length).setValues([missing]);
  }
}

// Auto-initialize sheets
function initTrackerSheets() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // 1. Coaching Log Sheet
  var coachingSheet = ss.getSheetByName(SHEET_COACHING);
  if (!coachingSheet) {
    coachingSheet = ss.insertSheet(SHEET_COACHING);
    coachingSheet.appendRow(COACHING_LOG_HEADERS);
  } else {
    ensureAppendOnlyHeaders_(coachingSheet, COACHING_LOG_HEADERS);
  }
  
  // 2. Training Log Sheet
  var trainingSheet = ss.getSheetByName(SHEET_TRAINING);
  if (!trainingSheet) {
    trainingSheet = ss.insertSheet(SHEET_TRAINING);
    trainingSheet.appendRow(TRAINING_LOG_HEADERS);
  } else {
    ensureAppendOnlyHeaders_(trainingSheet, TRAINING_LOG_HEADERS);
  }
  
  // 3. Execution Plan Sheet
  var planSheet = ss.getSheetByName(SHEET_PLAN);
  if (!planSheet) {
    planSheet = ss.insertSheet(SHEET_PLAN);
    planSheet.appendRow([
      "ID", "Date", "Type", "Title", "Assigned Agent", 
      "Topic", "Time Slot", "Status", "Recurring", "Linked Session ID"
    ]);
  }

  // 4. Settings Sheet
  var settingsSheet = ss.getSheetByName(SHEET_SETTINGS);
  if (!settingsSheet) {
    settingsSheet = ss.insertSheet(SHEET_SETTINGS);
    settingsSheet.appendRow(["Key", "Value"]);
    settingsSheet.appendRow(["Coaches", "Jane Doe, John Smith, Mercy Wanjiku"]);
    settingsSheet.appendRow(["FocusAreas", "Empathy, Active Listening, SOP Compliance, FCR, Product Knowledge, Call Control, System Navigation"]);
    settingsSheet.appendRow(["TrainingTopics", "New Onboarding, System Refresher, Empathy Deep-Dive, Escalation Handling, De-escalation Skills"]);
    settingsSheet.appendRow(["CoachingCcEmails", "teresia.nyokabi@food4education.org"]);
  } else {
    var data = settingsSheet.getDataRange().getValues();
    var hasCcSetting = false;
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] === "CoachingCcEmails") {
        hasCcSetting = true;
        break;
      }
    }
    if (!hasCcSetting) {
      settingsSheet.appendRow(["CoachingCcEmails", "teresia.nyokabi@food4education.org"]);
    }
  }

  // 5. Agent Directory Sheet
  var agentSheet = ss.getSheetByName("AgentDirectory");
  if (!agentSheet) {
    agentSheet = ss.insertSheet("AgentDirectory");
    agentSheet.appendRow(["Agent Name", "Email Address", "Team", "Status"]);
    agentSheet.appendRow(["Adah Achieng", "adah.achieng@food4education.org", "CX Team", "Active"]);
    agentSheet.appendRow(["Branice Arakot", "branice.mukasa@food4education.org", "CX Team", "Active"]);
    agentSheet.appendRow(["Brian Gichuhi", "brian.gichuhi@food4education.org", "CX Team", "Active"]);
    agentSheet.appendRow(["Eddy Wanjiku", "eddy.gathogo@food4education.org", "CX Team", "Active"]);
    agentSheet.appendRow(["Gerald Macharia", "gerald.macharia@food4education.org", "CX Team", "Active"]);
    agentSheet.appendRow(["Janipher Achieng", "janipher@food4education.org", "CX Team", "Active"]);
    agentSheet.appendRow(["Joan Wachira", "joan.wachira@food4education.org", "CX Team", "Active"]);
    agentSheet.appendRow(["John Gitungi", "john.gitungi@food4education.org", "CX Team", "Active"]);
    agentSheet.appendRow(["Lilian Kemunto", "lilian.kemunto@food4education.org", "CX Team", "Active"]);
    agentSheet.appendRow(["Magdalene Mukami", "magdaline.njeru@food4education.org", "CX Team", "Active"]);
    agentSheet.appendRow(["Nancy Waweru", "nancy.waweru@food4education.org", "CX Team", "Active"]);
    agentSheet.appendRow(["Peter Mwangi", "peter.mwangi@food4education.org", "CX Team", "Active"]);
    agentSheet.appendRow(["Rahab Munyua", "rahab.munyua@food4education.org", "CX Team", "Active"]);
  }
}

// Fetch all database tables
function getTrackerData() {
  try {
    initTrackerSheets();
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    return {
      coaching: getSheetData(ss.getSheetByName(SHEET_COACHING), MAX_TRACKER_LOG_ROWS),
      training: getSheetData(ss.getSheetByName(SHEET_TRAINING), MAX_TRACKER_LOG_ROWS),
      plans: getSheetData(ss.getSheetByName(SHEET_PLAN), MAX_TRACKER_LOG_ROWS),
      agents: getSheetData(ss.getSheetByName("AgentDirectory")),
      settings: getSettingsData(ss.getSheetByName(SHEET_SETTINGS)),
      scorecards: getExternalScorecardData(),
      needsFeedback: {
        needs: getSheetData(ss.getSheetByName("PreTrainingNeeds"), MAX_TRACKER_LOG_ROWS),
        feedback: getSheetData(ss.getSheetByName("PostTrainingFeedback"), MAX_TRACKER_LOG_ROWS)
      },
      deescalation: getDeescalationData()
    };
  } catch(err) {
    Logger.log("getTrackerData error: " + err.toString());
    return { 
      coaching: [], training: [], plans: [], agents: [], settings: {}, 
      scorecards: { inbound: [], outbound: [], survey: [], sms: [], email: [] },
      needsFeedback: { needs: [], feedback: [] },
      deescalation: [],
      error: err.toString() 
    };
  }
}

function getExternalScorecardData() {
  var data = { inbound: [], outbound: [], survey: [], sms: [], email: [] };
  try {
    var extSS = SpreadsheetApp.openById("1VIAYlOvb5TtonUSAEiECqrq7i-n_dFA70IkmRW9dQIo");
    var sheets = extSS.getSheets();
    
    for (var i = 0; i < sheets.length; i++) {
      var sheetName = sheets[i].getName();
      var sheetId = sheets[i].getSheetId();
      
      // Match by exact sheet ID or name
      if (sheetName === "Inbound Scorecard" || sheetId == 956377943) {
        data.inbound = getSheetData(sheets[i], MAX_TRACKER_AUDIT_ROWS_PER_TYPE);
      } else if (sheetName === "SMS Quality Audit" || sheetId == 283607713) {
        data.sms = getSheetData(sheets[i], MAX_TRACKER_AUDIT_ROWS_PER_TYPE);
      } else if (sheetName === "Outbound Call Scorecard (Campaign-Focused)" || sheetId == 1574188088) {
        data.outbound = getSheetData(sheets[i], MAX_TRACKER_AUDIT_ROWS_PER_TYPE);
      } else if (sheetName === "Consumer Satisfaction Survey Calls") {
        data.survey = getSheetData(sheets[i], MAX_TRACKER_AUDIT_ROWS_PER_TYPE);
      } else if (sheetName === "Email Audit Scorecard") {
        data.email = getSheetData(sheets[i], MAX_TRACKER_AUDIT_ROWS_PER_TYPE);
      }
    }
  } catch(e) {
    Logger.log("getExternalScorecardData error: " + e.toString());
  }
  
  // Fallbacks: If external SS fails to load, try active SS
  try {
    var localSS = SpreadsheetApp.getActiveSpreadsheet();
    if (data.inbound.length === 0) {
      var s = localSS.getSheetByName("Inbound Scorecard");
      if (s) data.inbound = getSheetData(s, MAX_TRACKER_AUDIT_ROWS_PER_TYPE);
    }
    if (data.sms.length === 0) {
      var s = localSS.getSheetByName("SMS Quality Audit");
      if (s) data.sms = getSheetData(s, MAX_TRACKER_AUDIT_ROWS_PER_TYPE);
    }
    if (data.outbound.length === 0) {
      var s = localSS.getSheetByName("Outbound Call Scorecard (Campaign-Focused)");
      if (s) data.outbound = getSheetData(s, MAX_TRACKER_AUDIT_ROWS_PER_TYPE);
    }
    if (data.survey.length === 0) {
      var surveySheet = localSS.getSheetByName("Consumer Satisfaction Survey Calls");
      if (surveySheet) data.survey = getSheetData(surveySheet, MAX_TRACKER_AUDIT_ROWS_PER_TYPE);
    }
    if (data.email.length === 0) {
      var emailSheet = localSS.getSheetByName("Email Audit Scorecard");
      if (emailSheet) data.email = getSheetData(emailSheet, MAX_TRACKER_AUDIT_ROWS_PER_TYPE);
    }
  } catch(localErr) {
    Logger.log("getExternalScorecardData local fallback error: " + localErr.toString());
  }
  
  return data;
}

function getDeescalationData() {
  try {
    var ss = SpreadsheetApp.openById("1vaEX7f0ayRKryBy_C5NcL2DMncKzJ2CpxpaA-p-JbL0");
    var sheets = ss.getSheets();
    var sheet = sheets[0];
    for (var i = 0; i < sheets.length; i++) {
      if (sheets[i].getSheetId() == 632483075) {
        sheet = sheets[i];
        break;
      }
    }
    return getSheetData(sheet, MAX_TRACKER_LOG_ROWS);
  } catch(e) {
    Logger.log("getDeescalationData error: " + e.toString());
    return [];
  }
}

// Setup and link Google Forms for Pre-Training Needs and Post-Training Feedback
function setupGoogleForms(customConfig) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Use customConfig or read from Settings or default to CX
  var titlePre = (customConfig && customConfig.titlePre) || getSettingValue("FormTitlePre") || "CX | Pre-Training Needs Assessment";
  var descPre = (customConfig && customConfig.descPre) || getSettingValue("FormDescPre") || "Please share your current knowledge level and learning expectations before the training session.";
  var titlePost = (customConfig && customConfig.titlePost) || getSettingValue("FormTitlePost") || "CX | Post-Training Feedback & Confidence Evaluation";
  var descPost = (customConfig && customConfig.descPost) || getSettingValue("FormDescPost") || "Please rate your confidence and provide feedback on the training session you completed.";
  
  // Save settings to spreadsheet so they persist
  saveSettingValue("FormTitlePre", titlePre);
  saveSettingValue("FormDescPre", descPre);
  saveSettingValue("FormTitlePost", titlePost);
  saveSettingValue("FormDescPost", descPost);

  // 1. Pre-Training Needs Form
  var formIdPre = getSettingValue("FormIdPre");
  var formPre;
  if (!formIdPre) {
    formPre = FormApp.create(titlePre);
    formPre.setDescription(descPre);
    
    var agentItem = formPre.addListItem();
    agentItem.setTitle("Agent Name");
    
    var topicItem = formPre.addListItem();
    topicItem.setTitle("Training Topic");
    
    var scaleItem = formPre.addScaleItem();
    scaleItem.setTitle("Rate your current knowledge/confidence in this topic");
    scaleItem.setBounds(1, 5);
    scaleItem.setLabels("Beginner / No Knowledge", "Expert / Highly Confident");
    
    var prevCoachedItem = formPre.addMultipleChoiceItem();
    prevCoachedItem.setTitle("Have you received formal coaching or training on this topic before?");
    prevCoachedItem.setChoiceValues(["Yes, within the last 3 months", "Yes, more than 3 months ago", "No, this is my first time"]);

    var textItem = formPre.addParagraphTextItem();
    textItem.setTitle("What specific areas, questions, or scenarios would you like this training to cover?");
    
    formPre.setDestination(FormApp.DestinationType.SPREADSHEET, ss.getId());
    formIdPre = formPre.getId();
    
    saveSettingValue("FormIdPre", formIdPre);
    saveSettingValue("FormUrlPre", formPre.getPublishedUrl());
  } else {
    try {
      formPre = FormApp.openById(formIdPre);
      formPre.setTitle(titlePre);
      formPre.setDescription(descPre);
    } catch(e) {
      formIdPre = null;
      // Re-create if opening existing failed
      formPre = FormApp.create(titlePre);
      formPre.setDescription(descPre);
      
      var agentItem = formPre.addListItem();
      agentItem.setTitle("Agent Name");
      
      var topicItem = formPre.addListItem();
      topicItem.setTitle("Training Topic");
      
      var scaleItem = formPre.addScaleItem();
      scaleItem.setTitle("Rate your current knowledge/confidence in this topic");
      scaleItem.setBounds(1, 5);
      scaleItem.setLabels("Beginner / No Knowledge", "Expert / Highly Confident");
      
      var prevCoachedItem = formPre.addMultipleChoiceItem();
      prevCoachedItem.setTitle("Have you received formal coaching or training on this topic before?");
      prevCoachedItem.setChoiceValues(["Yes, within the last 3 months", "Yes, more than 3 months ago", "No, this is my first time"]);

      var textItem = formPre.addParagraphTextItem();
      textItem.setTitle("What specific areas, questions, or scenarios would you like this training to cover?");
      
      formPre.setDestination(FormApp.DestinationType.SPREADSHEET, ss.getId());
      formIdPre = formPre.getId();
      
      saveSettingValue("FormIdPre", formIdPre);
      saveSettingValue("FormUrlPre", formPre.getPublishedUrl());
    }
  }

  // 2. Post-Training Feedback Form
  var formIdPost = getSettingValue("FormIdPost");
  var formPost;
  if (!formIdPost) {
    formPost = FormApp.create(titlePost);
    formPost.setDescription(descPost);
    
    var agentItem = formPost.addListItem();
    agentItem.setTitle("Agent Name");
    
    var topicItem = formPost.addListItem();
    topicItem.setTitle("Training Topic");
    
    var scaleItem = formPost.addScaleItem();
    scaleItem.setTitle("Rate your knowledge/confidence in this topic AFTER the training");
    scaleItem.setBounds(1, 5);
    scaleItem.setLabels("Low Confidence", "Very High Confidence");
    
    var scaleTrainer = formPost.addScaleItem();
    scaleTrainer.setTitle("Rate the trainer's preparation and delivery");
    scaleTrainer.setBounds(1, 5);
    scaleTrainer.setLabels("Poor", "Excellent");
    
    var scaleContent = formPost.addScaleItem();
    scaleContent.setTitle("Rate the content and training materials");
    scaleContent.setBounds(1, 5);
    scaleContent.setLabels("Poor", "Excellent");
    
    var pacingItem = formPost.addMultipleChoiceItem();
    pacingItem.setTitle("Rate the pacing of the training session");
    pacingItem.setChoiceValues(["Too Slow", "Just Right", "Too Fast"]);

    var applyItem = formPost.addParagraphTextItem();
    applyItem.setTitle("How will you apply what you learned today in your daily calls?");

    var textItem = formPost.addParagraphTextItem();
    textItem.setTitle("What recommendations, feedback, or future training topics do you suggest?");
    
    formPost.setDestination(FormApp.DestinationType.SPREADSHEET, ss.getId());
    formIdPost = formPost.getId();
    
    saveSettingValue("FormIdPost", formIdPost);
    saveSettingValue("FormUrlPost", formPost.getPublishedUrl());
  } else {
    try {
      formPost = FormApp.openById(formIdPost);
      formPost.setTitle(titlePost);
      formPost.setDescription(descPost);
    } catch(e) {
      formIdPost = null;
      // Re-create if opening existing failed
      formPost = FormApp.create(titlePost);
      formPost.setDescription(descPost);
      
      var agentItem = formPost.addListItem();
      agentItem.setTitle("Agent Name");
      
      var topicItem = formPost.addListItem();
      topicItem.setTitle("Training Topic");
      
      var scaleItem = formPost.addScaleItem();
      scaleItem.setTitle("Rate your knowledge/confidence in this topic AFTER the training");
      scaleItem.setBounds(1, 5);
      scaleItem.setLabels("Low Confidence", "Very High Confidence");
      
      var scaleTrainer = formPost.addScaleItem();
      scaleTrainer.setTitle("Rate the trainer's preparation and delivery");
      scaleTrainer.setBounds(1, 5);
      scaleTrainer.setLabels("Poor", "Excellent");
      
      var scaleContent = formPost.addScaleItem();
      scaleContent.setTitle("Rate the content and training materials");
      scaleContent.setBounds(1, 5);
      scaleContent.setLabels("Poor", "Excellent");
      
      var pacingItem = formPost.addMultipleChoiceItem();
      pacingItem.setTitle("Rate the pacing of the training session");
      pacingItem.setChoiceValues(["Too Slow", "Just Right", "Too Fast"]);

      var applyItem = formPost.addParagraphTextItem();
      applyItem.setTitle("How will you apply what you learned today in your daily calls?");

      var textItem = formPost.addParagraphTextItem();
      textItem.setTitle("What recommendations, feedback, or future training topics do you suggest?");
      
      formPost.setDestination(FormApp.DestinationType.SPREADSHEET, ss.getId());
      formIdPost = formPost.getId();
      
      saveSettingValue("FormIdPost", formIdPost);
      saveSettingValue("FormUrlPost", formPost.getPublishedUrl());
    }
  }
  
  // Update Dropdown Choices dynamically from Spreadsheet data
  syncFormChoices();
  
  // Find newly inserted response sheets and rename them
  renameFormResponseSheets(ss);
  
  return {
    preUrl: getSettingValue("FormUrlPre"),
    postUrl: getSettingValue("FormUrlPost")
  };
}

function renameFormResponseSheets(ss) {
  var sheets = ss.getSheets();
  for (var i = 0; i < sheets.length; i++) {
    var s = sheets[i];
    var name = s.getName();
    if (name.indexOf("Form Responses") !== -1 || name.indexOf("Pre-Training") !== -1 || name.indexOf("Post-Training") !== -1) {
      if (s.getLastColumn() > 0) {
        var headers = s.getRange(1, 1, 1, s.getLastColumn()).getValues()[0];
        var headersStr = headers.join(" ");
        if (headersStr.indexOf("learning needs") !== -1 || headersStr.indexOf("specific areas") !== -1) {
          if (name !== "PreTrainingNeeds") {
            try {
              var old = ss.getSheetByName("PreTrainingNeeds");
              if (old) ss.deleteSheet(old);
              s.setName("PreTrainingNeeds");
            } catch(e) {}
          }
        } else if (headersStr.indexOf("trainer's preparation") !== -1 || headersStr.indexOf("training materials") !== -1) {
          if (name !== "PostTrainingFeedback") {
            try {
              var old = ss.getSheetByName("PostTrainingFeedback");
              if (old) ss.deleteSheet(old);
              s.setName("PostTrainingFeedback");
            } catch(e) {}
          }
        }
      }
    }
  }
}

function syncFormChoices() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Get Agents list
  var agentSheet = ss.getSheetByName("AgentDirectory");
  var agents = [];
  if (agentSheet) {
    var data = agentSheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] && data[i][3] === "Active") {
        agents.push(data[i][0]);
      }
    }
  }
  if (agents.length === 0) agents = ["Adah Achieng", "Branice Arakot", "Brian Gichuhi", "Joan Wachira"];
  
  // Get Topics list
  var settingsSheet = ss.getSheetByName(SHEET_SETTINGS);
  var topics = [];
  if (settingsSheet) {
    var data = settingsSheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] === "TrainingTopics") {
        var val = data[i][1];
        if (val) {
          topics = val.split(",").map(function(s) { return s.trim(); });
        }
        break;
      }
    }
  }
  if (topics.length === 0) topics = ["New Onboarding", "System Refresher", "Empathy Deep-Dive", "De-escalation Skills"];
  
  // Update dropdown choices in Pre Form
  var formIdPre = getSettingValue("FormIdPre");
  if (formIdPre) {
    try {
      var form = FormApp.openById(formIdPre);
      var items = form.getItems();
      for (var i = 0; i < items.length; i++) {
        var item = items[i];
        if (item.getTitle() === "Agent Name") {
          item.asListItem().setChoiceValues(agents);
        } else if (item.getTitle() === "Training Topic") {
          item.asListItem().setChoiceValues(topics);
        }
      }
    } catch(e) {}
  }
  
  // Update dropdown choices in Post Form
  var formIdPost = getSettingValue("FormIdPost");
  if (formIdPost) {
    try {
      var form = FormApp.openById(formIdPost);
      var items = form.getItems();
      for (var i = 0; i < items.length; i++) {
        var item = items[i];
        if (item.getTitle() === "Agent Name") {
          item.asListItem().setChoiceValues(agents);
        } else if (item.getTitle() === "Training Topic") {
          item.asListItem().setChoiceValues(topics);
        }
      }
    } catch(e) {}
  }
}

function getPrefilledFormUrl(topic, agentName, type) {
  var formIdKey = (type === 'pre') ? "FormIdPre" : "FormIdPost";
  var formId = getSettingValue(formIdKey);
  
  if (!formId) {
    setupGoogleForms();
    formId = getSettingValue(formIdKey);
  }
  
  if (!formId) return "";
  
  try {
    var form = FormApp.openById(formId);
    var response = form.createResponse();
    var items = form.getItems();
    
    for (var i = 0; i < items.length; i++) {
      var item = items[i];
      var title = item.getTitle();
      if (title === "Agent Name" && agentName) {
        var agentResponse = item.asListItem().createResponse(agentName);
        response.withItemResponse(agentResponse);
      } else if (title === "Training Topic" && topic) {
        var topicResponse = item.asListItem().createResponse(topic);
        response.withItemResponse(topicResponse);
      }
    }
    
    return response.toPrefilledUrl();
  } catch(e) {
    Logger.log("getPrefilledFormUrl error: " + e.toString());
    return "";
  }
}

function getSettingValue(key) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_SETTINGS);
  if (!sheet) return "";
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === key) {
      return data[i][1];
    }
  }
  return "";
}

function saveSettingValue(key, val) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_SETTINGS);
  if (!sheet) return;
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === key) {
      sheet.getRange(i + 1, 2).setValue(val);
      return;
    }
  }
  sheet.appendRow([key, val]);
}

function getSheetData(sheet, requestedLimit) {
  if (!sheet) return [];
  var lastRow = sheet.getLastRow();
  var lastColumn = sheet.getLastColumn();
  if (lastRow <= 1 || lastColumn < 1) return [];
  var headers = sheet.getRange(1, 1, 1, lastColumn).getValues()[0];
  var parsedLimit = parseInt(requestedLimit, 10);
  var rowCount = isNaN(parsedLimit) ? lastRow - 1 : Math.min(Math.max(parsedLimit, 1), lastRow - 1);
  var startRow = lastRow - rowCount + 1;
  var data = sheet.getRange(startRow, 1, rowCount, lastColumn).getValues();
  var rows = [];
  for (var i = 0; i < data.length; i++) {
    var obj = {};
    for (var j = 0; j < headers.length; j++) {
      var key = toCamelCase(headers[j]);
      var val = data[i][j];
      if (val instanceof Date) {
        obj[key] = Utilities.formatDate(val, Session.getScriptTimeZone() || "GMT", "yyyy-MM-dd");
      } else {
        obj[key] = val;
      }
    }
    rows.push(obj);
  }
  return rows;
}

function getSettingsData(sheet) {
  if (!sheet) return {};
  var data = sheet.getDataRange().getValues();
  var settings = {};
  for (var i = 1; i < data.length; i++) {
    var key = data[i][0];
    var val = data[i][1];
    if (key === "Agents" || key === "Coaches" || key === "FocusAreas" || key === "TrainingTopics") {
      settings[key] = val ? val.split(",").map(function(s) { return s.trim(); }) : [];
    } else {
      settings[key] = val;
    }
  }
  return settings;
}

function toCamelCase(str) {
  return str.replace(/[^a-zA-Z0-9 ]/g, "")
            .toLowerCase()
            .replace(/[^a-zA-Z0-9]+(.)/g, function(m, chr) {
              return chr.toUpperCase();
            });
}

// Write/Update functions
function saveCoachingSession(session) {
  return saveRow(SHEET_COACHING, session, COACHING_LOG_HEADERS);
}

function saveTrainingSession(training) {
  return saveRow(SHEET_TRAINING, training, TRAINING_LOG_HEADERS);
}

function saveExecutionPlanItem(plan) {
  return saveRow(SHEET_PLAN, plan, [
    "ID", "Date", "Type", "Title", "Assigned Agent", 
    "Topic", "Time Slot", "Status", "Recurring", "Linked Session ID"
  ]);
}

function deleteRowItem(sheetName, id) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) return false;
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] == id) {
      sheet.deleteRow(i + 1);
      return true;
    }
  }
  return false;
}

function saveSettingsData(settings) {
  initTrackerSheets();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_SETTINGS);
  sheet.clear();
  sheet.appendRow(["Key", "Value"]);
  for (var key in settings) {
    var val = settings[key];
    if (Array.isArray(val)) {
      val = val.join(", ");
    }
    sheet.appendRow([key, val]);
  }
  return true;
}

function saveRow(sheetName, item, headers) {
  initTrackerSheets();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(sheetName);
  var data = sheet.getDataRange().getValues();
  var foundRowIndex = -1;
  
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] == item.id) {
      foundRowIndex = i + 1;
      break;
    }
  }
  
  var rowValues = [];
  for (var h = 0; h < headers.length; h++) {
    var key = toCamelCase(headers[h]);
    var val = item[key];
    if (val === undefined || val === null) {
      if (headers[h] === "Created At") val = new Date();
      else if (headers[h] === "Updated At") val = new Date();
      else val = "";
    }
    rowValues.push(val);
  }
  
  if (foundRowIndex !== -1) {
    var createdAtCol = headers.indexOf("Created At");
    if (createdAtCol !== -1 && data[foundRowIndex - 1][createdAtCol]) {
      rowValues[createdAtCol] = data[foundRowIndex - 1][createdAtCol];
    }
    var updatedAtCol = headers.indexOf("Updated At");
    if (updatedAtCol !== -1) {
      rowValues[updatedAtCol] = new Date();
    }
    sheet.getRange(foundRowIndex, 1, 1, rowValues.length).setValues([rowValues]);
  } else {
    sheet.appendRow(rowValues);
  }
  return true;
}


function ensureHeaders(sheet, expectedHeaders) {
  if (!expectedHeaders || expectedHeaders.length === 0) return [];
  var lastCol = sheet.getLastColumn();
  var currentHeaders = lastCol > 0
    ? sheet.getRange(1, 1, 1, lastCol).getDisplayValues()[0]
    : [];

  while (currentHeaders.length && !String(currentHeaders[currentHeaders.length - 1]).trim()) {
    currentHeaders.pop();
  }

  if (currentHeaders.length === 0) {
    if (sheet.getMaxColumns() < expectedHeaders.length) {
      sheet.insertColumnsAfter(sheet.getMaxColumns(), expectedHeaders.length - sheet.getMaxColumns());
    }
    sheet.getRange(1, 1, 1, expectedHeaders.length).setValues([expectedHeaders]);
    return expectedHeaders.slice();
  }

  var missingHeaders = expectedHeaders.filter(function(header) {
    return currentHeaders.indexOf(header) === -1;
  });
  if (missingHeaders.length > 0) {
    var requiredColumns = currentHeaders.length + missingHeaders.length;
    if (sheet.getMaxColumns() < requiredColumns) {
      sheet.insertColumnsAfter(sheet.getMaxColumns(), requiredColumns - sheet.getMaxColumns());
    }
    sheet.getRange(1, currentHeaders.length + 1, 1, missingHeaders.length).setValues([missingHeaders]);
    currentHeaders = currentHeaders.concat(missingHeaders);
  }
  return currentHeaders;
}

function alignRowToHeaders_(sheetHeaders, payloadHeaders, payloadRow) {
  var valuesByHeader = {};
  (payloadHeaders || []).forEach(function(header, index) {
    if (Object.prototype.hasOwnProperty.call(valuesByHeader, header)) {
      throw new Error('Duplicate audit header: ' + header);
    }
    valuesByHeader[header] = index < (payloadRow || []).length ? payloadRow[index] : '';
  });
  return (sheetHeaders || []).map(function(header) {
    return Object.prototype.hasOwnProperty.call(valuesByHeader, header)
      ? valuesByHeader[header]
      : '';
  });
}

// Agent Directory CRUD
function saveAgent(agent) {
  initTrackerSheets();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("AgentDirectory");
  var data = sheet.getDataRange().getValues();
  var foundRowIndex = -1;
  
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] == agent.agentName) {
      foundRowIndex = i + 1;
      break;
    }
  }
  
  var rowValues = [
    agent.agentName, 
    agent.emailAddress, 
    agent.team || "CX Team", 
    agent.status || "Active"
  ];
  
  if (foundRowIndex !== -1) {
    sheet.getRange(foundRowIndex, 1, 1, rowValues.length).setValues([rowValues]);
  } else {
    sheet.appendRow(rowValues);
  }
  return true;
}

function deleteAgent(name) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("AgentDirectory");
  if (!sheet) return false;
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] == name) {
      sheet.deleteRow(i + 1);
      return true;
    }
  }
  return false;
}

// Email coaching discussion notes / summary directly to the agent
function sendCoachingSummaryEmail(session, agentEmail) {
  if (!agentEmail) return false;
  
  // Retrieve CC emails from TrackerSettings sheet
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var settingsSheet = ss.getSheetByName(SHEET_SETTINGS);
  var ccEmails = "teresia.nyokabi@food4education.org"; // Default fallback
  
  if (settingsSheet) {
    var data = settingsSheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] === "CoachingCcEmails") {
        var val = data[i][1];
        if (val) {
          ccEmails = val;
        }
        break;
      }
    }
  }

  var subject = "Coaching Session Summary - " + session.focusArea + " - " + session.date;
  
  // Format HTML body beautifully for Food4Education
  var htmlBody = 
    "<div style='font-family: Georgia, serif; max-width: 600px; margin: 0 auto; padding: 20px; border: 1px solid #e4e2d8; border-radius: 12px; background-color: #f4faf3;'>" +
      "<div style='background-color: #24631E; color: white; padding: 15px 20px; border-radius: 8px 8px 0 0; text-align: center;'>" +
        "<h2 style='margin: 0; font-size: 20px;'>Coaching Discussion Summary</h2>" +
        "<p style='margin: 5px 0 0 0; font-size: 13px; opacity: 0.8;'>Food4Education | Tap2eat Q&T Operations</p>" +
      "</div>" +
      "<div style='padding: 20px; background-color: white; border-radius: 0 0 8px 8px; color: #1F1B1B;'>" +
        "<p>Hi <strong>" + session.agentName + "</strong>,</p>" +
        "<p>Here is a summary of the coaching session conducted on <strong>" + session.date + "</strong> by <strong>" + session.coachName + "</strong>.</p>" +
        
        "<table style='width: 100%; border-collapse: collapse; margin: 20px 0; font-size: 13px;'>" +
          "<tr style='background-color: #f7f5ee;'><td style='padding: 8px; font-weight: bold; width: 140px;'>Focus Area</td><td style='padding: 8px; color: #24631E; font-weight: bold;'>" + session.focusArea + "</td></tr>" +
          "<tr><td style='padding: 8px; font-weight: bold;'>Coaching Type</td><td style='padding: 8px;'>" + session.type + "</td></tr>" +
          "<tr style='background-color: #f7f5ee;'><td style='padding: 8px; font-weight: bold;'>Pre-Coaching Score</td><td style='padding: 8px;'>" + (session.precoachingScore ? session.precoachingScore + "%" : "-") + "</td></tr>" +
          "<tr><td style='padding: 8px; font-weight: bold;'>Post-Coaching Score</td><td style='padding: 8px;'>" + (session.postcoachingScore ? session.postcoachingScore + "%" : "-") + "</td></tr>" +
          "<tr style='background-color: #f7f5ee;'><td style='padding: 8px; font-weight: bold;'>Follow-up Date</td><td style='padding: 8px;'>" + (session.followupDate || "-") + "</td></tr>" +
        "</table>" +
        
        "<div style='margin-bottom: 20px; padding: 12px; background-color: #fdfcea; border-left: 4px solid #FF8C00; border-radius: 4px; font-size: 13px;'>" +
          "<strong style='color: #BF6900;'>Key Observations:</strong>" +
          "<p style='margin: 6px 0 0 0; line-height: 1.5; white-space: pre-wrap;'>" + (session.observations || "None recorded.") + "</p>" +
        "</div>" +
        
        "<div style='margin-bottom: 20px; padding: 12px; background-color: #eaf3e8; border-left: 4px solid #24631E; border-radius: 4px; font-size: 13px;'>" +
          "<strong style='color: #24631E;'>Action Plan / Commitments:</strong>" +
          "<p style='margin: 6px 0 0 0; line-height: 1.5; white-space: pre-wrap;'>" + (session.actionItems || "None recorded.") + "</p>" +
        "</div>" +
        
        "<p style='font-size: 13px; line-height: 1.5;'>Please review this feedback and work towards implementing the commitments we agreed upon during our session.</p>" +
        "<p style='font-size: 13px; font-weight: bold; margin-top: 25px;'>Keep up the great work!</p>" +
        "<p style='font-size: 12px; color: #7f8c8d; border-top: 1px solid #eeeeee; padding-top: 15px; margin-top: 25px;'>This is an automated Q&T summary notification. If you have questions, please reach out to your Q&T Lead.</p>" +
      "</div>" +
    "</div>";

  MailApp.sendEmail({
    to: agentEmail,
    subject: subject,
    htmlBody: htmlBody,
    cc: ccEmails
  });
  return true;
}

function buildTrackerDateTime_(dateValue, timeValue) {
  if (!dateValue) throw new Error("A session date is required.");
  var dateParts = String(dateValue).split("-");
  var timeParts = String(timeValue || "10:00").split(":");
  return new Date(
    Number(dateParts[0]), Number(dateParts[1]) - 1, Number(dateParts[2]),
    Number(timeParts[0] || 10), Number(timeParts[1] || 0), 0, 0
  );
}

function sendCoachingCalendarInvite(session) {
  if (!session || !session.agentEmail) throw new Error("The selected agent has no email address.");
  var start = buildTrackerDateTime_(session.date, session.sessionTime);
  var end = new Date(start.getTime() + 45 * 60 * 1000);
  var description = [
    "Coaching focus: " + (session.focusArea || ""),
    "Pre-coaching QA score: " + (session.precoachingScore === "" ? "Not recorded" : session.precoachingScore + "%"),
    "Source Audit ID: " + (session.sourceAuditId || "Not linked"),
    "Coach: " + (session.coachName || ""),
    "Action plan will be documented in the Q&T Coaching Hub."
  ].join("\n");
  var event = CalendarApp.getDefaultCalendar().createEvent(
    "Coaching | " + (session.focusArea || "Quality improvement"), start, end,
    { description: description, guests: session.agentEmail, sendInvites: true }
  );
  return { success: true, eventId: event.getId() };
}

function sendTrainingCalendarInvite(training) {
  var guests = String(training.attendeeEmails || "").split(",").map(function(email) { return email.trim(); }).filter(String);
  if (guests.length === 0) throw new Error("Add at least one attendee email before sending invitations.");
  var start = buildTrackerDateTime_(training.date, training.startTime);
  var duration = Math.max(Number(training.durationHrs || 1), 0.5);
  var end = new Date(start.getTime() + duration * 60 * 60 * 1000);
  var event = CalendarApp.getDefaultCalendar().createEvent(
    "Training | " + (training.trainingTitle || training.topicmodule || "CX learning session"), start, end,
    {
      description: "Topic: " + (training.topicmodule || "") + "\nPre-training average: " + (training.preassessmentAvg === "" ? "Not recorded" : training.preassessmentAvg + "%") + "\nTrainer: " + (training.trainerName || ""),
      guests: guests.join(","),
      sendInvites: true
    }
  );
  return { success: true, eventId: event.getId(), invitees: guests.length };
}

function sendCoachingFollowupEmail(session, agentEmail) {
  if (!agentEmail) throw new Error("The selected agent has no email address.");
  if (!session.followupNotes) throw new Error("Follow-up notes are required.");
  var subject = "Coaching Follow-up | " + (session.focusArea || "Quality improvement") + " | " + session.date;
  var htmlBody = "<div style='font-family:Georgia,serif;max-width:640px;margin:auto;color:#1F1B1B'>" +
    "<div style='background:#24631E;color:white;padding:16px 20px'><h2 style='margin:0;font-size:20px'>Coaching Follow-up Record</h2></div>" +
    "<div style='border:1px solid #E5E5C9;border-top:0;padding:20px'>" +
    "<p>Hi <strong>" + session.agentName + "</strong>,</p>" +
    "<p>This documents the follow-up to our coaching on <strong>" + session.focusArea + "</strong>.</p>" +
    "<p><strong>Pre-score:</strong> " + session.precoachingScore + "% &nbsp; <strong>Post-score:</strong> " + (session.postcoachingScore === "" ? "Pending" : session.postcoachingScore + "%") + "</p>" +
    "<div style='background:#FAFAF5;border-left:4px solid #FFB200;padding:14px;white-space:pre-wrap'>" + session.followupNotes + "</div>" +
    "<p><strong>Next follow-up:</strong> " + (session.followupDate || "Not scheduled") + "</p>" +
    "<p style='font-size:12px;color:#56564F'>This email forms part of the Q&T coaching documentation trail.</p></div></div>";
  MailApp.sendEmail({ to: agentEmail, subject: subject, htmlBody: htmlBody });
  return true;
}

function sendTrainingFollowupEmail(training) {
  var recipients = String(training.attendeeEmails || "").split(",").map(function(email) { return email.trim(); }).filter(String);
  if (recipients.length === 0) throw new Error("Add attendee emails before sending follow-up documentation.");
  var delta = training.preassessmentAvg !== "" && training.postassessmentAvg !== ""
    ? Number(training.postassessmentAvg) - Number(training.preassessmentAvg) : null;
  var subject = "Training Follow-up | " + (training.trainingTitle || training.topicmodule || "CX learning session");
  var htmlBody = "<div style='font-family:Georgia,serif;max-width:640px;margin:auto;color:#1F1B1B'>" +
    "<div style='background:#1F1B1B;color:white;padding:16px 20px'><h2 style='margin:0;font-size:20px'>Training Follow-up Record</h2></div>" +
    "<div style='border:1px solid #E5E5C9;border-top:0;padding:20px'><p><strong>Topic:</strong> " + training.topicmodule + "</p>" +
    "<p><strong>Pre-assessment:</strong> " + training.preassessmentAvg + "% &nbsp; <strong>Post-assessment:</strong> " + (training.postassessmentAvg === "" ? "Pending" : training.postassessmentAvg + "%") +
    (delta === null ? "" : " &nbsp; <strong>Movement:</strong> " + (delta >= 0 ? "+" : "") + delta.toFixed(1) + " points") + "</p>" +
    "<div style='background:#FAFAF5;border-left:4px solid #FF8C00;padding:14px;white-space:pre-wrap'>" + (training.notes || "No additional notes recorded.") + "</div>" +
    "<p style='font-size:12px;color:#56564F'>This email forms part of the Q&T training documentation trail.</p></div></div>";
  MailApp.sendEmail({ to: recipients.join(","), subject: subject, htmlBody: htmlBody });
  return true;
}
