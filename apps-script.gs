// QA Audit Script — Per-Agent Folder + Auto Email + Diagnostics
// Copy this entire file into your Apps Script editor.
// Then: Deploy → Manage Deployments → edit pencil → New version → Deploy → Authorize

function doPost(e) {
  try {
    var d = JSON.parse(e.postData.contents);
    var doc = SpreadsheetApp.getActiveSpreadsheet();
    var fileUrls = [];
    var fileUrlStr = "No files attached";
    var attachments = [];
    var agentFolderUrl = "";

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
          var values = sheet.getDataRange().getValues();
          var displayValues = sheet.getDataRange().getDisplayValues();
          var headers = values[0];

          // 1. Ensure new columns/headers are present in the sheet
          if (d.headers && d.headers.length > 0) {
            ensureHeaders(sheet, d.headers);
            // Re-read values and headers after ensuring they are correct
            values = sheet.getDataRange().getValues();
            displayValues = sheet.getDataRange().getDisplayValues();
            headers = values[0];
          }

          var foundRowIndex = -1;
          var sheetTimezone = doc.getSpreadsheetTimeZone();

          for (var r = values.length - 1; r >= 1; r--) {
            var cellVal = values[r][0];
            var displayVal = displayValues[r][0];
            var isMatched = false;

            // Check exact matches first
            if (cellVal == d.timestamp || displayVal == d.timestamp) {
              isMatched = true;
            } else {
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
                foundRowIndex = r + 1; // 1-based index
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
      if (d.headers && d.headers.length > 0) {
        ensureHeaders(sheet, d.headers);
      }

      // Upload files to Google Drive (if any and folder is resolved)
      if (attachments.length > 0 && agentFolder) {
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
      for (var i = 0; i < d.row.length; i++) {
        if (d.row[i] === "[UPLOADING...]") d.row[i] = fileUrlStr;
      }
      sheet.appendRow(d.row);
    }

    // --- Send Email ---
    if (d.emailSettings && d.emailSettings.to && d.emailSettings.send !== false) {
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
      .createTextOutput(JSON.stringify({ status: "success", debug: diag }))
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

// Database Sheet names
var SHEET_COACHING = "CoachingLog";
var SHEET_TRAINING = "TrainingLog";
var SHEET_PLAN = "ExecutionPlan";
var SHEET_SETTINGS = "TrackerSettings";

// Auto-initialize sheets
function initTrackerSheets() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // 1. Coaching Log Sheet
  var coachingSheet = ss.getSheetByName(SHEET_COACHING);
  if (!coachingSheet) {
    coachingSheet = ss.insertSheet(SHEET_COACHING);
    coachingSheet.appendRow([
      "ID", "Date", "Agent Name", "Coach Name", "Type", 
      "Focus Area", "Pre-Coaching Score (%)", "Post-Coaching Score (%)", 
      "Observations", "Action Items", "Follow-up Date", "Status", 
      "Created At", "Updated At"
    ]);
  }
  
  // 2. Training Log Sheet
  var trainingSheet = ss.getSheetByName(SHEET_TRAINING);
  if (!trainingSheet) {
    trainingSheet = ss.insertSheet(SHEET_TRAINING);
    trainingSheet.appendRow([
      "ID", "Date", "Training Title", "Trainer Name", "Type", 
      "Topic/Module", "Duration (Hrs)", "Attendees Count", "Pre-Assessment Avg (%)", 
      "Post-Assessment Avg (%)", "Delivery Method", "Status", "Notes", 
      "Created At", "Updated At"
    ]);
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
      coaching: getSheetData(ss.getSheetByName(SHEET_COACHING)),
      training: getSheetData(ss.getSheetByName(SHEET_TRAINING)),
      plans: getSheetData(ss.getSheetByName(SHEET_PLAN)),
      agents: getSheetData(ss.getSheetByName("AgentDirectory")),
      settings: getSettingsData(ss.getSheetByName(SHEET_SETTINGS)),
      scorecards: {
        inbound: getSheetData(ss.getSheetByName("Inbound Scorecard")),
        sms: getSheetData(ss.getSheetByName("SMS Quality Audit")),
        outbound: getSheetData(ss.getSheetByName("Outbound Call Scorecard (Campaign-Focused)"))
      },
      needsFeedback: {
        needs: getSheetData(ss.getSheetByName("PreTrainingNeeds")),
        feedback: getSheetData(ss.getSheetByName("PostTrainingFeedback"))
      },
      deescalation: getDeescalationData()
    };
  } catch(err) {
    Logger.log("getTrackerData error: " + err.toString());
    return { 
      coaching: [], training: [], plans: [], agents: [], settings: {}, 
      scorecards: { inbound: [], sms: [], outbound: [] },
      needsFeedback: { needs: [], feedback: [] },
      deescalation: [],
      error: err.toString() 
    };
  }
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
    return getSheetData(sheet);
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

function getSheetData(sheet) {
  if (!sheet) return [];
  var data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];
  var headers = data[0];
  var rows = [];
  for (var i = 1; i < data.length; i++) {
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
  return saveRow(SHEET_COACHING, session, [
    "ID", "Date", "Agent Name", "Coach Name", "Type", 
    "Focus Area", "Pre-Coaching Score (%)", "Post-Coaching Score (%)", 
    "Observations", "Action Items", "Follow-up Date", "Status", 
    "Created At", "Updated At"
  ]);
}

function saveTrainingSession(training) {
  return saveRow(SHEET_TRAINING, training, [
    "ID", "Date", "Training Title", "Trainer Name", "Type", 
    "Topic/Module", "Duration (Hrs)", "Attendees Count", "Pre-Assessment Avg (%)", 
    "Post-Assessment Avg (%)", "Delivery Method", "Status", "Notes", 
    "Created At", "Updated At"
  ]);
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
  if (!expectedHeaders || expectedHeaders.length === 0) return;
  var lastCol = sheet.getLastColumn();
  var maxCols = sheet.getMaxColumns();
  if (maxCols < expectedHeaders.length) {
    sheet.insertColumnsAfter(maxCols, expectedHeaders.length - maxCols);
  }
  var currentHeaders = lastCol > 0 ? sheet.getRange(1, 1, 1, lastCol).getValues()[0] : [];
  var isMismatch = currentHeaders.length !== expectedHeaders.length;
  if (!isMismatch) {
    for (var i = 0; i < expectedHeaders.length; i++) {
      if (currentHeaders[i] !== expectedHeaders[i]) {
        isMismatch = true;
        break;
      }
    }
  }
  if (isMismatch) {
    sheet.getRange(1, 1, 1, expectedHeaders.length).setValues([expectedHeaders]);
  }
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
    "<div style='font-family: Arial, sans-serif; max-width: 600px; margin: 0 auto; padding: 20px; border: 1px solid #e4e2d8; border-radius: 12px; background-color: #f4faf3;'>" +
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
