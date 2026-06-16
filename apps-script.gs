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
            var headersModified = false;
            for (var h = 0; h < d.headers.length; h++) {
              var hName = d.headers[h];
              if (headers.indexOf(hName) === -1) {
                sheet.getRange(1, h + 1).setValue(hName);
                headersModified = true;
              }
            }
            if (headersModified) {
              // Re-read values and headers after dynamically adding them
              values = sheet.getDataRange().getValues();
              displayValues = sheet.getDataRange().getDisplayValues();
              headers = values[0];
            }
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
        sheet.appendRow(d.headers);
      }

      // Upload files to Google Drive (if any and folder is resolved)
      if (attachments.length > 0 && agentFolder) {
        for (var j = 0; j < attachments.length; j++) {
          try {
            var blob = attachments[j];
            var file = agentFolder.createFile(blob);
            if (d.agentEmail && file) {
              try { file.addViewer(d.agentEmail); } catch(e) {}
            }
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

function doGet() {
  return ContentService
    .createTextOutput(JSON.stringify({ status: "ok" }))
    .setMimeType(ContentService.MimeType.JSON);
}
