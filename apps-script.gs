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
          var foundRowIndex = -1;
          var targetTime = new Date(d.timestamp).getTime();

          for (var r = values.length - 1; r >= 1; r--) {
            var cellVal = values[r][0];
            var cellTime = 0;
            if (cellVal instanceof Date) {
              cellTime = cellVal.getTime();
            } else if (cellVal) {
              cellTime = new Date(cellVal).getTime();
            }

            // Compare times (5-second tolerance) and check if agent name is present in the row
            if (Math.abs(cellTime - targetTime) < 5000) {
              var rowStr = JSON.stringify(values[r]);
              if (!d.agentName || rowStr.indexOf(d.agentName) !== -1) {
                foundRowIndex = r + 1; // 1-based index
                break;
              }
            }
          }

          if (foundRowIndex !== -1) {
            var headers = values[0];
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
            diag += ' | SheetUpdate=ROW_NOT_FOUND (ts=' + d.timestamp + ', targetTime=' + targetTime + ')';
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
