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

    if (!d.emailOnly) {
      // --- Write to Google Sheet ---
      var sheetName = d.auditType;
      var sheet = doc.getSheetByName(sheetName);
      if (!sheet) {
        sheet = doc.insertSheet(sheetName);
        sheet.appendRow(d.headers);
      }

      // --- Resolve Agent Folder in Drive ---
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

      // --- Upload files to Drive ---
      if (d.files && d.files.length > 0 && agentFolder) {
        var today = new Date();
        var dateStr = today.getFullYear() + "-"
          + String(today.getMonth() + 1).padStart(2, "0") + "-"
          + String(today.getDate()).padStart(2, "0");
        var safeName = d.agentName.replace(/[^a-zA-Z0-9]/g, "_");
        var safeType = d.auditType.replace(/[^a-zA-Z0-9]/g, "_");

        for (var j = 0; j < d.files.length; j++) {
          try {
            var f = d.files[j];
            var ext = f.fileName.split(".").pop();
            var markedName = safeName + "_" + safeType + "_" + dateStr + "_" + (j+1) + "." + ext;

            var parts = f.fileData.split(",");
            var b64 = parts.length > 1 ? parts[1] : parts[0];
            var bytes = Utilities.base64Decode(b64);
            var file = agentFolder.createFile(markedName, bytes, f.fileMimeType || "audio/mpeg");

            if (d.agentEmail && file) {
              try { file.addViewer(d.agentEmail); } catch(e) {}
            }

            fileUrls.push(file.getUrl());
            attachments.push(file); // Store the file for email attachment
            diag += ' | File' + (j+1) + '=OK';
          } catch (fileErr) {
            diag += ' | File' + (j+1) + '=FAIL:' + fileErr.toString();
          }
        }
      } else {
        diag += ' | Upload skipped: files=' + (d.files ? d.files.length : 0) + ' agentFolder=' + (agentFolder ? 'OK' : 'null');
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
