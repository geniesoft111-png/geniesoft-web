/**
 * GenieSoft Technologies - Google Apps Script Form Handler
 * 
 * SETUP INSTRUCTIONS:
 * 1. Open Google Sheets (https://sheets.google.com) and create a new Spreadsheet (e.g., "GenieSoft Website Leads").
 * 2. In the top menu, click: Extensions > Apps Script.
 * 3. Replace all code in the script editor with this file's code.
 * 4. Click "Save" (Ctrl+S / Cmd+S).
 * 5. Click "Deploy" (top right) > "Manage deployments" (or "New deployment").
 * 6. Make sure:
 *    - Execute as: "Me"
 *    - Who has access: "Anyone"  <-- CRITICAL! Must be "Anyone", NOT "Only myself"
 * 7. Click "Deploy" and authorize access when requested.
 */

function doPost(e) {
  var lock = LockService.getScriptLock();
  // Wait up to 30 seconds for concurrent requests
  lock.tryLock(30000);

  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getActiveSheet();

    // Parse incoming data
    var data = {};

    // 1. Try parsing JSON body
    if (e.postData && e.postData.contents) {
      try {
        data = JSON.parse(e.postData.contents);
      } catch (err) {
        // Fallback if not JSON
      }
    }

    // 2. Fallback to URL/Form parameters
    if (!data || Object.keys(data).length === 0) {
      data = e.parameter || {};
    }

    // Ensure header row exists if sheet is empty
    if (sheet.getLastRow() === 0) {
      sheet.appendRow([
        "Timestamp",
        "Name",
        "Email",
        "Phone",
        "Subject",
        "Message"
      ]);
      // Style header row
      var headerRange = sheet.getRange(1, 1, 1, 6);
      headerRange.setFontWeight("bold");
      headerRange.setBackground("#0A1931");
      headerRange.setFontColor("#FFFFFF");
    }

    // Prepare row data
    var timestamp = new Date();
    var name = data.name || "";
    var email = data.email || "";
    var phone = data.phone || "";
    var subject = data.subject || "";
    var message = data.message || "";

    // Append inquiry to sheet
    sheet.appendRow([timestamp, name, email, phone, subject, message]);

    // Optional: Send email notification to admin
    try {
      var adminEmail = "info@geniesoft.tech";
      var emailSubject = "New Website Lead: " + (subject || "General Inquiry");
      var emailBody = "New Lead received from GenieSoft Website:\n\n" +
        "Name: " + name + "\n" +
        "Email: " + email + "\n" +
        "Phone: " + phone + "\n" +
        "Subject: " + subject + "\n\n" +
        "Message:\n" + message + "\n\n" +
        "Date: " + timestamp.toString();

      MailApp.sendEmail(adminEmail, emailSubject, emailBody);
    } catch (mailErr) {
      Logger.log("Email notification error: " + mailErr);
    }

    return ContentService
      .createTextOutput(JSON.stringify({ result: "success" }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    return ContentService
      .createTextOutput(JSON.stringify({ result: "error", error: error.toString() }))
      .setMimeType(ContentService.MimeType.JSON);

  } finally {
    lock.releaseLock();
  }
}

function doGet(e) {
  return ContentService
    .createTextOutput(JSON.stringify({ status: "online", message: "GenieSoft Form Handler is running." }))
    .setMimeType(ContentService.MimeType.JSON);
}
