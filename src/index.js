// Requirements:
//  1. Receive .xsl via email
//  2. Convert .xsl to google sheet/csv
//  3. Append data to master sheet
//  4. Append data to live input sheet

const excludeHeaders = true;
const gmailLabelPathIncoming = "zAutomation/Incoming";
const gmailLabelPathProcessed = "zAutomation/Processed";
const driveTemporaryFolderId = "12sIBE9u-R5P5Vj1Bpvt6QQZKC7UqZBCA";
const driveTargetSheetIds = [
  "1Av6K4tp_BFJwmS471nx4UjuNuCUvcb6KgRrTf5ObgVE",
  "1bbZZt_oUaTyyMYmkIkpbi7bOE2Al0WjkGJrv9j5VQRM",
];

function getGmailIncomingLabel() {
  return GmailApp.getUserLabelByName(gmailLabelPathIncoming);
}

function getGmailProcessedLabel() {
  return GmailApp.getUserLabelByName(gmailLabelPathProcessed);
}

function getDriveTemporaryFolder() {
  return DriveApp.getFolderById(driveTemporaryFolderId);
}

function getDriveTargetSheets() {
  return driveTargetSheetIds.map((sheetId) => {
    try {
      return SpreadsheetApp.openById(sheetId);
    } catch (error) {
      throw new Error("Failed to access sheet " + sheetId);
    }
  });
}

function validateConfig() {
  if (typeof Drive === "undefined") {
    Logger.log(
      "You must enable Advanced Services & enable Drive https://developers.google.com/apps-script/guides/services/advanced#enabling_advanced_services"
    );
    return false;
  }
  var incomingLabel = getGmailIncomingLabel();
  if (!incomingLabel) {
    Logger.log(
      "Unable to find \"Incoming\" Gmail label by path",
      gmailLabelPathIncoming
    );
    return false;
  }
  var processedLabel = getGmailProcessedLabel();
  if (!processedLabel) {
    Logger.log(
      "Unable to find \"Processed\" Gmail label by path",
      gmailLabelPathProcessed
    );
    return false;
  }
  var temporaryFolder = getDriveTemporaryFolder();
  if (!temporaryFolder) {
    Logger.log(
      "Unable to find Temporary Drive Folder label by id",
      driveTemporaryFolderId
    );
    return false;
  }
  try {
    getDriveTargetSheets();
  } catch (error) {
    Logger.log("Unable to access Target Drive Sheets", error);
    return false;
  }

  Logger.log("Validation Successful");
  return true;
}

// eslint-disable-next-line no-unused-vars
function processIncomingAttachments() {
  if (!validateConfig()) {
    return;
  }

  var incomingLabel = getGmailIncomingLabel();
  var processedLabel = getGmailProcessedLabel();
  // The rest of the script only runs if there's at least one unread thread
  if (incomingLabel.getUnreadCount() === 0) {
    Logger.log("Found no unread messages in label", incomingLabel.getName());
    return;
  }

  var threads = incomingLabel.getThreads();
  // Going through every thread in the order received
  for (var i = threads.length - 1; i >= 0; i--) {
    var currentThread = threads[i];

    // Collecting messages from unread threads
    if (!currentThread.isUnread()) {
      continue;
    }

    var messages = GmailApp.getMessagesForThread(currentThread);

    // Collecting attachments from each message
    for (var message of messages) {
      Logger.log(
        "Processing message %s received at %s",
        message.getSubject(),
        message.getDate().toISOString()
      );
      var attachments = message.getAttachments();
      // Processing each attachment
      for (var attachment of attachments) {
        if (attachment.getName() && !attachment.getName().match(/.*\.xls/)) {
          Logger.log(
            "Cannot process attachment that is not xls. %s",
            attachment.getName()
          );
          continue;
        }
        var xslFile = attachment.copyBlob();
        processXlsFile(xslFile);
      }
    }

    currentThread.markRead().refresh();
    incomingLabel.removeFromThread(currentThread);
    processedLabel.addToThread(currentThread);
  }
}

function processXlsFile(xlsFile) {
  const temporarySheet = uploadXlsToSheets(xlsFile);

  for (const targetSheet of getDriveTargetSheets()) {
    appendDataToEnd(targetSheet, temporarySheet);
  }

  deleteFile(temporarySheet);
}

function deleteFile(file) {
  try {
    DriveApp.getFileById(file.getId()).setTrashed(true);
  } catch (error) {
    Logger.log("An error occurred deleting the file. " + error);
    throw new Error("Failed deleting the file " + file.getId());
  }
  Logger.log("File deleted");
}

function appendDataToEnd(targetSheet, sourceSheet) {
  var targetActiveSheet = targetSheet.getActiveSheet();
  var sourceActiveSheet = sourceSheet.getActiveSheet();

  var targetDataRange = targetActiveSheet.getRange(
    targetActiveSheet.getLastRow() + 1,
    1,
    sourceActiveSheet.getLastRow(),
    sourceActiveSheet.getLastColumn()
  );
  var sourceDataRange = sourceActiveSheet.getRange(
    1,
    1,
    sourceActiveSheet.getLastRow(),
    sourceActiveSheet.getLastColumn()
  );

  targetDataRange.setValues(sourceDataRange.getValues());
}

function uploadXlsToSheets(xlsFile) {
  var temporaryFolder = getDriveTemporaryFolder();

  var temporaryFileName = "tmp_" + new Date().toISOString();

  Logger.log("Creating temporary file %s", temporaryFileName);

  var metadata = {
    parents: [{ id: temporaryFolder.getId() }],
    title: temporaryFileName,
    mimeType: MimeType.MICROSOFT_EXCEL,
  };
  var options = {
    convert: true,
  };
  let file;
  try {
    file = Drive.Files.insert(metadata, xlsFile, options);
  } catch (error) {
    Logger.log("An error occurred uploading the file. " + error);
    throw new Error("Failed to upload temporary file");
  }

  let spreadsheet;
  try {
    spreadsheet = SpreadsheetApp.openById(file.getId());
  } catch (error) {
    Logger.log("An error occurred opening the spreadsheet. " + error);
    throw new Error("Failed to open temporary spreadsheet");
  }

  if (excludeHeaders) {
    spreadsheet.getActiveSheet().deleteRow(1);
  }

  return spreadsheet;
}
