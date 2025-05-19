function saveAttachmentsToStructuredFolders() {
  var mainFolderId = "1ZRw-1xhUv3K2-Q6T1Z2aCoFZGetqwfun"; // 🔁 Replace with your actual folder ID
  var subjectFolderMap = {
    "SBI Card ELITE Monthly Statement": "sbi",
    "HDFC Bank - Diners Club International Credit Card Statement": "hdfc",
    "ICICI Bank Credit Card Statement": "icici"
  };

  var mainFolder = DriveApp.getFolderById(mainFolderId);

  Object.keys(subjectFolderMap).forEach(subject => {
    var cardType = subjectFolderMap[subject];
    var threads = GmailApp.search('subject:"' + subject + '"');

    threads.forEach(thread => {
      var messages = thread.getMessages();
      messages.forEach(message => {
        var attachments = message.getAttachments();

        attachments.forEach(attachment => {
          var originalFileName = attachment.getName();
          if (originalFileName && originalFileName.toLowerCase().endsWith(".pdf")) {
            var { year, month } = extractYearMonth(originalFileName, cardType, message.getSubject());

            if (year && month) {
              var cardFolder = getOrCreateFolder(mainFolder, cardType);
              var renamedFileName = month + "_" + year + ".pdf"; // 💡 MM_YYYY.pdf format

              if (!fileExists(cardFolder, renamedFileName)) {
                // cardFolder.createFile(attachment.getAs(MimeType.PDF)).setName(renamedFileName);
                //if (attachment.getContentType() === MimeType.PDF) {
                  cardFolder.createFile(attachment.copyBlob()).setName(renamedFileName);
                //}
                Logger.log("✅ Saved: " + renamedFileName + " to folder: " + cardType);
              } else {
                Logger.log("⏩ Skipped (already exists): " + renamedFileName);
              }
            } else {
              Logger.log("⚠️ Could not determine date for: " + originalFileName);
            }
          } else {
            Logger.log("❌ Ignored non-PDF: " + originalFileName);
          }
        });
      });
    });
  });
}

// Extract year and month from filename based on card type
function extractYearMonth(fileName, cardType, subjectLine) {
  try {
    var year = "", month = "";

    if (cardType === "icici") {
      var match = subjectLine.match(/to\s+([A-Za-z]+)\s+\d{1,2}\s+(\d{4})$/i);
      if (match) {
        var monthName = match[1];
        year = match[2];
        month = getMonthNumberFromName(monthName);
      }
    } else if (cardType === "sbi") {
      var parts = fileName.split("_");
      if (parts.length > 1) {
        var dateStr = parts[1].replace(".pdf", "");
        if (dateStr.length === 8) { // Format: DDMMYYYY
          month = dateStr.slice(2, 4);
          year = dateStr.slice(4);
        }
      }
    } else if (cardType === "hdfc") {
      var temp = fileName.split("_")[1];
      temp = temp.split(".")[0];
      var tokens = temp.split("-");
      year = tokens[2];
      month = tokens[1];
    }

    // Ensure month is always 2-digit
    month = ("0" + parseInt(month)).slice(-2);

    return { year, month };
  } catch (e) {
    Logger.log("❌ Error extracting date from: " + fileName + " -> " + e.message);
    return { year: null, month: null };
  }
}

// Check if a file already exists in the folder
function fileExists(folder, fileName) {
  var files = folder.getFilesByName(fileName);
  return files.hasNext();
}

// Create or retrieve folder
function getOrCreateFolder(parent, name) {
  var folders = parent.getFoldersByName(name);
  if (folders.hasNext()) {
    return folders.next();
  }
  return parent.createFolder(name);
}

function getMonthNumberFromName(monthName) {
  var months = {
    january: 1, february: 2, march: 3, april: 4,
    may: 5, june: 6, july: 7, august: 8,
    september: 9, october: 10, november: 11, december: 12
  };
  return months[monthName.toLowerCase()] || null;
}