// Google Apps Script — Novel Scraper Drive Uploader
// Deploy as: Execute as "Me", Access "Anyone"
// Then paste the web app URL into scraper server as APPS_SCRIPT_URL

var FOLDER_ID = "1UyCUOcPTQLGSkII4DoEPd-_gKGZwLO9E";

function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);
    var filename = data.filename;
    var folderId = data.folder_id || FOLDER_ID;
    var subfolderName = data.subfolder_name || "";

    var bytes = Utilities.base64Decode(data.content);
    var blob = Utilities.newBlob(
      bytes,
      "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
      filename
    );

    var folder = DriveApp.getFolderById(folderId);

    // Find or create a subfolder named after the book
    if (subfolderName) {
      var subFolders = folder.getFoldersByName(subfolderName);
      if (subFolders.hasNext()) {
        folder = subFolders.next();
      } else {
        folder = folder.createFolder(subfolderName);
      }
    }

    // Replace existing file if same name
    var existing = folder.getFilesByName(filename);
    while (existing.hasNext()) {
      existing.next().setTrashed(true);
    }

    folder.createFile(blob);

    return ContentService
      .createTextOutput(JSON.stringify({ status: "ok", filename: filename }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: "error", message: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// Test via GET to confirm the script is deployed
function doGet(e) {
  return ContentService
    .createTextOutput(JSON.stringify({ status: "ok", message: "Novel Scraper uploader ready" }))
    .setMimeType(ContentService.MimeType.JSON);
}
