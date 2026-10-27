/**
 * Backend for the audio survey (index.html).
 *
 * Deploy this as a Web App:
 *   1. Go to https://script.google.com -> New project, paste this file in as Code.gs.
 *   2. Deploy > New deployment > type: Web app.
 *        - Execute as: Me
 *        - Who has access: Anyone
 *   3. Authorize the requested permissions (Drive + Sheets) when prompted.
 *   4. Copy the Web app URL into CONFIG.ENDPOINT_URL in index.html.
 *   5. Run setup() once from the Apps Script editor (select it in the toolbar, click Run)
 *      to pre-create the Sheet + Drive folder and grant permissions up front.
 *
 * On first request (if setup() wasn't run manually) the script will create:
 *   - A Google Sheet named "Audio Survey Responses"
 *   - A Drive folder named "Audio Survey Recordings"
 * and remember their IDs in Script Properties, so they're reused on every submission.
 */

var ANSWER_QUESTION_IDS = ["q1", "q2", "q3"];
var AUDIO_QUESTION_IDS = ["q4", "q5", "q6"];

function getOrCreateSpreadsheet() {
  var props = PropertiesService.getScriptProperties();
  var id = props.getProperty("SPREADSHEET_ID");
  var ss;
  if (id) {
    try {
      ss = SpreadsheetApp.openById(id);
    } catch (e) {
      ss = null;
    }
  }
  if (!ss) {
    ss = SpreadsheetApp.create("Audio Survey Responses");
    props.setProperty("SPREADSHEET_ID", ss.getId());
  }
  var sheet = ss.getSheets()[0];
  if (sheet.getLastRow() === 0) {
    var header = ["Server Timestamp", "Submitted At"]
      .concat(ANSWER_QUESTION_IDS)
      .concat(AUDIO_QUESTION_IDS.map(function (id) { return id + "_audio_link"; }));
    sheet.appendRow(header);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function getOrCreateFolder() {
  var props = PropertiesService.getScriptProperties();
  var id = props.getProperty("DRIVE_FOLDER_ID");
  var folder;
  if (id) {
    try {
      folder = DriveApp.getFolderById(id);
    } catch (e) {
      folder = null;
    }
  }
  if (!folder) {
    folder = DriveApp.createFolder("Audio Survey Recordings");
    props.setProperty("DRIVE_FOLDER_ID", folder.getId());
  }
  return folder;
}

/** Run this once manually from the Apps Script editor to pre-authorize and create resources. */
function setup() {
  var sheet = getOrCreateSpreadsheet();
  var folder = getOrCreateFolder();
  Logger.log("Spreadsheet: " + sheet.getParent().getUrl());
  Logger.log("Drive folder: " + folder.getUrl());
}

function doGet(e) {
  return ContentService.createTextOutput("Audio survey endpoint is live.")
    .setMimeType(ContentService.MimeType.TEXT);
}

function doPost(e) {
  var result;
  try {
    if (!e || !e.postData || !e.postData.contents) {
      throw new Error("No data received.");
    }
    var data = JSON.parse(e.postData.contents);
    var answers = data.answers || {};
    var audioItems = data.audio || [];
    var submittedAt = data.submittedAt || new Date().toISOString();

    var folder = getOrCreateFolder();
    var sheet = getOrCreateSpreadsheet();

    // Give each submission its own subfolder so recordings from the same
    // respondent stay grouped together.
    var safeStamp = submittedAt.replace(/[:.]/g, "-");
    var submissionFolder = folder.createFolder(safeStamp);

    var audioLinks = {};
    audioItems.forEach(function (item) {
      if (!item.base64) return;
      var ext = item.extension || "webm";
      var filename = (item.questionId || "audio") + "_" + safeStamp + "." + ext;
      var bytes = Utilities.base64Decode(item.base64);
      var blob = Utilities.newBlob(bytes, item.mimeType || "audio/webm", filename);
      var file = submissionFolder.createFile(blob);
      audioLinks[item.questionId] = file.getUrl();
    });

    var row = [new Date(), submittedAt];
    ANSWER_QUESTION_IDS.forEach(function (qid) {
      row.push(answers[qid] !== undefined ? answers[qid] : "");
    });
    AUDIO_QUESTION_IDS.forEach(function (qid) {
      row.push(audioLinks[qid] || "");
    });
    sheet.appendRow(row);

    result = { status: "ok" };
  } catch (err) {
    result = { status: "error", message: err && err.message ? err.message : String(err) };
  }
  return ContentService.createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}
