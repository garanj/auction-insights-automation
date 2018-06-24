/**
 * Auction insights downloader script.
 *
 * This script requires specific set up in order to operate correctly. See
 * guidance at : https://github.com/plemont/auction-insights-automation
 */
// Email address for notificatons of processing errors.
var EMAIL_RECIPIENT = 'INSERT_RECIPIENT_HERE';

// In this example script, each report downloaded is written to Drive. This
// holds the ID of the folder to write files to.
var OUTPUT_FOLDER_ID = 'INSERT_DRIVE_FOLDER_ID';

// Spreadsheet should have the following headings in order:
// - Unique Report Name - The report name as set up in AdWords.
// - Description - Free text description
// - Days Before Alert - The number of days since the last report after which
//       to raise an alert for a missing alert.
// - Last Received - The date of the last received report.
var REPORT_MAP_SPREADSHEET_URL = 'INSERT_SPREADSHEET_URL_HERE';

var SUBJECT_REGEX = /^AdWords Report Request \| (.*)$/;
var URL_REGEX = /(https:\/\/adwords\.google\.com\/aw_reporting\/email_download\S+)/;
var OUTPUT_FOLDER = null;
var REPORT_MAP = {};
var MILLIS_PER_DAY = 86400000;

function main() {
  loadReportMap();
  OUTPUT_FOLDER = DriveApp.getFolderById(OUTPUT_FOLDER_ID);
  processThreads();
  
  checkLastReportDates();
}

function processThreads() {
  var threads = GmailApp.getInboxThreads();
  for (var i = 0; i < threads.length; i++) {
    var thread = threads[i];
    processThread(thread);
  }
}

function processThread(thread) {
  var messages = thread.getMessages();
  for (var i = 0; i < messages.length; i++) {
    var message = messages[i];
    // If the message is starred, it has already been processed and some
    // error occurred. Therefore, don't process again.    
    if (!message.isStarred()) {
      processMessage(message);
    }
  }
}

function processMessage(message) {
  var error = null;
  var subject = message.getSubject();
  var reportName = getReportNameFromSubject(message);
  if (reportName && REPORT_MAP[reportName]) {
    var url = getDownloadUrlFromBody(message);
    if (url) {
      var content = downloadContentFromMessageUrl(url);
      if (content) {
        var csv = convertDownloadToCsv(content);
        if (csv) {
          // Do something with the CSV. e.g. save to Drive
          var date = Utilities.formatDate(message.getDate(),
            Session.getScriptTimeZone(), 'yyyy-MM-dd');
          var filename = date + '-' + reportName;
          OUTPUT_FOLDER.createFile(filename, csv, MimeType.CSV);
          
          markLastReportDate(reportName, message.getDate());
          // Delete the message.
          message.moveToTrash();
        } else {
          error = 'Error in converting download to csv for message: ' + subject;
        }
      } else {
        error = 'Error in downloading content for message: ' + subject;
      }
    } else {
      error = 'No download URL found for message: ' + subject;
    }
  } else {
    error = 'Unexpected subject found for message: ' + subject;
  }
  
  if (error) {
    message.star();
    sendErrorEmail(error);
  }
}

function getReportNameFromSubject(message) {
  var subject = message.getSubject();
  var subjectMatches = SUBJECT_REGEX.exec(subject);
  if (subjectMatches && subjectMatches.length) {
    return subjectMatches[1]; 
  } 
}

function getDownloadUrlFromBody(message) {
  var body = message.getBody();
  var bodyMatches = URL_REGEX.exec(body);
  if (bodyMatches && bodyMatches.length) {
    return bodyMatches[1]; 
  } 
}

function downloadContentFromMessageUrl(url) {
  var options = {muteHttpExceptions: true};
  var response = UrlFetchApp.fetch(url, options);
  if (response.getResponseCode() === 200) {
    return response.getContent();      
  }
}

function convertDownloadToCsv(content) {
  try {
    var blob = Utilities.newBlob(content, 'application/x-gzip');
    var csv = Utilities.ungzip(blob).getDataAsString();
  } catch (exception) {
    return;     
  }
  return csv;
}

function sendErrorEmail(error) {
  MailApp.sendEmail(EMAIL_RECIPIENT, 'Auction Insights script alert', error);
}

function loadReportMap() {
  var spreadsheet = SpreadsheetApp.openByUrl(REPORT_MAP_SPREADSHEET_URL);
  var sheet = spreadsheet.getActiveSheet();
  var data = sheet.getDataRange().getValues();
  
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    
    REPORT_MAP[row[0]] = {
      description: row[1],
      daysBeforeAlert: row[2],
      lastReport: row[3]
    }
  }
}

function markLastReportDate(reportName, date) {
  var spreadsheet = SpreadsheetApp.openByUrl(REPORT_MAP_SPREADSHEET_URL);
  var sheet = spreadsheet.getActiveSheet();
  var data = sheet.getDataRange().getValues();
  
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    if (row[0] === reportName && (!row[3] || row[3] < date)) {
      sheet.getRange(i + 1, 4, 1, 1).setValue(date);
      return;
    }
  }
}

function checkLastReportDates() {
  var spreadsheet = SpreadsheetApp.openByUrl(REPORT_MAP_SPREADSHEET_URL);
  var sheet = spreadsheet.getActiveSheet();
  var data = sheet.getDataRange().getValues();
  
  var now = new Date().getTime();
  var overdueReports = [];
  
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var reportName = row[0];
    var daysBeforeAlert = row[2];
    var lastReceived = row[3] ? row[3].getTime() : 0;
    if (now - lastReceived > daysBeforeAlert * MILLIS_PER_DAY) {
      overdueReports.push(reportName);
    }
  }
  
  if (overdueReports.length) {
    MailApp.sendEmail(EMAIL_RECIPIENT, 'Auction Insights script alert',
                      'Email reports were expected for the following reports,' +
                      ' but have not been seen: ' + overdueReports.join(', '));
  }
}
