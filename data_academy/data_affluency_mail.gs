/**
 * Sends emails with data from the current spreadsheet.
 */
function sendEmails() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var startRow = 1; // First row of data to process
  var numRows = 5; // Number of rows to process
  // Fetch the range of cells
  var dataRange = sheet.getRange(startRow, 1, numRows, 2);
  // Fetch values for each row in the Range.
  var data = dataRange.getValues();
  for (i in data) {
    var row = data[i];
    var emailAddress = row[0]; // First column
    //var message = row[1]; // Second column
    var subject = 'Data Barometer: How are we doing?';
    MailApp.sendEmail(emailAddress, subject, "Hello Fellow Babbelonian!\n\nAt Analytics, we want to make sure that we're making the most out of our data.\n\nLet us know how you're feeling in this 100% *anonymous* 4-question survey\nhttps://www.surveymonkey.de/r/85XRHCM.\n\nPS:\nSince this email is sent to 10 random data-users, you might get this email again in the future. Please answer the survey every time you receive it.\nReply to this message if you have questions or comments.\n\nThanks for your help :)");
  }
}
