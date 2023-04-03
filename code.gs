function formSubmitReply(e) {
  var userEmail = e.values[2]; // Column D for email address
  var sheet = SpreadsheetApp.getActiveSheet();
  var lastRow = sheet.getLastRow();
  var ticketNumber = sheet.getRange(lastRow, getColIndexByName("Ticket Number")).getValue(); // Added ticket number
  var fullName = sheet.getRange(lastRow, getColIndexByName("Full Name")).getValue(); // Added full name
  var shortDescription = sheet.getRange(lastRow, getColIndexByName("Short Description")).getValue(); // Added short description
  var detailedDescription = sheet.getRange(lastRow, getColIndexByName("Detailed Description")).getValue(); // Added detailed description
  var severity = sheet.getRange(lastRow, getColIndexByName("Severity")).getValue(); // Added severity

  // Set the status of the new ticket to 'New'.
  // Column J is the Status column
  sheet.getRange(lastRow, getColIndexByName("Status")).setValue("New");

  // Calculate how many other 'New' tickets are ahead of this one
  var numNew = 0;
  for (var i = 2; i < lastRow; i++) {
    if (sheet.getRange(i, getColIndexByName("Status")).getValue() == "New") {
      numNew++;
    }
  }
  MailApp.sendEmail(userEmail,
                    ticketNumber + " has been created", // Modified subject
                    "Dear " + fullName + ",\n\nThank you for contacting G6line Tech support. We have received your support request with the following details:\n\n" +
                    "Ticket Number: " + ticketNumber + "\n" +
                    "Short Description: " + shortDescription + "\n" +
                    "Detailed Description: " + detailedDescription + "\n" +
                    "Severity: " + severity + "\n\n" +
                    "Our support team will begin working on your request as soon as possible. Requests are prioritized based on urgency and type of request. You are currently number " +
                    (numNew + 1) + " in the queue.\n\nSincerely," +
                    "\n\nThe G6line Tech Support Team"); // Modified body
}

function emailStatusUpdates() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var row = sheet.getActiveRange().getRowIndex();
  var ticketNumber = sheet.getRange(row, getColIndexByName("Ticket Number")).getValue(); // Added ticket number
  var fullName = sheet.getRange(row, getColIndexByName("Full Name")).getValue(); // Added full name
  var shortDescription = sheet.getRange(row, getColIndexByName("Short Description")).getValue(); // Added short description
  var detailedDescription = sheet.getRange(row, getColIndexByName("Detailed Description")).getValue(); // Added detailed description
  var severity = sheet.getRange(row, getColIndexByName("Severity")).getValue(); // Added severity
  var customerNotes = sheet.getRange(row, getColIndexByName("Help Desk Notes")).getValue(); // Added customer notes
  var status = sheet.getRange(row, getColIndexByName("Status")).getValue(); // Added status
  var resolution = sheet.getRange(row, getColIndexByName("Resolution")).getValue(); // Added resolution
  var userEmail = sheet.getRange(row, getColIndexByName("Email Address")).getValue(); // Email Address
  var subject = "Status Update: " + ticketNumber; // Modified subject
  var body = "Dear " + fullName + ",\n\nThe G6line Tech support team has updated the status of your support request with the following details:\n\n" +
"Short Description: " + shortDescription + "\n" +
"Detailed Description: " + detailedDescription + "\n" +
"Severity: " + severity + "\n" +
"Customer Notes: " + customerNotes + "\n" +
"Status: " + status + "\n" +
"Resolution: " + resolution + "\n\n" +
"If you have any further questions or concerns, please do not hesitate to contact us at g6linetech@gmail.com. We are always happy to assist you.\n\nSincerely," +
"\n\nThe G6line Tech Support Team"; // Modified body

MailApp.sendEmail(userEmail, subject, body, {name:"G6line Tech Support Team"}); // Modified sender name
}

function onOpen() {
var subMenus = [{name:"Send Status Email (choose a row)", functionName: "emailStatusUpdates"}];
SpreadsheetApp.getActiveSpreadsheet().addMenu("Tech Support Menu", subMenus); // Modified menu name
}

function getColIndexByName(colName) {
var sheet = SpreadsheetApp.getActiveSheet();
var numColumns = sheet.getLastColumn();
var row = sheet.getRange(1, 1, 1, numColumns).getValues();
for (i in row[0]) {
var name = row[0][i];
if (name == colName) {
return parseInt(i) + 1;
}
}
return -1;
}
