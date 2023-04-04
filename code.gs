function formSubmitReply(e) {
  var userEmail = e.values[2]; // Column C for email address
  var sheet = SpreadsheetApp.getActiveSheet();
  var newRowIndex = e.range.getRow(); // Get the row index of the new row
  var fullName = e.values[1]; // Added full name
  var shortDescription = e.values[3]; // Added short description
  var detailedDescription = e.values[4]; // Added detailed description
  var severity = e.values[5]; // Added severity

  // Column I is the Status column
  sheet.getRange(newRowIndex, getColIndexByName("Status")).setValue("New"); // Set "New" status on the new row

  // Calculate how many other 'New' tickets are ahead of this one
  var numNew = 0;
  for (var i = 2; i < newRowIndex; i++) { // Only loop through rows up to the new row index
    if (sheet.getRange(i, getColIndexByName("Status")).getValue() == "New") {
      numNew++;
    }
  }

  MailApp.sendEmail(userEmail,
                    newRowIndex + " has been created", // Modified subject
                    "Dear " + fullName + ",\n\nThank you for contacting G6line Tech support. We have received your support request with the following details:\n\n" +
                    "Short Description: " + shortDescription + "\n" +
                    "Detailed Description: " + detailedDescription + "\n" +
                    "Severity: " + severity + "\n\n" +
                    "Our support team will begin working on your request as soon as possible. Requests are prioritized based on urgency and type of request. You are currently number " +
                    (numNew + 1) + " in the queue.\n\nSincerely," +
                    "\n\nThe G6line Tech Support Team"); // Modified body
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

function emailStatusUpdates() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var row = sheet.getActiveRange().getRowIndex();
  var lastRow = sheet.getLastRow();
  var fullName = sheet.getRange(row, getColIndexByName("Full Name")).getValue(); // Added full name
  var shortDescription = sheet.getRange(row, getColIndexByName("Short Description")).getValue(); // Added short description
  var detailedDescription = sheet.getRange(row, getColIndexByName("Detailed Description")).getValue(); // Added detailed description
  var severity = sheet.getRange(row, getColIndexByName("Severity")).getValue(); // Added severity
  var helpdeskNotes = sheet.getRange(row, getColIndexByName("Help Desk Notes")).getValue(); // Added customer notes
  var status = sheet.getRange(row, getColIndexByName("Status")).getValue(); // Added status
  var resolution = sheet.getRange(row, getColIndexByName("Resolution")).getValue(); // Added resolution
  var userEmail = sheet.getRange(row, getColIndexByName("Email Address")).getValue(); // Email Address
  var subject = "Ticket Update: HDT#" + lastRow; // Modified subject
  var body = "Dear " + fullName + ",\n\nThe G6line Tech support team has updated the status of your support request with the following details:\n\n" +
"Short Description: " + shortDescription + "\n" +
"Detailed Description: " + detailedDescription + "\n" +
"Severity: " + severity + "\n" +
"Help Desk Notes: " + helpdeskNotes + "\n" +
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
