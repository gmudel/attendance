function markAttendance() {
  var ui = SpreadsheetApp.getUi();
  var formUrl = ui.prompt('Please enter the form link').getResponseText();
  var weekString = ui.prompt('Please enter the week').getResponseText();
  var week = parseInt(weekString);
  try {
    var form = FormApp.openByUrl(formUrl);
  }
  catch(e) {
    ui.alert("Error opening form. Make sure you're the form owner and that you're using the form's edit link (not the public one that you gave to students). Contact Gabe Mudel if the issue persists.");
    return;
  }
  var sheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheetData = sheet.getDataRange().getValues();

  var notFound = [];
  var formResponses = form.getResponses();
  for (var i = 0; i < formResponses.length; i++) {
    var formResponse = formResponses[i];
    var email = formResponse.getRespondentEmail()
    var timestamp = formResponse.getTimestamp()
    if(!markStudent_(sheetData, email, timestamp, week)) {
      notFound.push(email);
    }
  }
  markUnenrolledStudents_(sheet, sheetData, notFound, week);
}

function markUnenrolledStudents_(sheet, sheetData, notFound, week) {
  var notFoundRow = getNotFoundIndex_(sheetData);
  for(var i = 0; i < notFound.length; ++i) {
    a1 = toa1_(notFoundRow + i, week);
    sheet.getRange(a1).setValue(notFound[i]);
  }
}

function getNotFoundIndex_(data) {
  let idx = data.findIndex(row => row[0] === "Not found:");
  if(idx == -1) {
    idx = data.length + 1;
    a1 = toa1_(idx, 0);
    SpreadsheetApp.getActiveSpreadsheet().getRange(a1).setValue("Not found:")
  }
  return idx;
}

function markStudent_(data, studentEmail, timestamp, week) {
  row = data.findIndex(row => row[0] === studentEmail);
  if (row != -1) {
    a1Notation = toa1_(row, week);
    SpreadsheetApp.getActiveSpreadsheet().getRange(a1Notation).setValue(timestamp);
    return true;
  }
  else {
    return false;
  }
}
function toa1_(row, col){
  return String.fromCharCode(65 + col) + (++row).toString();
}
