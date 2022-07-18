function showFeedbackDialog() {
  var widget = HtmlService.createHtmlOutputFromFile("Index.html");
  widget.setHeight(150);
  widget.setWidth(200);
  SpreadsheetApp.getUi().showModalDialog(widget, "Create Sheet");
}

function appendData(data) {
  var name = data.new_date;
  var prev_date = data.past_date;
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Template').copyTo(ss);

  sheet.setName(name);

  /* Make the new sheet active */
  ss.setActiveSheet(sheet);
  ss.moveActiveSheet(2);
  sheet.getRange("Y2").setValue(name);
  let curr_date = new Date(sheet.getRange("Y2").getValue());
  let day = curr_date.getDay();
  switch (day) {
    case 1:
      day = 'Mon';
      break;
    case 2:
      day = 'Tue';
      break;
    case 3:
      day = 'Wed';
      break;
    case 4:
      day = 'Thurs';
      break;
    case 5:
      day = 'Fri';
      break;
    default:
      day = "Invalid"
      break;
  }
  sheet.getRange("AB2").setValue(day);

  const prevSheet = ss.getSheetByName(prev_date);
  const vals = prevSheet.getRange("N13:N34").getValues();

  sheet.getRange("D13:D34").setValues(vals);

}
