
const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
let currRow = 3;


function addStockDialog() {
  var widget = HtmlService.createHtmlOutputFromFile("Dialog.html");
  widget.setHeight(150);
  widget.setWidth(200);
  SpreadsheetApp.getUi().showModalDialog(widget, "Add Stock");
}

function switchDialog() {
  var widget = HtmlService.createHtmlOutputFromFile("Switch.html");
  widget.setHeight(200);
  widget.setWidth(200);
  SpreadsheetApp.getUi().showModalDialog(widget, "Set Pouch");
}

function switchData(data) {
  determineRow();
  Logger.log(data.halfRadios);
  if (data.halfRadios.toString().trim() === "Aluminum") {
    getPrev(16);
    sheet.getRange(currRow - 1, 16).setFormula("=P" + (currRow - 2) + "-F" + (currRow - 1) + "*40 - H" + (currRow - 1));
    sheet.getRange(currRow - 1, 16).setFontWeight(null);
    sheet.getRange(currRow - 1, 16).setBackground(null);
    sheet.getRange(currRow - 1, 17).setValue(sheet.getRange(currRow - 2, 17).getValue());
  } else if (data.halfRadios.toString().trim() === "Clear") {
    getPrev(17);
    sheet.getRange(currRow - 1, 17).setFormula("=Q" + (currRow - 2) + "-F" + (currRow - 1) + "*40 - H" + (currRow - 1) + "+ G" + (currRow - 1) + "*40 - I" + (currRow - 1));
    sheet.getRange(currRow - 1, 17).setFontWeight(null);
    sheet.getRange(currRow - 1, 17).setBackground(null);
    sheet.getRange(currRow - 1, 16).setValue(sheet.getRange(currRow - 2, 16).getValue());
  }

  if (data.fullRadios.toString().trim() === "Aluminum") {
    getPrev(18);
    sheet.getRange(currRow - 1, 18).setFormula("=R" + (currRow - 2) + "- G" + (currRow - 1) + "*40 - I" + (currRow - 1));
    sheet.getRange(currRow - 1, 18).setFontWeight(null);
    sheet.getRange(currRow - 1, 19).setValue(sheet.getRange(currRow - 2, 19).getValue());
  } else if (data.fullRadios.toString().trim() === "Paper") {
    getPrev(19);
    sheet.getRange(currRow - 1, 19).setFormula("=S" + (currRow - 2) + "- G" + (currRow -1 ) + "*40 - I" + (currRow - 1));
    sheet.getRange(currRow - 1, 19).setFontWeight(null);
    sheet.getRange(currRow - 1, 18).setValue(sheet.getRange(currRow - 2, 18).getValue());
  }

}

function appendData(data) {
  
  SpreadsheetApp.getActiveSpreadsheet().toast(data.select + ": " + data.amount);

  determineRow();
  if(sheet.getRange(currRow, 1, 1, 25).isBlank()) {
    sheet.getRange(currRow - 1, 1, 2, 25).copyTo(sheet.getRange(currRow, 1, 1, 25), {contentsOnly:true});
    sheet.getRange(currRow, 1, 1, 25).setBackground("#e6b8af");
    sheet.getRange(currRow, 1).setValue("TOTAL");
    sheet.getRange(currRow, 1, 1, 25).setFontWeight("bold");
    getPrevTotal(16);
    getPrevTotal(17);
    getPrevTotal(18);
    getPrevTotal(19);
    sheet.getRange(currRow, 2, 1, 2).clearContent();
    sheet.getRange(currRow, 6, 1, 8).clearContent();

    sheet.getRange(4, 1, 1, 25).copyTo(sheet.getRange(currRow + 1, 1, 1, 25), {contentsOnly:false});
    sheet.getRange(currRow - 1, 16, 1, 4).copyTo(sheet.getRange(currRow + 1, 16, 1, 4), {contentsOnly: false});
    sheet.getRange(currRow + 1, 16, 1, 4).setBackground("#e6b8af");
  }
  
  let driedCol = parseInt(sheet.getRange(1421, 2).getValue());
  let updateCol = determineCol(data.select, driedCol);

  let prevTotal = parseInt(sheet.getRange(currRow - 1, updateCol).getValue());
  sheet.getRange(currRow, updateCol).setValue(prevTotal + parseInt(data.amount));
  
}


/**
 * Sends email to specified addresses once stock for products reach desginated threshold. 
 */
const sendEmail = () => {
  const stockVals = sheet.getRange(currRow - 1, 14, 1, 4).getValues();
  const silicaStock = sheet.getRange(currRow - 1, 23).getValue();
  let email = "sam@route66int.com";
  let email2 = "monica@route66int.com";
  let subject = "Low Stock Alert";
  let tail = ". \n\nIf you believe this email was a mistake, please contact the system administrator. \n\nThis is an auto-generated message, please do not reply.";
  let message = ""
  let sendEmail = false;
  if (stockVals[0][0] < 700) {
    message += "The stock of Half Box is currently at " + stockVals[0][0] + ", which is below the current threshold of 700.\n";
    sendEmail = true;
  }

  if (stockVals[0][1] < 1000) {
    message += "The stock of Full Box is currently at " + stockVals[0][1] + ", which is below the current threshold of 1000.\n";
    sendEmail = true;
  }

  if (stockVals[0][2] < 40000) {
    message += "The stock of Half Pouch is currently at " + stockVals[0][2] + ", which is below the current threshold of 40000.\n";
    sendEmail = true;
  }

  if (stockVals[0][3] < 5000) {
    message += "The stock of Full Pouch is currently at " + stockVals[0][3] + ", which is below the current threshold of 5000.\n";
    sendEmail = true;
  }

  if (silicaStock < 50000) {
    message += "The stock of Silica Gel is currently at " + silicaStock + ", which is below the current threshold of 50000.\n";
    sendEmail = true;
  }
  message += tail;
  if (sendEmail) {
    MailApp.sendEmail(email, subject, message);
    MailApp.sendEmail(email2, subject, message);
    Logger.log("Email Sent");
  }
}

const determineCol = (name, driedCol) => {

  if (name === "Dried Seaweed") {
    return driedCol + 2;
  }
  if (name === "Half Box") {
    return 14;
  }
  if (name === "Full Box") {
    return 15;
  }
  if (name === "Half Pouch") {
    return 16;
  }
  if (name === "Clear Pouch") {
    return 17;
  }
  if (name === "Full Pouch") {
    return 18;
  }
  if (name === "Full Paper") {
    return 19;
  }
  if (name === "Handrolled") {
    return 20;
  }
  if (name === "SHRD 260g") {
    return 21;
  }
  if (name === "SHRD 40g") {
    return 22;
  }
  if (name === "Silica Gel") {
    return 23;
  }
};


const determineRow = () => {
  const vals = sheet.getRange("A" + currRow + ":A").getValues();
  for (let row in vals) {
    if (vals[row][0].toString() == "") {
      break;
    }
    currRow++
    Logger.log(currRow);
  }
}

const getPrev = (col) => {
  let tempRow = currRow - 2;
  while (sheet.getRange(tempRow, col).getValue().toString() === "") {
    tempRow--;
  }
  let val = sheet.getRange(tempRow, col).getValue();
  Logger.log(val);
  sheet.getRange(currRow - 2, col).setValue(val);
}

const getPrevTotal = (col) => {
  let tempRow = currRow - 1;
  while (sheet.getRange(tempRow, col).getValue().toString() === "") {
    tempRow--;
  }
  let val = sheet.getRange(tempRow, col).getValue();
  Logger.log(val);
  sheet.getRange(currRow, col).setValue(val);
}


