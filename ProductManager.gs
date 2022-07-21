/*-----------------------------------------------------Global Variables----------------------------------------------------------*/

const spreadsheetId = "1NwOF-DMM6J0105uwfCpnuzy27f50K8Wa22_mMfm8RnA";
const invss = SpreadsheetApp.openById("1Z1mLhJc99480yT8R6Q8hRgvuakMHyVA8JQI7ikr9o8w");
const invsheet = invss.getSheets()[0];
var currRow = 5;

/*----------------------------------------------------- Functions ---------------------------------------------------------------*/
/**
 * Creates a pop-up dialog, connects it with Index.html file
 */
function showFeedbackDialog() {
  var widget = HtmlService.createHtmlOutputFromFile("Index.html");
  widget.setHeight(150);
  widget.setWidth(200);
  SpreadsheetApp.getUi().showModalDialog(widget, "Create Sheet");
}

/**
 * Retrieves data from the pop-up dialog form and uses data to create new blank copy of the template.
 */
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
  const vals = prevSheet.getRange("N14:N35").getValues();

  sheet.getRange("D14:D35").setValues(vals);
  dataImport();
}
/**
 * Imports the data from the current sheet to the Factory Inventory Spreadsheet. 
 */
function dataImport() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  if (sheet.getSheetName() === "Template") {
    Logger.log("template");
    return;
  }

  const date = new Date(sheet.getRange("Y2").getValue()).getTime();

  const dateVals = invsheet.getRange("A5:A").getValues();
  for (let row in dateVals) {
    const factDate = new Date(dateVals[row][0]).getTime();
    const val = dateVals[row][0].toString();
    if (date === factDate) {
      break;
    } else if (date < factDate) {
      Logger.log("Date doesn't exist in current log!");
      SpreadsheetApp.getActiveSpreadsheet().toast("Date doesn't exist in current inventory sheet");
    } else if (val === "" && isNaN(factDate)) {
      invsheet.getRange(currRow, 1).setValue(sheet.getRange("Y2").getValue());
      break;
    }
    currRow++;
  }

  const driedVals = sheet.getRange("AD3:AK3").getValues();
  let driedNameArr = driedVals[0][0].toString().split(" ");
  let driedName = driedNameArr[2];

  const driedPack = driedVals[0][7] / 60;
  const driedResult = driedVals[0][4] + driedPack;

  let driedCol = 2;

  const inputVals = invsheet.getRange("B2:C2").getValues();

  for (let col in inputVals[0]) {
    if (inputVals[0][col].toString().includes(driedName)) {
      driedCol += parseInt(col);
    }
  }
  invsheet.getRange(1421, 2).setValue(driedCol);

  invsheet.getRange(currRow, driedCol).setValue(driedResult);
  const outputFormula = invsheet.getRange(currRow - 1, driedCol + 2).getFormulaR1C1();
  invsheet.getRange(currRow, driedCol + 2).setFormulaR1C1(outputFormula);

  const boxFormulas = invsheet.getRange(currRow - 1, 14, 1, 7).getFormulasR1C1();
  invsheet.getRange(currRow, 14, 1, 7).setFormulasR1C1(boxFormulas);
  const silicaFormula = invsheet.getRange(currRow - 1, 23).getFormulaR1C1();
  invsheet.getRange(currRow, 23).setFormulaR1C1(silicaFormula);


  let halfBoxSum = 0;
  let halfPackSum = 0;
  let fullBoxSum = 0;
  let fullPackSum = 0;

  const productVal1 = sheet.getRange("AD7:AK7").getValues();
  if (productVal1[0][0].toString().includes("Half")) {
    halfBoxSum += productVal1[0][4];
    halfPackSum += productVal1[0][7];
  } else if (productVal1[0][0].toString().includes("Full")) {
    fullBoxSum += productVal1[0][4];
    fullPackSum += productVal1[0][7];
  } else if (productVal1[0][0].toString() === "Shredded 40g") {
    invsheet.getRange(currRow, 12, 1, 1).setValue(productVal1[0][4]);
    invsheet.getRange(currRow, 13, 1, 1).setValue(productVal1[0][7]);
  } else if (productVal1[0][0].toString() === "Shredded 200g") {
    invsheet.getRange(currRow, 10, 1, 1).setValue(productVal1[0][4]);
    invsheet.getRange(currRow, 11, 1, 1).setValue(productVal1[0][7]);
  }

  const productVal2 = sheet.getRange("AD8:AK8").getValues();
  if (productVal2[0][0].toString().includes("Half")) {
    halfBoxSum += productVal1[0][4];
    halfPackSum += productVal1[0][7];
  } else if (productVal2[0][0].toString().includes("Full")) {
    fullBoxSum += productVal1[0][4];
    fullPackSum += productVal1[0][7];
  } else if (productVal2[0][0].toString() === "Shredded 40g") {
    invsheet.getRange(currRow, 12, 1, 1).setValue(productVal2[0][4]);
    invsheet.getRange(currRow, 13, 1, 1).setValue(productVal2[0][7]);
  } else if (productVal2[0][0].toString() === "Shredded 200g") {
    invsheet.getRange(currRow, 10, 1, 1).setValue(productVal2[0][4]);
    invsheet.getRange(currRow, 11, 1, 1).setValue(productVal2[0][7]);
  }

  const productVal602 = sheet.getRange("AH9:AK9").getValues();
  halfPackSum += productVal602[0][0];
  fullPackSum += productVal602[0][3];

  const productValBulk = sheet.getRange("AH10").getValue();

  invsheet.getRange(currRow, 6, 1, 4).setValues([[halfBoxSum, fullBoxSum, halfPackSum, fullPackSum]]);

  invsheet.getRange(1, 1).setValue(currRow);

}

/**
 * Sends email to specified addresses once stock for products reach desginated threshold. 
 */
const sendEmail = () => {
  dataImport();
  const stockVals = invsheet.getRange(currRow, 14, 1, 4).getValues();
  const silicaStock = invsheet.getRange(currRow, 23).getValue();
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
    message += "The stock of Silica gel is currently at " + silicaStock + ", which is below the current threshold of 50000.\n";
    sendEmail = true;
  }
  message += tail;
  if (sendEmail) {
    MailApp.sendEmail(email, subject, message);
    MailApp.sendEmail(email, subject, message);
    Logger.log("Email Sent");
  }
}

