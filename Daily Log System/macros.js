/**
 * The AddButton() and AddButton2() macros have the same functionality with different ranges.
 * Determines the row based on dropdown menu and then adds the "BOX" and "PACK" values into its
 * corresponding row.
 */
function AddButton() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const name = spreadsheet.getRange("AD7").getValue();
  const vals = spreadsheet.getRange("AH7:AK7").getValues();

  const productVals = spreadsheet.getRange("A18:C29").getValues();

  Logger.log(name);
  Logger.log("Box: " + vals[0][0] + " Packs: " + vals[0][3]);

  for (row in productVals) {
    let realRow = 18 + parseInt(row);
    if (productVals[row][0].toString() === name.toString()) {
      spreadsheet.getRange(realRow, 6, 1, 3).setValues([[vals[0][0], [], vals[0][3]]]);
    }
  }
};

function AddButton2() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const name = spreadsheet.getRange("AD8").getValue();
  const vals = spreadsheet.getRange("AH8:AK8").getValues();

  const productVals = spreadsheet.getRange("A18:C29").getValues();

  Logger.log(name);
  Logger.log("Box: " + vals[0][0] + " Packs: " + vals[0][3]);

  for (row in productVals) {
    let realRow = 18 + parseInt(row);
    if (productVals[row][0].toString() === name.toString()) {
      spreadsheet.getRange(realRow, 6, 1, 3).setValues([[vals[0][0], [], vals[0][3]]]);
    }
  }
};

/**
 * This macro functions as a button, displaying a pop-up dialog which collects user-input.
 */
function CreateSheet() {
  showFeedbackDialog();
};