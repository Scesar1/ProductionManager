function createOnEditTrigger() {
  var triggers = ScriptApp.getProjectTriggers();
  var shouldCreateTrigger = true;
  triggers.forEach(function (trigger) {
    if(trigger.getEventType() === ScriptApp.EventType.ON_EDIT && trigger.getHandlerFunction() === "dataImport") {
      shouldCreateTrigger = false; 
    }
  });

   if(shouldCreateTrigger) {
    ScriptApp.newTrigger("dataImport")
      .forSpreadsheet(SpreadsheetApp.openById(spreadsheetId))
      .onEdit()
      .create()
  }
}
