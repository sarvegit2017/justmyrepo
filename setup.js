//mycode
function setupQuizSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var quizSheetName = "quiz";
  var datastoreSheetName = "datastore";
  
  // Ensure 'quiz' sheet exists
  var quizSheet = ss.getSheetByName(quizSheetName);
  if (!quizSheet) {
    quizSheet = ss.insertSheet(quizSheetName);
    Logger.log('Sheet "' + quizSheetName + '" has been created.');
  } else {
    Logger.log('Sheet "' + quizSheetName + '" already exists.');
  }
  
  // Setup dropdown in B2
  var datastoreSheet = ss.getSheetByName(datastoreSheetName);
  if (!datastoreSheet) {
    Logger.log('Sheet "' + datastoreSheetName + '" not found.');
    return;
  }
  
  var data = datastoreSheet.getDataRange().getValues();
  var categories = [];
  for (var i = 1; i < data.length; i++) {
    if (data[i][1]) {
      categories.push(data[i][1]);
    }
  }
  var uniqueCategories = Array.from(new Set(categories));
  
  var dropdownCell = quizSheet.getRange("B2");
  var rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(uniqueCategories, true)
    .setAllowInvalid(false)
    .build();
  dropdownCell.setDataValidation(rule);
  
  // Clear any existing trigger to avoid duplication
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() == "onEditQuizSheet") {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
  
  Logger.log('Setup complete. Dropdown and headers are ready.');
}

// Helper function to get random sample of size 'n' from array 'arr'
function getRandomSample(arr, n) {
  var result = [];
  var taken = [];
  
  while (result.length < n) {
    var index = Math.floor(Math.random() * arr.length);
    if (taken.indexOf(index) === -1) {
      result.push(arr[index]);
      taken.push(index);
    }
  }
  
  return result;
}