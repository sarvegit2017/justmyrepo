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
  
  // Set up trigger for onEdit
  ScriptApp.newTrigger("onEditQuizSheet")
    .forSpreadsheet(ss)
    .onEdit()
    .create();
  
  Logger.log('Setup complete. Dropdown, headers, and trigger are ready.');
}

function onEditQuizSheet(e) {
  var ss = e.source;
  var quizSheet = ss.getSheetByName("quiz");
  var datastoreSheet = ss.getSheetByName("datastore");
  
  if (!quizSheet || !datastoreSheet) return;
  
  var editedRange = e.range;
  if (quizSheet.getName() == e.range.getSheet().getName() && editedRange.getA1Notation() == "B2") {
    var selectedCategory = e.value;
    
    // Clear display area first (from row 3, cols B to D)
    quizSheet.getRange("B3:D").clearContent();
    quizSheet.getRange("B4:D").clearContent();
    
    // Always set headers
    quizSheet.getRange("B3").setValue("SL#");
    quizSheet.getRange("C3").setValue("Category");
    quizSheet.getRange("D3").setValue("Questions");
    
    if (!selectedCategory) return;
    
    var data = datastoreSheet.getDataRange().getValues();
    var filteredData = [];
    
    for (var i = 1; i < data.length; i++) {
      if (data[i][1] == selectedCategory) {
        filteredData.push([data[i][0], data[i][1], data[i][2]]);
      }
    }
    
    if (filteredData.length > 0) {
      if (filteredData.length > 5) {
        filteredData = getRandomSample(filteredData, 5);
      }
      quizSheet.getRange(4, 2, filteredData.length, 3).setValues(filteredData);
    } else {
      quizSheet.getRange("B4").setValue("No data for selected category.");
    }
  }
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
