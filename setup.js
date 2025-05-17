function setupQuizSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var quizSheet = ss.getSheetByName('Quiz');
  
  if (!quizSheet) {
    quizSheet = ss.insertSheet('Quiz');
  } else {
    quizSheet.clear();
  }

  // Setup UI layout
  quizSheet.getRange('B2').setValue('Select Category');
  quizSheet.getRange('B4').setValue('Question');
  quizSheet.getRange('B6').setValue('Next Question');
  quizSheet.getRange('B8').setValue('Show Answer');
  quizSheet.getRange('B10').setValue('Answer');

  quizSheet.getRange('E1').setValue('Questions Asked:');
  quizSheet.getRange('F1').setValue(0);

  quizSheet.getRange('B13').setValue('Right Count:');
  quizSheet.getRange('C13').setValue(0);
  quizSheet.getRange('B14').setValue('Right');
  quizSheet.getRange('C14').insertCheckboxes();

  quizSheet.getRange('B15').setValue('Wrong Count:');
  quizSheet.getRange('C15').setValue(0);
  quizSheet.getRange('B16').setValue('Wrong');
  quizSheet.getRange('C16').insertCheckboxes();

  quizSheet.getRange('C6').insertCheckboxes();
  quizSheet.getRange('C8').insertCheckboxes();

  // Dropdown: category selection
  var dataSheet = ss.getSheetByName('datastore');
  var lastRow = dataSheet.getLastRow();
  var categoryRange = dataSheet.getRange('B2:B' + lastRow);
  var rule = SpreadsheetApp.newDataValidation()
    .requireValueInRange(categoryRange, true)
    .setAllowInvalid(false)
    .build();
  quizSheet.getRange('C2').setDataValidation(rule);

  quizSheet.getRange('C2').clearContent();
  quizSheet.getRange('C4').clearContent();
  quizSheet.getRange('C6').setValue(false);
  quizSheet.getRange('C8').setValue(false);
  quizSheet.getRange('C10').clearContent();
  quizSheet.getRange('C14').setValue(false);
  quizSheet.getRange('C16').setValue(false);
  quizSheet.getRange('Z1').clearContent();
}
