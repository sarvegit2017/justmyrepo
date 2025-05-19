function setupQuizSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var quizSheet = ss.getSheetByName('Quiz');
  
  if (!quizSheet) {
    quizSheet = ss.insertSheet('Quiz');
  } else {
    quizSheet.clear();
  }

  // Set column widths for better readability
  quizSheet.setColumnWidth(1, 20);  // Margin column
  quizSheet.setColumnWidth(2, 120); // Label column
  quizSheet.setColumnWidth(3, 300); // Content column
  quizSheet.setColumnWidth(4, 30);  // Spacing column
  quizSheet.setColumnWidth(5, 120); // Stats label column
  quizSheet.setColumnWidth(6, 80);  // Stats value column

  // Title and instructions
  quizSheet.getRange('B1:C1').merge();
  quizSheet.getRange('B1').setValue('QUIZ TOOL').setFontWeight('bold').setHorizontalAlignment('center');
  quizSheet.getRange('B1:C1').setBackground('#f0f0f0');
  
  // Create sections with borders and backgrounds
  // Setup section
  quizSheet.getRange('B2:C3').setBorder(true, true, true, true, true, true);
  quizSheet.getRange('B2').setValue('Select Category').setFontWeight('bold');
  quizSheet.getRange('B3').setValue('Start Quiz').setFontWeight('bold');
  
  // Question section
  quizSheet.getRange('B5:C5').merge();
  quizSheet.getRange('B5').setValue('CURRENT QUESTION').setFontWeight('bold').setHorizontalAlignment('center');
  quizSheet.getRange('B5:C5').setBackground('#e6f2ff');
  
  quizSheet.getRange('B6:C8').setBorder(true, true, true, true, true, true);
  quizSheet.getRange('B6:C8').setBackground('#f5f9ff');
  quizSheet.getRange('B6').setValue('Question:').setFontWeight('bold');
  
  // Answer section
  quizSheet.getRange('B10:C10').merge();
  quizSheet.getRange('B10').setValue('ANSWER').setFontWeight('bold').setHorizontalAlignment('center');
  quizSheet.getRange('B10:C10').setBackground('#e6f2ff');
  
  quizSheet.getRange('B11:C13').setBorder(true, true, true, true, true, true);
  quizSheet.getRange('B11:C13').setBackground('#f5f9ff');
  quizSheet.getRange('B11').setValue('Show Answer').setFontWeight('bold');
  quizSheet.getRange('B12').setValue('Answer:').setFontWeight('bold');
  
  // Navigation section
  quizSheet.getRange('B15:C15').merge();
  quizSheet.getRange('B15').setValue('NAVIGATION').setFontWeight('bold').setHorizontalAlignment('center');
  quizSheet.getRange('B15:C15').setBackground('#e6f2ff');
  
  quizSheet.getRange('B16:C16').setBorder(true, true, true, true, true, true);
  quizSheet.getRange('B16:C16').setBackground('#f5f9ff');
  quizSheet.getRange('B16').setValue('Next Question').setFontWeight('bold');
  
  // Results section
  quizSheet.getRange('B18:C18').merge();
  quizSheet.getRange('B18').setValue('RESULTS').setFontWeight('bold').setHorizontalAlignment('center');
  quizSheet.getRange('B18:C18').setBackground('#e6f2ff');
  
  quizSheet.getRange('B19:C22').setBorder(true, true, true, true, true, true);
  quizSheet.getRange('B19:C22').setBackground('#f5f9ff');
  quizSheet.getRange('B19').setValue('Right').setFontWeight('bold');
  quizSheet.getRange('B20').setValue('Right Count:').setFontWeight('bold');
  quizSheet.getRange('B21').setValue('Wrong').setFontWeight('bold');
  quizSheet.getRange('B22').setValue('Wrong Count:').setFontWeight('bold');

  // Quiz statistics in the right column
  quizSheet.getRange('E2:F2').merge();
  quizSheet.getRange('E2').setValue('QUIZ STATISTICS').setFontWeight('bold').setHorizontalAlignment('center');
  quizSheet.getRange('E2:F2').setBackground('#e6f2ff');
  
  quizSheet.getRange('E3:F5').setBorder(true, true, true, true, true, true);
  quizSheet.getRange('E3:F5').setBackground('#f5f9ff');
  quizSheet.getRange('E3').setValue('Questions Asked:').setFontWeight('bold');
  quizSheet.getRange('E4').setValue('Progress:').setFontWeight('bold');
  quizSheet.getRange('E5').setValue('Score:').setFontWeight('bold');
  quizSheet.getRange('F5').setValue('0%');

  // Add checkboxes and set initial values
  quizSheet.getRange('C3').insertCheckboxes();
  quizSheet.getRange('C11').insertCheckboxes();
  quizSheet.getRange('C16').insertCheckboxes();
  quizSheet.getRange('C19').insertCheckboxes();
  quizSheet.getRange('C21').insertCheckboxes();

  // Dropdown: category selection
  var dataSheet = ss.getSheetByName('datastore');
  var lastRow = dataSheet.getLastRow();
  var categoryRange = dataSheet.getRange('B2:B' + lastRow);
  var rule = SpreadsheetApp.newDataValidation()
    .requireValueInRange(categoryRange, true)
    .setAllowInvalid(false)
    .build();
  quizSheet.getRange('C2').setDataValidation(rule);

  // Initialize values
  quizSheet.getRange('C2').clearContent();
  quizSheet.getRange('C3').setValue(false);
  quizSheet.getRange('C6').clearContent();
  quizSheet.getRange('C11').setValue(false);
  quizSheet.getRange('C12').clearContent();
  quizSheet.getRange('C16').setValue(false);
  quizSheet.getRange('C19').setValue(false);
  quizSheet.getRange('C20').setValue(0);
  quizSheet.getRange('C21').setValue(false);
  quizSheet.getRange('C22').setValue(0);
  quizSheet.getRange('F3').setValue(0);
  quizSheet.getRange('F4').setValue('0/5');
  
  // Store answer in a hidden cell
  quizSheet.getRange('Z1').clearContent();
}