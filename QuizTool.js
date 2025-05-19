/**
 * Quiz Tool - Main Script
 * 
 * This script creates an interactive quiz using Google Sheets.
 * Questions are pulled from a 'datastore' sheet and presented
 * in a clean user interface that tracks progress and scores.
 */

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Quiz Tool')
      .addItem('Setup Quiz Sheet', 'setupQuizSheet')
      .addToUi();
}

// Call this function to create or reset the Quiz sheet
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
  
  // Setup section
  quizSheet.getRange('B2:C3').setBorder(true, true, true, true, true, true);
  quizSheet.getRange('B2').setValue('Select Category').setFontWeight('bold');
  quizSheet.getRange('B3').setValue('Start Quiz').setFontWeight('bold');
  
  // Question section
  quizSheet.getRange('B5:C5').merge();
  quizSheet.getRange('B5').setValue('CURRENT QUESTION').setFontWeight('bold').setHorizontalAlignment('center');
  quizSheet.getRange('B5:C5').setBackground('#e6f2ff');
  
  quizSheet.getRange('B6:C6').setBorder(true, true, true, true, true, true);
  quizSheet.getRange('B6:C6').setBackground('#f5f9ff');
  quizSheet.getRange('B6').setValue('Question:').setFontWeight('bold');
  
  // Answer section
  quizSheet.getRange('B8:C8').merge();
  quizSheet.getRange('B8').setValue('ANSWER').setFontWeight('bold').setHorizontalAlignment('center');
  quizSheet.getRange('B8:C8').setBackground('#e6f2ff');
  
  quizSheet.getRange('B9:C10').setBorder(true, true, true, true, true, true);
  quizSheet.getRange('B9:C10').setBackground('#f5f9ff');
  quizSheet.getRange('B9').setValue('Show Answer').setFontWeight('bold');
  quizSheet.getRange('B10').setValue('Answer:').setFontWeight('bold');
  
  // Results section (moved up to replace the navigation section)
  quizSheet.getRange('B12:C12').merge();
  quizSheet.getRange('B12').setValue('RESULTS').setFontWeight('bold').setHorizontalAlignment('center');
  quizSheet.getRange('B12:C12').setBackground('#e6f2ff');
  
  quizSheet.getRange('B13:C16').setBorder(true, true, true, true, true, true);
  quizSheet.getRange('B13:C16').setBackground('#f5f9ff');
  quizSheet.getRange('B13').setValue('Right').setFontWeight('bold');
  quizSheet.getRange('B14').setValue('Right Count:').setFontWeight('bold');
  quizSheet.getRange('B15').setValue('Wrong').setFontWeight('bold');
  quizSheet.getRange('B16').setValue('Wrong Count:').setFontWeight('bold');

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
  quizSheet.getRange('C9').insertCheckboxes();
  quizSheet.getRange('C13').insertCheckboxes();
  quizSheet.getRange('C15').insertCheckboxes();

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
  quizSheet.getRange('C9').setValue(false);
  quizSheet.getRange('C10').clearContent();
  quizSheet.getRange('C13').setValue(false);
  quizSheet.getRange('C14').setValue(0);
  quizSheet.getRange('C15').setValue(false);
  quizSheet.getRange('C16').setValue(0);
  quizSheet.getRange('F3').setValue(0);
  quizSheet.getRange('F4').setValue('0/5');
  
  // Store answer in a hidden cell
  quizSheet.getRange('Z1').clearContent();
  
  // Set up onEdit trigger
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === "onEdit") {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
  
  ScriptApp.newTrigger("onEdit")
    .forSpreadsheet(ss)
    .onEdit()
    .create();
    
  SpreadsheetApp.getUi().alert('Quiz has been set up successfully! Select a category and click "Start Quiz" to begin.');
}

function startQuiz() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var quizSheet = ss.getSheetByName('Quiz');
  
  var selectedCategory = quizSheet.getRange('C2').getValue();
  if (!selectedCategory) {
    SpreadsheetApp.getUi().alert('Please select a category first.');
    quizSheet.getRange('C3').setValue(false);
    return;
  }
  
  // Reset counters
  quizSheet.getRange('F3').setValue(0);  // Questions Asked
  quizSheet.getRange('F4').setValue('0/5'); // Progress
  quizSheet.getRange('F5').setValue('0%'); // Score
  quizSheet.getRange('C14').setValue(0); // Right Count
  quizSheet.getRange('C16').setValue(0); // Wrong Count
  
  // Show the first question
  showRandomQuestion();
  
  // Uncheck the Start Quiz checkbox
  quizSheet.getRange('C3').setValue(false);
}

function showRandomQuestion() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var quizSheet = ss.getSheetByName('Quiz');
  var dataSheet = ss.getSheetByName('datastore');

  // Check if we've already completed 5 questions
  var countCell = quizSheet.getRange('F3');  // Questions Asked cell
  var count = countCell.getValue();
  
  if (count >= 5) {
    quizSheet.getRange('F3').setValue("Quiz complete");
    quizSheet.getRange('C6').setValue("Quiz complete! You answered 5 questions.");
    
    // Calculate and show final score
    var rightCount = quizSheet.getRange('C14').getValue();
    var percentage = Math.round((rightCount / 5) * 100);
    quizSheet.getRange('F5').setValue(percentage + '%');
    quizSheet.getRange('C10').setValue('Your final score: ' + percentage + '% (' + rightCount + ' out of 5 correct)');
    return;
  }

  var selectedCategory = quizSheet.getRange('C2').getValue();
  if (!selectedCategory) {
    SpreadsheetApp.getUi().alert('Please select a category first.');
    return;
  }

  var data = dataSheet.getDataRange().getValues();
  var qaPairs = [];

  for (var i = 1; i < data.length; i++) {
    if (data[i][1] === selectedCategory) {
      qaPairs.push({ question: data[i][2], answer: data[i][3] });
    }
  }

  if (qaPairs.length === 0) {
    quizSheet.getRange('C6').setValue('No questions found in this category.');
    quizSheet.getRange('C10').setValue('');
    return;
  }

  var random = qaPairs[Math.floor(Math.random() * qaPairs.length)];
  quizSheet.getRange('C6').setValue(random.question);
  quizSheet.getRange('Z1').setValue(random.answer);
  quizSheet.getRange('C10').setValue('');

  // Update the question counter and progress
  var newCount = (!isNaN(count) ? count : 0) + 1;
  countCell.setValue(newCount);
  quizSheet.getRange('F4').setValue(newCount + '/5'); // Update progress

  // Reset checkboxes
  quizSheet.getRange('C9').setValue(false); // Show Answer
  quizSheet.getRange('C13').setValue(false); // Right
  quizSheet.getRange('C15').setValue(false); // Wrong
}

function showAnswer() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var quizSheet = ss.getSheetByName('Quiz');
  var answer = quizSheet.getRange('Z1').getValue();

  quizSheet.getRange('C10').setValue(answer || 'No answer available.');
  quizSheet.getRange('C9').setValue(false);
}

function updateCountAndAskNext(type) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var quizSheet = ss.getSheetByName('Quiz');
  
  // Update the appropriate counter based on right/wrong
  var cell = type === 'right' ? 'C14' : 'C16';
  var count = quizSheet.getRange(cell).getValue();
  quizSheet.getRange(cell).setValue((!isNaN(count) ? count : 0) + 1);

  // Update current score percentage
  var rightCount = quizSheet.getRange('C14').getValue();
  var questionsAsked = quizSheet.getRange('F3').getValue();
  var percentage = Math.round((rightCount / questionsAsked) * 100);
  quizSheet.getRange('F5').setValue(percentage + '%');

  // Reset the checkboxes
  quizSheet.getRange('C13').setValue(false); // Right checkbox
  quizSheet.getRange('C15').setValue(false); // Wrong checkbox

  // Check if we've completed 5 questions
  if (questionsAsked >= 5) {
    // Don't ask next question, just show completion message
    quizSheet.getRange('F3').setValue("Quiz complete");
    quizSheet.getRange('C6').setValue("Quiz complete! You answered 5 questions.");
    
    // Calculate and show final score
    percentage = Math.round((rightCount / 5) * 100);
    quizSheet.getRange('F5').setValue(percentage + '%');
    quizSheet.getRange('C10').setValue("Your final score: " + percentage + "% (" + rightCount + " out of 5 correct)");
    return;
  }

  // Automatically show the next question
  showRandomQuestion();
}

function onEdit(e) {
  var ss = e.source;
  var sheet = ss.getSheetByName('Quiz');
  var range = e.range;
  if (!sheet || sheet.getName() !== range.getSheet().getName()) return;

  var cell = range.getA1Notation();

  // Handle Start Quiz checkbox
  if (cell === 'C3' && range.getValue() === true) {
    startQuiz();
  }

  // Handle Show Answer checkbox
  if (cell === 'C9' && range.getValue() === true) {
    showAnswer();
  }

  // Handle Right checkbox
  if (cell === 'C13' && range.getValue() === true) {
    updateCountAndAskNext('right');
  }

  // Handle Wrong checkbox
  if (cell === 'C15' && range.getValue() === true) {
    updateCountAndAskNext('wrong');
  }
}