function setupTestResultsSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let testSheet = ss.getSheetByName('testResults');

  if (!testSheet) {
    testSheet = ss.insertSheet('testResults');
    testSheet.getRange('A1:C1').setValues([['Test Name', 'Result', 'Details']]);
    testSheet.getRange('A1:C1').setFontWeight('bold');
    testSheet.setFrozenRows(1);
    testSheet.autoResizeColumns(1, 3);
    testSheet.setColumnWidth(3, 300);
  }
  return testSheet;
}

function clearTestResults() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let testSheet = ss.getSheetByName('testResults');

  if (testSheet) {
    const lastRow = testSheet.getLastRow();
    if (lastRow > 1) {
      testSheet.getRange(2, 1, lastRow - 1, 3).clearContent();
      testSheet.getRange(2, 1, lastRow - 1, 3).clearFormat();
    }
  }
}

function recordTestResult(testName, passed, details) {
  const testSheet = setupTestResultsSheet();
  const lastRow = Math.max(1, testSheet.getLastRow());

  testSheet.getRange(lastRow + 1, 1, 1, 3).setValues([
    [
      testName,
      passed ? 'PASSED' : 'FAILED',
      details
    ]
  ]);

  const resultCell = testSheet.getRange(lastRow + 1, 2);
  resultCell.setBackground(passed ? '#b7e1cd' : '#f4c7c3');

  testSheet.autoResizeColumns(1, 3);
}

function isRangeProtected(sheet, rangeA1) {
  const protections = sheet.getProtections(SpreadsheetApp.ProtectionType.RANGE);
  return protections.some(protection => {
    const protectedRange = protection.getRange();
    return protectedRange.getA1Notation() === rangeA1;
  });
}

function createMockEditEvent(spreadsheet, range, newValue, oldValue) {
  range.setValue(newValue);

  return {
    source: spreadsheet,
    range: range,
    value: newValue,
    oldValue: oldValue
  };
}




// Test that Show Answer checkbox is enabled when quiz is started
function testShowAnswerCheckboxEnabledWhenQuizStarted() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const quizSheet = ss.getSheetByName('quiz');
  const datastoreSheet = ss.getSheetByName('datastore');
  
  // Get a valid category
  const data = datastoreSheet.getDataRange().getValues();
  const validCategories = [...new Set(data.slice(1).map(row => row[1]).filter(Boolean))];
  const testCategory = validCategories[0];
  
  // Setup quiz
  quizSheet.getRange('A1').setValue(testCategory);
  quizSheet.getRange('B2').setValue(false);
  quizSheet.getRange('B7').setValue(false);
  
  // Start quiz
  const startQuizEvent = {
    source: ss,
    range: quizSheet.getRange('B2'),
    value: true,
    oldValue: false
  };
  handleCheckboxEdit(startQuizEvent);
  
  // Check if Show Answer checkbox is not protected (enabled)
  const isShowAnswerProtected = isRangeProtected(quizSheet, 'B7');
  
  recordTestResult(
    'Show Answer checkbox should be enabled when quiz is started',
    !isShowAnswerProtected,
    !isShowAnswerProtected ? 
      '✓ Show Answer checkbox is properly enabled when quiz is started' :
      '✗ Show Answer checkbox is still disabled when quiz is started'
  );
  
  return !isShowAnswerProtected;
}





// Test that Show Answer updates with new question when moving to next question
function testShowAnswerUpdatesWithNextQuestion() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const quizSheet = ss.getSheetByName('quiz');
  const datastoreSheet = ss.getSheetByName('datastore');
  
  // Get a category with multiple questions
  const data = datastoreSheet.getDataRange().getValues();
  const categoryQuestionCounts = {};
  
  data.slice(1).forEach(row => {
    if (row[1]) {
      categoryQuestionCounts[row[1]] = (categoryQuestionCounts[row[1]] || 0) + 1;
    }
  });
  
  let testCategory = null;
  for (const [category, count] of Object.entries(categoryQuestionCounts)) {
    if (count >= 2) {
      testCategory = category;
      break;
    }
  }
  
  if (!testCategory) {
    recordTestResult(
      'Show Answer should update with new question when moving to next question',
      false,
      '✗ No category with multiple questions found for testing'
    );
    return false;
  }
  
  // Setup quiz
  quizSheet.getRange('A1').setValue(testCategory);
  quizSheet.getRange('B2').setValue(false);
  quizSheet.getRange('B7').setValue(false);
  quizSheet.getRange('C1').setValue(0); // Reset counter
  quizSheet.getRange('D1').setValue(''); // Clear used questions
  
  // Start quiz
  const startQuizEvent = {
    source: ss,
    range: quizSheet.getRange('B2'),
    value: true,
    oldValue: false
  };
  handleCheckboxEdit(startQuizEvent);
  
  // Check Show Answer checkbox
  const showAnswerEvent = {
    source: ss,
    range: quizSheet.getRange('B7'),
    value: true,
    oldValue: false
  };
  handleCheckboxEdit(showAnswerEvent);
  
  // Get first question and answer
  const firstQuestion = quizSheet.getRange('A4').getValue();
  const firstAnswer = quizSheet.getRange('A8').getValue();
  
  // Move to next question by clicking Right
  const rightClickEvent = {
    source: ss,
    range: quizSheet.getRange('B5'),
    value: true,
    oldValue: false
  };
  handleCheckboxEdit(rightClickEvent);
  
  // Get second question and answer
  const secondQuestion = quizSheet.getRange('A4').getValue();
  const secondAnswer = quizSheet.getRange('A8').getValue();
  
  const testPassed = firstQuestion !== secondQuestion && 
                     firstAnswer !== secondAnswer && 
                     firstAnswer !== '' && 
                     secondAnswer !== '' &&
                     !secondQuestion.includes('Quiz Complete');
  
  recordTestResult(
    'Show Answer should update with new question when moving to next question',
    testPassed,
    testPassed ? 
      '✓ Show Answer correctly updates with new question' :
      `✗ Show Answer update failed. First Q: "${firstQuestion}", Second Q: "${secondQuestion}", First A: "${firstAnswer}", Second A: "${secondAnswer}"`
  );
  
  return testPassed;
}

// Test that Show Answer is cleared when quiz ends
function testShowAnswerClearedWhenQuizEnds() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const quizSheet = ss.getSheetByName('quiz');
  const datastoreSheet = ss.getSheetByName('datastore');
  
  // Get a valid category
  const data = datastoreSheet.getDataRange().getValues();
  const validCategories = [...new Set(data.slice(1).map(row => row[1]).filter(Boolean))];
  const testCategory = validCategories[0];
  
  // Setup quiz
  quizSheet.getRange('A1').setValue(testCategory);
  quizSheet.getRange('B2').setValue(false);
  quizSheet.getRange('B7').setValue(false);
  quizSheet.getRange('C1').setValue(0); // Reset counter
  quizSheet.getRange('D1').setValue(''); // Clear used questions
  
  // Start quiz
  const startQuizEvent = {
    source: ss,
    range: quizSheet.getRange('B2'),
    value: true,
    oldValue: false
  };
  handleCheckboxEdit(startQuizEvent);
  
  // Check Show Answer checkbox
  const showAnswerEvent = {
    source: ss,
    range: quizSheet.getRange('B7'),
    value: true,
    oldValue: false
  };
  handleCheckboxEdit(showAnswerEvent);
  
  // Verify answer is displayed
  const answerBeforeStop = quizSheet.getRange('A8').getValue();
  
  // Stop quiz
  const stopQuizEvent = {
    source: ss,
    range: quizSheet.getRange('B2'),
    value: false,
    oldValue: true
  };
  handleCheckboxEdit(stopQuizEvent);
  
  // Check answer after stopping quiz
  const answerAfterStop = quizSheet.getRange('A8').getValue();
  
  const testPassed = answerBeforeStop !== '' && answerAfterStop === '';
  
  recordTestResult(
    'Show Answer should be cleared when quiz ends',
    testPassed,
    testPassed ? 
      '✓ Show Answer is correctly cleared when quiz ends' :
      `✗ Show Answer clear failed. Before stop: "${answerBeforeStop}", After stop: "${answerAfterStop}"`
  );
  
  return testPassed;
}

/**
 * Run all unit tests and record results in the test results sheet
 */
function runAllTestsGit() {
  clearTestResults();

  // Existing tests
  
  // New Show Answer tests

  testShowAnswerCheckboxEnabledWhenQuizStarted();
  //testShowAnswerDisplaysCorrectAnswer();
  //testShowAnswerHidesAnswerWhenUnchecked();
  testShowAnswerUpdatesWithNextQuestion();
  testShowAnswerClearedWhenQuizEnds();
}