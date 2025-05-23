// Unit Tests for Quiz Tool
// This file contains separate unit tests that can be run independently

/**
 * Set up the test results sheet if it doesn't exist
 * @returns {GoogleAppsScript.Spreadsheet.Sheet} The test results sheet
 */
function setupTestResultsSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let testSheet = ss.getSheetByName('testResults');

  // Create the sheet if it doesn't exist
  if (!testSheet) {
    testSheet = ss.insertSheet('testResults');

    // Set up header row only when creating a new sheet
    testSheet.getRange('A1:C1').setValues([['Test Name', 'Result', 'Details']]);
    testSheet.getRange('A1:C1').setFontWeight('bold');

    // Format the sheet
    testSheet.setFrozenRows(1);
    testSheet.autoResizeColumns(1, 3);
    testSheet.setColumnWidth(3, 300); // Make the details column wider
  }
  return testSheet;
}

/**
 * Clear all test results in the test results sheet
 */
function clearTestResults() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let testSheet = ss.getSheetByName('testResults');

  if (testSheet) {
    // Clear existing content except header
    const lastRow = testSheet.getLastRow();
    if (lastRow > 1) {
      testSheet.getRange(2, 1, lastRow - 1, 3).clearContent();
      testSheet.getRange(2, 1, lastRow - 1, 3).clearFormat();
    }
  }
}

/**
 * Record a test result in the test results sheet
 * @param {string} testName - Name of the test
 * @param {boolean} passed - Whether the test passed
 * @param {string} details - Additional details about the test result
 */
function recordTestResult(testName, passed, details) {
  const testSheet = setupTestResultsSheet();
  const lastRow = Math.max(1, testSheet.getLastRow());

  // Add new row with test results
  testSheet.getRange(lastRow + 1, 1, 1, 3).setValues([
    [
      testName,
      passed ? 'PASSED' : 'FAILED',
      details
    ]
  ]);

  // Set background color based on result
  const resultCell = testSheet.getRange(lastRow + 1, 2);
  resultCell.setBackground(passed ? '#b7e1cd' : '#f4c7c3'); // Light green for pass, light red for fail

  // Auto-resize columns to fit content
  testSheet.autoResizeColumns(1, 3);
}
/**
 * Helper function to check if a range is protected
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - The sheet to check
 * @param {string} rangeA1 - The A1 notation of the range to check
 * @returns {boolean} True if the range is protected, false otherwise
 */
function isRangeProtected(sheet, rangeA1) {
  const protections = sheet.getProtections(SpreadsheetApp.ProtectionType.RANGE);
  return protections.some(protection => {
    const protectedRange = protection.getRange();
    return protectedRange.getA1Notation() === rangeA1;
  });
}

/**
 * Test that verifies questions are not repeated during a quiz session
 * Returns true if the test passes, false otherwise
 */
function testQuestionsNotRepeated() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const quizSheet = ss.getSheetByName('quiz');
  const datastoreSheet = ss.getSheetByName('datastore');
  
  // Get valid category from datastore with multiple questions
  const data = datastoreSheet.getDataRange().getValues();
  const validCategories = [...new Set(data.slice(1).map(row => row[1]).filter(Boolean))];
  const testCategory = validCategories.length > 0 ? validCategories[0] : 'Rishis';
  
  // Get all questions for the test category
  const categoryQuestions = data.filter((row, index) => index !== 0 && row[1] === testCategory);
  
  // Reset the quiz state
  quizSheet.getRange('A1').setValue(testCategory);
  quizSheet.getRange('B2').setValue(false);
  quizSheet.getRange('A4').setValue('');
  quizSheet.getRange('C1').setValue(0);
  quizSheet.getRange('D1').setValue(''); // Reset used questions
  quizSheet.getRange('B5').setValue(false);
  quizSheet.getRange('B6').setValue(false);
  
  // Start the quiz
  const startQuizEvent = {
    source: ss,
    range: quizSheet.getRange('B2'),
    value: true,
    oldValue: false
  };
  handleCheckboxEdit(startQuizEvent);
  
  const usedQuestions = [];
  let hasRepeats = false;
  let debugInfo = [];
  
  // Collect the first question
  const firstQuestion = quizSheet.getRange('A4').getValue();
  if (firstQuestion && firstQuestion !== '') {
    usedQuestions.push(firstQuestion);
    debugInfo.push(`Q1: "${firstQuestion}"`);
  }
  
  // Simulate answering questions and check for repetition
  const maxQuestions = Math.min(5, categoryQuestions.length);
  for (let i = 1; i < maxQuestions; i++) {
    // Answer the current question (use Right checkbox)
    const rightClickEvent = {
      source: ss,
      range: quizSheet.getRange('B5'),
      value: true,
      oldValue: false
    };
    handleCheckboxEdit(rightClickEvent);
    
    const currentQuestion = quizSheet.getRange('A4').getValue();
    
    // Check if this is a completion message
    if (currentQuestion && currentQuestion.toString().includes('Quiz Complete')) {
      debugInfo.push(`Quiz completed at question ${i + 1}`);
      break;
    }
    
    if (currentQuestion && currentQuestion !== '') {
      debugInfo.push(`Q${i + 1}: "${currentQuestion}"`);
      
      // Check if this question was already used
      if (usedQuestions.includes(currentQuestion)) {
        hasRepeats = true;
        debugInfo.push(`REPEAT DETECTED: "${currentQuestion}" was already used`);
      } else {
        usedQuestions.push(currentQuestion);
      }
    }
  }
  
  const testPassed = !hasRepeats && usedQuestions.length > 1;
  
  // Record result in the test sheet
  recordTestResult(
    'Questions should not be repeated during a quiz session', 
    testPassed, 
    testPassed ? 
      `✓ No repeated questions found. Used ${usedQuestions.length} unique questions` : 
      `✗ ${hasRepeats ? 'Repeated questions detected' : 'Not enough questions to test'}. Debug: ${debugInfo.join(' | ')}`
  );
  
  return testPassed;
}
/**
 * Run all unit tests and record results in the test results sheet
 */
function runAllTestsGit() {
  // Clear previous results before running new tests
  clearTestResults();
  // Test 1: Category change clears question cell
  testQuestionsNotRepeated();
}

