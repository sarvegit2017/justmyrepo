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
 * Test that verifies cell A4 is cleared when category in A1 is changed
 * Returns true if the test passes, false otherwise
 */
function testCategoryClearsQuestionCell() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const quizSheet = ss.getSheetByName('quiz');
  
  // Setup: Put some text in A4 to verify it gets cleared
  quizSheet.getRange('A4').setValue('Test question that should be cleared');
  quizSheet.getRange('B2').setValue(true); // Set checkbox to checked
  
  // Create a mock edit event for changing the category in A1
  const mockEvent = {
    source: ss,
    range: quizSheet.getRange('A1'),
    value: 'Some Category',
    oldValue: 'Previous Category'
  };
  
  // Call the handler function with our mock event
  handleCheckboxEdit(mockEvent);
  
  // Check if A4 is cleared
  const a4Value = quizSheet.getRange('A4').getValue();
  const a4Cleared = a4Value === '';
  
  // Record result in the test sheet
  recordTestResult(
    'When category is changed, cell A4 should be cleared', 
    a4Cleared, 
    a4Cleared ? 
      '✓ A4 cell was properly cleared' : 
      `✗ A4 cell was not cleared. Current value: "${a4Value}"`
  );
  
  return a4Cleared;
}

/**
 * Test that verifies checkbox in B2 is unchecked when category in A1 is changed
 * Returns true if the test passes, false otherwise
 */
function testCategoryClearsCheckbox() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const quizSheet = ss.getSheetByName('quiz');
  
  // Setup: Check the checkbox to verify it gets cleared
  quizSheet.getRange('B2').setValue(true); // Set checkbox to checked
  
  // Create a mock edit event for changing the category in A1
  const mockEvent = {
    source: ss,
    range: quizSheet.getRange('A1'),
    value: 'Some Category',
    oldValue: 'Previous Category'
  };
  
  // Call the handler function with our mock event
  handleCheckboxEdit(mockEvent);
  
  // Verify the checkbox was unchecked
  const checkboxValue = quizSheet.getRange('B2').getValue();
  const checkboxUnchecked = checkboxValue === false;
  
  // Record result in the test sheet
  recordTestResult(
    'When category is changed, checkbox in B2 should be cleared', 
    checkboxUnchecked, 
    checkboxUnchecked ? 
      '✓ B2 checkbox was properly unchecked' : 
      '✗ B2 checkbox was not unchecked'
  );
  
  return checkboxUnchecked;
}

/**
 * Test that verifies Right and Wrong checkboxes are disabled when Start Quiz is unchecked
 * Returns true if the test passes, false otherwise
 */
function testRightWrongCheckboxesDisabledWhenQuizNotStarted() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const quizSheet = ss.getSheetByName('quiz');
  const datastoreSheet = ss.getSheetByName('datastore');
  
  // Get valid category from datastore
  const data = datastoreSheet.getRange('B2:B' + datastoreSheet.getLastRow()).getValues();
  const validCategories = [...new Set(data.flat().filter(Boolean))];
  const validCategory = validCategories.length > 0 ? validCategories[0] : 'Rishis';
  
  // Setup: Ensure Start Quiz is unchecked and select a valid category
  quizSheet.getRange('B2').setValue(false); // Uncheck Start Quiz
  quizSheet.getRange('A1').setValue(validCategory); // Set a valid category
  
  // Create a mock edit event for unchecking Start Quiz
  const mockEvent = {
    source: ss,
    range: quizSheet.getRange('B2'),
    value: false,
    oldValue: true
  };
  
  // Call the handler function with our mock event
  handleCheckboxEdit(mockEvent);
  
  // Check if Right and Wrong checkboxes are protected (disabled)
  const rightProtected = isRangeProtected(quizSheet, 'B5');
  const wrongProtected = isRangeProtected(quizSheet, 'B6');
  
  // Also check if the checkboxes are unchecked
  const rightUnchecked = quizSheet.getRange('B5').getValue() === false;
  const wrongUnchecked = quizSheet.getRange('B6').getValue() === false;
  
  const testPassed = rightProtected && wrongProtected && rightUnchecked && wrongUnchecked;
  
  // Record result in the test sheet
  recordTestResult(
    'Right and Wrong checkboxes should be disabled when Start Quiz is unchecked', 
    testPassed, 
    testPassed ? 
      '✓ Right and Wrong checkboxes are properly disabled and unchecked' : 
      `✗ Right protected: ${rightProtected}, Wrong protected: ${wrongProtected}, Right unchecked: ${rightUnchecked}, Wrong unchecked: ${wrongUnchecked}`
  );
  
  return testPassed;
}

/**
 * Test that verifies Right and Wrong checkboxes are enabled when Start Quiz is checked
 * Returns true if the test passes, false otherwise
 */
function testRightWrongCheckboxesEnabledWhenQuizStarted() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const quizSheet = ss.getSheetByName('quiz');
  const datastoreSheet = ss.getSheetByName('datastore');
  
  // Get valid category from datastore
  const data = datastoreSheet.getRange('B2:B' + datastoreSheet.getLastRow()).getValues();
  const validCategories = [...new Set(data.flat().filter(Boolean))];
  const testCategory = validCategories.length > 0 ? validCategories[0] : 'Rishis';
  
  // Setup: Set category and check Start Quiz
  quizSheet.getRange('A1').setValue(testCategory);
  quizSheet.getRange('B2').setValue(true); // Check Start Quiz
  
  // Create a mock edit event for checking Start Quiz
  const mockEvent = {
    source: ss,
    range: quizSheet.getRange('B2'),
    value: true,
    oldValue: false
  };
  
  // Call the handler function with our mock event
  handleCheckboxEdit(mockEvent);
  
  // Check if Right and Wrong checkboxes are NOT protected (enabled)
  const rightNotProtected = !isRangeProtected(quizSheet, 'B5');
  const wrongNotProtected = !isRangeProtected(quizSheet, 'B6');
  
  // Also check if a question was loaded
  const questionLoaded = quizSheet.getRange('A4').getValue() !== '';
  
  const testPassed = rightNotProtected && wrongNotProtected && questionLoaded;
  
  // Record result in the test sheet
  recordTestResult(
    'Right and Wrong checkboxes should be enabled when Start Quiz is checked', 
    testPassed, 
    testPassed ? 
      '✓ Right and Wrong checkboxes are properly enabled and question loaded' : 
      `✗ Right enabled: ${rightNotProtected}, Wrong enabled: ${wrongNotProtected}, Question loaded: ${questionLoaded}`
  );
  
  return testPassed;
}

/**
 * Test that verifies Right and Wrong checkboxes auto-uncheck when clicked while Start Quiz is unchecked
 * Returns true if the test passes, false otherwise
 */
function testRightWrongCheckboxesAutoUncheckWhenQuizNotStarted() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const quizSheet = ss.getSheetByName('quiz');
  const datastoreSheet = ss.getSheetByName('datastore');
  
  // Get valid category from datastore
  const data = datastoreSheet.getRange('B2:B' + datastoreSheet.getLastRow()).getValues();
  const validCategories = [...new Set(data.flat().filter(Boolean))];
  const validCategory = validCategories.length > 0 ? validCategories[0] : 'Rishis';
  
  // Setup: Ensure Start Quiz is unchecked
  quizSheet.getRange('B2').setValue(false); // Uncheck Start Quiz
  quizSheet.getRange('A1').setValue(validCategory); // Set a valid category
  
  // Try to check the Right checkbox when Start Quiz is unchecked
  const mockEventRight = {
    source: ss,
    range: quizSheet.getRange('B5'),
    value: true,
    oldValue: false
  };
  
  // Call the handler function
  handleCheckboxEdit(mockEventRight);
  
  // Check if Right checkbox was auto-unchecked
  const rightAutoUnchecked = quizSheet.getRange('B5').getValue() === false;
  
  // Try to check the Wrong checkbox when Start Quiz is unchecked
  const mockEventWrong = {
    source: ss,
    range: quizSheet.getRange('B6'),
    value: true,
    oldValue: false
  };
  
  // Call the handler function
  handleCheckboxEdit(mockEventWrong);
  
  // Check if Wrong checkbox was auto-unchecked
  const wrongAutoUnchecked = quizSheet.getRange('B6').getValue() === false;
  
  const testPassed = rightAutoUnchecked && wrongAutoUnchecked;
  
  // Record result in the test sheet
  recordTestResult(
    'Right and Wrong checkboxes should auto-uncheck when clicked while Start Quiz is unchecked', 
    testPassed, 
    testPassed ? 
      '✓ Right and Wrong checkboxes properly auto-unchecked' : 
      `✗ Right auto-unchecked: ${rightAutoUnchecked}, Wrong auto-unchecked: ${wrongAutoUnchecked}`
  );
  
  return testPassed;
}

/**
 * Test that verifies quiz completes after 5 questions and shows completion message
 * Returns true if the test passes, false otherwise
 */
function testQuizCompletesAfter5Questions() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const quizSheet = ss.getSheetByName('quiz');
  const datastoreSheet = ss.getSheetByName('datastore');
  
  // Get valid category from datastore
  const data = datastoreSheet.getRange('B2:B' + datastoreSheet.getLastRow()).getValues();
  const validCategories = [...new Set(data.flat().filter(Boolean))];
  const testCategory = validCategories.length > 0 ? validCategories[0] : 'Rishis';
  
  // Reset the quiz state first
  quizSheet.getRange('A1').setValue(testCategory);
  quizSheet.getRange('B2').setValue(false);
  quizSheet.getRange('A4').setValue('');
  quizSheet.getRange('C1').setValue(0);
  quizSheet.getRange('B5').setValue(false);
  quizSheet.getRange('B6').setValue(false);
  
  // Start the quiz manually by directly calling the logic
  quizSheet.getRange('B2').setValue(true);
  
  // Manually trigger the start quiz logic
  const filtered = datastoreSheet.getDataRange().getValues().filter((row, index) => index !== 0 && row[1] === testCategory);
  if (filtered.length > 0) {
    const random = filtered[Math.floor(Math.random() * filtered.length)];
    quizSheet.getRange('A4').setValue(random[2]);
    quizSheet.getRange('C1').setValue(1);
    toggleRightWrongCheckboxes(quizSheet, true);
  }
  
  let debugInfo = [];
  debugInfo.push(`Initial state: Counter = ${quizSheet.getRange('C1').getValue()}, Question = "${quizSheet.getRange('A4').getValue()}"`);
  
  // Manually simulate answering 4 questions (we already have question 1 loaded)
  for (let i = 0; i < 4; i++) {
    const currentCount = getQuestionCounter(quizSheet);
    debugInfo.push(`Before question ${i + 2}: Counter = ${currentCount}`);
    
    // Call showNextQuestion directly
    showNextQuestion(quizSheet, ss);
    
    const newCount = getQuestionCounter(quizSheet);
    const questionText = quizSheet.getRange('A4').getValue();
    debugInfo.push(`After question ${i + 2}: Counter = ${newCount}, Question = "${questionText}"`);
  }
  
  // Now we should be at question 5. Call showNextQuestion one more time
  const beforeFinalCount = getQuestionCounter(quizSheet);
  debugInfo.push(`Before final call: Counter = ${beforeFinalCount}`);
  
  showNextQuestion(quizSheet, ss);
  
  // Check the final state
  const finalQuestionText = quizSheet.getRange('A4').getValue();
  const finalCounter = quizSheet.getRange('C1').getValue();
  const startQuizUnchecked = quizSheet.getRange('B2').getValue() === false;
  const rightCheckboxDisabled = isRangeProtected(quizSheet, 'B5');
  const wrongCheckboxDisabled = isRangeProtected(quizSheet, 'B6');
  
  debugInfo.push(`Final state: Counter = ${finalCounter}, Question = "${finalQuestionText}", Start Quiz unchecked = ${startQuizUnchecked}`);
  
  const isCompletionMessage = finalQuestionText && finalQuestionText.toString().includes('Quiz Complete');
  const questionCounterReset = finalCounter === 0;
  
  const testPassed = isCompletionMessage && startQuizUnchecked && rightCheckboxDisabled && wrongCheckboxDisabled && questionCounterReset;
  
  // Record result in the test sheet with detailed debugging info
  recordTestResult(
    'Quiz should complete after 5 questions with completion message', 
    testPassed, 
    testPassed ? 
      '✓ Quiz properly completed after 5 questions with all expected behaviors' : 
      `✗ Debug info: ${debugInfo.join(' | ')} | Final checks: Completion message found: ${isCompletionMessage}, Counter reset: ${questionCounterReset}, Start Quiz unchecked: ${startQuizUnchecked}, Right disabled: ${rightCheckboxDisabled}, Wrong disabled: ${wrongCheckboxDisabled}`
  );
  
  return testPassed;
}

/**
 * Test that verifies question counter resets when category is changed
 * Returns true if the test passes, false otherwise
 */
function testQuestionCounterResetsOnCategoryChange() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const quizSheet = ss.getSheetByName('quiz');
  const datastoreSheet = ss.getSheetByName('datastore');
  
  // Get valid categories from datastore
  const data = datastoreSheet.getRange('B2:B' + datastoreSheet.getLastRow()).getValues();
  const validCategories = [...new Set(data.flat().filter(Boolean))];
  const category1 = validCategories.length > 0 ? validCategories[0] : 'Rishis';
  const category2 = validCategories.length > 1 ? validCategories[1] : validCategories[0];
  
  // Setup: Start a quiz and answer some questions
  quizSheet.getRange('A1').setValue(category1);
  quizSheet.getRange('B2').setValue(true);
  
  // Start the quiz
  const startQuizEvent = {
    source: ss,
    range: quizSheet.getRange('B2'),
    value: true,
    oldValue: false
  };
  handleCheckboxEdit(startQuizEvent);
  
  // Answer 2 questions to set counter
  for (let i = 0; i < 2; i++) {
    const rightClickEvent = {
      source: ss,
      range: quizSheet.getRange('B5'),
      value: true,
      oldValue: false
    };
    handleCheckboxEdit(rightClickEvent);
  }
  
  // Verify counter is at 3 (1 initial + 2 additional)
  const counterBefore = quizSheet.getRange('C1').getValue();
  
  // Change category
  const categoryChangeEvent = {
    source: ss,
    range: quizSheet.getRange('A1'),
    value: category2,
    oldValue: category1
  };
  handleCheckboxEdit(categoryChangeEvent);
  
  // Check if counter was reset
  const counterAfter = quizSheet.getRange('C1').getValue();
  const counterReset = counterAfter === 0;
  
  const testPassed = counterBefore > 0 && counterReset;
  
  // Record result in the test sheet
  recordTestResult(
    'Question counter should reset when category is changed', 
    testPassed, 
    testPassed ? 
      '✓ Question counter properly reset when category changed' : 
      `✗ Counter before: ${counterBefore}, Counter after: ${counterAfter}, Reset: ${counterReset}`
  );
  
  return testPassed;
}

/**
 * Run all unit tests and record results in the test results sheet
 */
function runAllTests() {
  // Clear previous results before running new tests
  clearTestResults();
  
  // Test 1: Category change clears question cell
  testCategoryClearsQuestionCell();
  
  // Test 2: Category change unchecks checkbox
  testCategoryClearsCheckbox();
  
  // Test 3: Right and Wrong checkboxes disabled when quiz not started
  testRightWrongCheckboxesDisabledWhenQuizNotStarted();
  
  // Test 4: Right and Wrong checkboxes enabled when quiz started
  testRightWrongCheckboxesEnabledWhenQuizStarted();
  
  // Test 5: Right and Wrong checkboxes auto-uncheck when clicked while quiz not started
  testRightWrongCheckboxesAutoUncheckWhenQuizNotStarted();
  
  // Test 6: Quiz completes after 5 questions
  testQuizCompletesAfter5Questions();
  
  // Test 7: Question counter resets on category change
  testQuestionCounterResetsOnCategoryChange();
  
  // Add more tests here as needed
}

/**
 * Run a specific test by name
 * Can be used to run individual tests without clearing all results
 */
function runTest(testName) {
  switch(testName) {
    case 'clearQuestionCell':
      testCategoryClearsQuestionCell();
      break;
    case 'clearCheckbox':
      testCategoryClearsCheckbox();
      break;
    case 'disableRightWrongWhenQuizNotStarted':
      testRightWrongCheckboxesDisabledWhenQuizNotStarted();
      break;
    case 'enableRightWrongWhenQuizStarted':
      testRightWrongCheckboxesEnabledWhenQuizStarted();
      break;
    case 'autoUncheckRightWrongWhenQuizNotStarted':
      testRightWrongCheckboxesAutoUncheckWhenQuizNotStarted();
      break;
    case 'quizCompletesAfter5Questions':
      testQuizCompletesAfter5Questions();
      break;
    case 'questionCounterResetsOnCategoryChange':
      testQuestionCounterResetsOnCategoryChange();
      break;
    default:
      throw new Error(`Test "${testName}" not found`);
  }
}