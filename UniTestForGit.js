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
 * Helper function to create a proper mock edit event
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} spreadsheet - The spreadsheet
 * @param {GoogleAppsScript.Spreadsheet.Range} range - The range being edited
 * @param {*} newValue - The new value
 * @param {*} oldValue - The old value
 * @returns {Object} Mock event object
 */
function createMockEditEvent(spreadsheet, range, newValue, oldValue) {
  // Set the actual value in the range first
  range.setValue(newValue);
  
  return {
    source: spreadsheet,
    range: range,
    value: newValue,
    oldValue: oldValue
  };
}

/**
 * Test that verifies questions are not repeated during a quiz session
 * Returns true if the test passes, false otherwise
 */
function testQuestionsNotRepeated() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const quizSheet = ss.getSheetByName('quiz');
  const datastoreSheet = ss.getSheetByName('datastore');
  
  try {
    // Get valid category from datastore with multiple questions
    const data = datastoreSheet.getDataRange().getValues();
    const categoryQuestions = {};
    
    // Count questions per category
    data.slice(1).forEach(row => {
      if (row[1] && row[2]) { // Category and Question exist
        if (!categoryQuestions[row[1]]) {
          categoryQuestions[row[1]] = [];
        }
        categoryQuestions[row[1]].push(row[2]);
      }
    });
    
    // Find a category with at least 3 questions for testing
    let testCategory = null;
    let testCategoryQuestions = [];
    
    for (const [category, questions] of Object.entries(categoryQuestions)) {
      if (questions.length >= 3) {
        testCategory = category;
        testCategoryQuestions = questions;
        break;
      }
    }
    
    if (!testCategory) {
      recordTestResult(
        'Questions should not be repeated during a quiz session', 
        false, 
        '✗ No category found with at least 3 questions for testing'
      );
      return false;
    }
    
    // Reset the quiz state completely
    quizSheet.getRange('A1').setValue('');
    quizSheet.getRange('B2').setValue(false);
    quizSheet.getRange('A4').setValue('');
    quizSheet.getRange('C1').setValue('');
    quizSheet.getRange('D1').setValue('');
    quizSheet.getRange('B5').setValue(false);
    quizSheet.getRange('B6').setValue(false);
    
    // Clear any existing protections
    const protections = quizSheet.getProtections(SpreadsheetApp.ProtectionType.RANGE);
    protections.forEach(protection => {
      const range = protection.getRange();
      if (range.getA1Notation() === 'B5' || range.getA1Notation() === 'B6') {
        protection.remove();
      }
    });
    
    // Step 1: Set category
    const categoryEvent = createMockEditEvent(ss, quizSheet.getRange('A1'), testCategory, '');
    handleCheckboxEdit(categoryEvent);
    
    // Step 2: Start the quiz
    const startQuizEvent = createMockEditEvent(ss, quizSheet.getRange('B2'), true, false);
    handleCheckboxEdit(startQuizEvent);
    
    const usedQuestions = [];
    let hasRepeats = false;
    let debugInfo = [];
    
    // Collect the first question
    let firstQuestion = quizSheet.getRange('A4').getValue();
    if (firstQuestion && firstQuestion !== '' && !firstQuestion.toString().includes('No questions')) {
      usedQuestions.push(firstQuestion.toString());
      debugInfo.push(`Q1: "${firstQuestion}"`);
    } else {
      recordTestResult(
        'Questions should not be repeated during a quiz session', 
        false, 
        '✗ No first question generated. First question: ' + firstQuestion
      );
      return false;
    }
    
    // Simulate answering questions and check for repetition (test up to 4 more questions)
    for (let i = 1; i < 5; i++) {
      // Answer the current question (use Right checkbox)
      const rightClickEvent = createMockEditEvent(ss, quizSheet.getRange('B5'), true, false);
      handleCheckboxEdit(rightClickEvent);
      
      // Small delay to ensure the question is updated
      Utilities.sleep(100);
      
      const currentQuestion = quizSheet.getRange('A4').getValue();
      
      // Check if this is a completion message
      if (currentQuestion && currentQuestion.toString().includes('Quiz Complete')) {
        debugInfo.push(`Quiz completed at question ${i + 1}`);
        break;
      }
      
      // Check if this is an error message
      if (currentQuestion && currentQuestion.toString().includes('No questions')) {
        debugInfo.push(`No more questions at question ${i + 1}`);
        break;
      }
      
      if (currentQuestion && currentQuestion !== '') {
        const questionStr = currentQuestion.toString();
        debugInfo.push(`Q${i + 1}: "${questionStr}"`);
        
        // Check if this question was already used
        if (usedQuestions.includes(questionStr)) {
          hasRepeats = true;
          debugInfo.push(`REPEAT DETECTED: "${questionStr}" was already used`);
          break;
        } else {
          usedQuestions.push(questionStr);
        }
      } else {
        debugInfo.push(`Q${i + 1}: Empty question`);
        break;
      }
    }
    
    const testPassed = !hasRepeats && usedQuestions.length >= 2;
    
    // Record result in the test sheet
    recordTestResult(
      'Questions should not be repeated during a quiz session', 
      testPassed, 
      testPassed ? 
        `✓ No repeated questions found. Used ${usedQuestions.length} unique questions from category "${testCategory}"` : 
        `✗ ${hasRepeats ? 'Repeated questions detected' : 'Not enough unique questions generated'}. Category: "${testCategory}", Questions used: ${usedQuestions.length}. Debug: ${debugInfo.join(' | ')}`
    );
    
    return testPassed;
    
  } catch (error) {
    recordTestResult(
      'Questions should not be repeated during a quiz session', 
      false, 
      `✗ Test failed with error: ${error.toString()}`
    );
    return false;
  }
}

/**
 * Run all unit tests and record results in the test results sheet
 */
function runAllTestsGit() {
  // Clear previous results before running new tests
  clearTestResults();
  // Test 1: Questions should not be repeated during a quiz session
  testQuestionsNotRepeated();
}