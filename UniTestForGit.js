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
 * Run all unit tests and record results in the test results sheet
 */
function runAllTests() {
  // Clear previous results before running new tests
  clearTestResults();
  // Test 1: Category change clears question cell
  testCategoryClearsQuestionCell();
}

