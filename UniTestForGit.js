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

// Test that verifies cell A4 is cleared when category in A1 is changed
function testCategoryClearsQuestionCell() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const quizSheet = ss.getSheetByName('quiz');

  quizSheet.getRange('A4').setValue('Test question that should be cleared');
  quizSheet.getRange('B2').setValue(true);

  const mockEvent = {
    source: ss,
    range: quizSheet.getRange('A1'),
    value: 'Some Category',
    oldValue: 'Previous Category'
  };

  handleCheckboxEdit(mockEvent);

  const a4Value = quizSheet.getRange('A4').getValue();
  const a4Cleared = a4Value === '';

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
function runAllTestsGit() {
  clearTestResults();

  testCategoryClearsQuestionCell
}