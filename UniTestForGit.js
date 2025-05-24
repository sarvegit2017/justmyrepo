function setupTestResultsSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let testSheet = ss.getSheetByName('testResults');

  if (!testSheet) {
    testSheet = ss.insertSheet('testResults');
    testSheet.getRange('A1:D1').setValues([['Unit Test Name', 'Test Name', 'Result', 'Details']]);
    testSheet.getRange('A1:D1').setFontWeight('bold');
    testSheet.setFrozenRows(1);
    testSheet.autoResizeColumns(1, 4);
    testSheet.setColumnWidth(4, 300);
  }
  return testSheet;
}

function clearTestResults() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let testSheet = ss.getSheetByName('testResults');

  if (testSheet) {
    const lastRow = testSheet.getLastRow();
    if (lastRow > 1) {
      testSheet.getRange(2, 1, lastRow - 1, 4).clearContent();
      testSheet.getRange(2, 1, lastRow - 1, 4).clearFormat();
    }
  }
}

function recordTestResult(unitTestName, testName, passed, details) {
  const testSheet = setupTestResultsSheet();
  const lastRow = Math.max(1, testSheet.getLastRow());

  testSheet.getRange(lastRow + 1, 1, 1, 4).setValues([
    [
      unitTestName,
      testName,
      passed ? 'PASSED' : 'FAILED',
      details
    ]
  ]);

  const resultCell = testSheet.getRange(lastRow + 1, 3);
  resultCell.setBackground(passed ? '#b7e1cd' : '#f4c7c3');

  testSheet.autoResizeColumns(1, 4);
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





// Test mixed right and wrong answers
function testMixedAnswers() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const quizSheet = ss.getSheetByName('quiz');

  // Setup initial state
  quizSheet.getRange('A1').setValue('Test Category');
  quizSheet.getRange('B2').setValue(true); // Start quiz
  quizSheet.getRange('A4').setValue('Test Question');
  quizSheet.getRange('B9').setValue(0); // Reset right count
  quizSheet.getRange('B10').setValue(0); // Reset wrong count

  // Enable the checkboxes by removing protection (simulate quiz started state)
  toggleRightWrongCheckboxes(quizSheet, true);

  // Sequence: Right, Wrong, Right, Wrong, Right
  const sequence = [
    { checkbox: 'B5', expected: { right: 1, wrong: 0 } },
    { checkbox: 'B6', expected: { right: 1, wrong: 1 } },
    { checkbox: 'B5', expected: { right: 2, wrong: 1 } },
    { checkbox: 'B6', expected: { right: 2, wrong: 2 } },
    { checkbox: 'B5', expected: { right: 3, wrong: 2 } }
  ];

  let allCorrect = true;
  let details = '';

  sequence.forEach((step, index) => {
    const checkboxRange = quizSheet.getRange(step.checkbox);
    checkboxRange.setValue(true); // Set the checkbox value first
    
    const mockEvent = {
      source: ss,
      range: checkboxRange,
      value: true,
      oldValue: false
    };

    handleCheckboxEdit(mockEvent);

    const rightCount = quizSheet.getRange('B9').getValue();
    const wrongCount = quizSheet.getRange('B10').getValue();

    if (rightCount !== step.expected.right || wrongCount !== step.expected.wrong) {
      allCorrect = false;
      details += `✗ Step ${index + 1}: Expected R=${step.expected.right}, W=${step.expected.wrong}, Got R=${rightCount}, W=${wrongCount}. `;
    } else {
      details += `✓ Step ${index + 1}: R=${rightCount}, W=${wrongCount}. `;
    }

    // Reset checkboxes for next iteration
    quizSheet.getRange('B5').setValue(false);
    quizSheet.getRange('B6').setValue(false);
  });

  recordTestResult(
    'testMixedAnswers',
    'Mixed right and wrong answers should increment both counters correctly',
    allCorrect,
    details
  );

  return allCorrect;
}





/**
 * Run all unit tests and record results in the test results sheet
 */
function runAllTestsGit() {
  clearTestResults();


  // Score Increment Tests

  testMixedAnswers();

}