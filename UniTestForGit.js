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




// Test that verifies right answer count increments when Right checkbox is checked
function testRightAnswerIncrement() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const quizSheet = ss.getSheetByName('quiz');

  // Setup initial state
  quizSheet.getRange('A1').setValue('Test Category');
  quizSheet.getRange('B2').setValue(true); // Start quiz
  quizSheet.getRange('A4').setValue('Test Question');
  quizSheet.getRange('B9').setValue(0); // Reset right count
  quizSheet.getRange('B10').setValue(0); // Reset wrong count

  // Mock checking Right checkbox
  const mockEvent = {
    source: ss,
    range: quizSheet.getRange('B5'),
    value: true,
    oldValue: false
  };

  const initialRightCount = quizSheet.getRange('B9').getValue();
  handleCheckboxEdit(mockEvent);
  const finalRightCount = quizSheet.getRange('B9').getValue();

  const incrementedCorrectly = (finalRightCount === initialRightCount + 1);

  recordTestResult(
    'testRightAnswerIncrement',
    'Right answer count should increment when Right checkbox is checked',
    incrementedCorrectly,
    incrementedCorrectly ?
      `✓ Right count incremented from ${initialRightCount} to ${finalRightCount}` :
      `✗ Right count did not increment correctly: ${initialRightCount} to ${finalRightCount}`
  );

  return incrementedCorrectly;
}

// Test that verifies wrong answer count increments when Wrong checkbox is checked
function testWrongAnswerIncrement() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const quizSheet = ss.getSheetByName('quiz');

  // Setup initial state
  quizSheet.getRange('A1').setValue('Test Category');
  quizSheet.getRange('B2').setValue(true); // Start quiz
  quizSheet.getRange('A4').setValue('Test Question');
  quizSheet.getRange('B9').setValue(0); // Reset right count
  quizSheet.getRange('B10').setValue(0); // Reset wrong count

  // Mock checking Wrong checkbox
  const mockEvent = {
    source: ss,
    range: quizSheet.getRange('B6'),
    value: true,
    oldValue: false
  };

  const initialWrongCount = quizSheet.getRange('B10').getValue();
  handleCheckboxEdit(mockEvent);
  const finalWrongCount = quizSheet.getRange('B10').getValue();

  const incrementedCorrectly = (finalWrongCount === initialWrongCount + 1);

  recordTestResult(
    'testWrongAnswerIncrement',
    'Wrong answer count should increment when Wrong checkbox is checked',
    incrementedCorrectly,
    incrementedCorrectly ?
      `✓ Wrong count incremented from ${initialWrongCount} to ${finalWrongCount}` :
      `✗ Wrong count did not increment correctly: ${initialWrongCount} to ${finalWrongCount}`
  );

  return incrementedCorrectly;
}

// Test multiple right answers increment correctly
function testMultipleRightAnswers() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const quizSheet = ss.getSheetByName('quiz');

  // Setup initial state
  quizSheet.getRange('A1').setValue('Test Category');
  quizSheet.getRange('B2').setValue(true); // Start quiz
  quizSheet.getRange('A4').setValue('Test Question');
  quizSheet.getRange('B9').setValue(0); // Reset right count

  let allIncrementsCorrect = true;
  let details = '';

  // Check Right checkbox 3 times
  for (let i = 1; i <= 3; i++) {
    const mockEvent = {
      source: ss,
      range: quizSheet.getRange('B5'),
      value: true,
      oldValue: false
    };

    const beforeCount = quizSheet.getRange('B9').getValue();
    handleCheckboxEdit(mockEvent);
    const afterCount = quizSheet.getRange('B9').getValue();

    if (afterCount !== i) {
      allIncrementsCorrect = false;
      details += `✗ After ${i} clicks, expected ${i} but got ${afterCount}. `;
    } else {
      details += `✓ Click ${i}: ${beforeCount} → ${afterCount}. `;
    }

    // Reset checkbox for next iteration
    quizSheet.getRange('B5').setValue(false);
  }

  recordTestResult(
    'testMultipleRightAnswers',
    'Multiple right answer selections should increment correctly',
    allIncrementsCorrect,
    details
  );

  return allIncrementsCorrect;
}

// Test multiple wrong answers increment correctly
function testMultipleWrongAnswers() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const quizSheet = ss.getSheetByName('quiz');

  // Setup initial state
  quizSheet.getRange('A1').setValue('Test Category');
  quizSheet.getRange('B2').setValue(true); // Start quiz
  quizSheet.getRange('A4').setValue('Test Question');
  quizSheet.getRange('B10').setValue(0); // Reset wrong count

  let allIncrementsCorrect = true;
  let details = '';

  // Check Wrong checkbox 3 times
  for (let i = 1; i <= 3; i++) {
    const mockEvent = {
      source: ss,
      range: quizSheet.getRange('B6'),
      value: true,
      oldValue: false
    };

    const beforeCount = quizSheet.getRange('B10').getValue();
    handleCheckboxEdit(mockEvent);
    const afterCount = quizSheet.getRange('B10').getValue();

    if (afterCount !== i) {
      allIncrementsCorrect = false;
      details += `✗ After ${i} clicks, expected ${i} but got ${afterCount}. `;
    } else {
      details += `✓ Click ${i}: ${beforeCount} → ${afterCount}. `;
    }

    // Reset checkbox for next iteration
    quizSheet.getRange('B6').setValue(false);
  }

  recordTestResult(
    'testMultipleWrongAnswers',
    'Multiple wrong answer selections should increment correctly',
    allIncrementsCorrect,
    details
  );

  return allIncrementsCorrect;
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
    const mockEvent = {
      source: ss,
      range: quizSheet.getRange(step.checkbox),
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



// Test that scores are preserved during quiz progress
function testScorePreservationDuringQuizProgress() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const quizSheet = ss.getSheetByName('quiz');

  // Setup initial state
  quizSheet.getRange('A1').setValue('Test Category');
  quizSheet.getRange('B2').setValue(true); // Start quiz
  quizSheet.getRange('A4').setValue('Test Question 1');
  quizSheet.getRange('B9').setValue(2); // Set some initial scores
  quizSheet.getRange('B10').setValue(1);

  // Simulate answering correctly (this would typically trigger next question)
  const mockEvent = {
    source: ss,
    range: quizSheet.getRange('B5'),
    value: true,
    oldValue: false
  };

  const rightCountBefore = quizSheet.getRange('B9').getValue();
  const wrongCountBefore = quizSheet.getRange('B10').getValue();

  handleCheckboxEdit(mockEvent);

  const rightCountAfter = quizSheet.getRange('B9').getValue();
  const wrongCountAfter = quizSheet.getRange('B10').getValue();

  // Should increment right count but preserve wrong count
  const preservedCorrectly = (rightCountAfter === rightCountBefore + 1 && wrongCountAfter === wrongCountBefore);

  recordTestResult(
    'testScorePreservationDuringQuizProgress',
    'Scores should be preserved and increment correctly during quiz progress',
    preservedCorrectly,
    preservedCorrectly ?
      `✓ Scores preserved and incremented: Right ${rightCountBefore}→${rightCountAfter}, Wrong ${wrongCountBefore}→${wrongCountAfter}` :
      `✗ Scores not handled correctly: Right ${rightCountBefore}→${rightCountAfter}, Wrong ${wrongCountBefore}→${wrongCountAfter}`
  );

  return preservedCorrectly;
}

/**
 * Run all unit tests and record results in the test results sheet
 */
function runAllTestsGit() {
  clearTestResults();


  // Score Increment Tests
  testRightAnswerIncrement();
  testWrongAnswerIncrement();
  testMultipleRightAnswers();
  testMultipleWrongAnswers();
  testMixedAnswers();

  testScorePreservationDuringQuizProgress();
}