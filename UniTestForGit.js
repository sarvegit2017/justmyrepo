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



// Test that verifies used questions list is properly tracked in cell D1
function testUsedQuestionsTracking() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const quizSheet = ss.getSheetByName('quiz');
  const datastoreSheet = ss.getSheetByName('datastore');

  const data = datastoreSheet.getDataRange().getValues();
  const validCategories = [...new Set(data.slice(1).map(row => row[1]).filter(Boolean))];
  const testCategory = validCategories.length > 0 ? validCategories[0] : 'Rishis';

  quizSheet.getRange('A1').setValue(testCategory);
  quizSheet.getRange('B2').setValue(false);
  quizSheet.getRange('A4').setValue('');
  quizSheet.getRange('C1').setValue(0);
  quizSheet.getRange('D1').setValue('');
  quizSheet.getRange('B5').setValue(false);
  quizSheet.getRange('B6').setValue(false);

  const initialUsedQuestions = getUsedQuestions(quizSheet);
  const initiallyEmpty = Array.isArray(initialUsedQuestions) && initialUsedQuestions.length === 0;

  const startQuizEvent = {
    source: ss,
    range: quizSheet.getRange('B2'),
    value: true,
    oldValue: false
  };
  handleCheckboxEdit(startQuizEvent);

  const firstQuestion = quizSheet.getRange('A4').getValue();
  const usedQuestionsAfterFirst = getUsedQuestions(quizSheet);
  const firstQuestionTracked = usedQuestionsAfterFirst.includes(firstQuestion);

  const rightClickEvent = {
    source: ss,
    range: quizSheet.getRange('B5'),
    value: true,
    oldValue: false
  };
  handleCheckboxEdit(rightClickEvent);

  const secondQuestion = quizSheet.getRange('A4').getValue();
  const usedQuestionsAfterSecond = getUsedQuestions(quizSheet);
  const secondQuestionTracked = !secondQuestion.includes('Quiz Complete') && usedQuestionsAfterSecond.includes(secondQuestion);
  const bothQuestionsTracked = usedQuestionsAfterSecond.length >= 2 || secondQuestion.includes('Quiz Complete');

  const testPassed = initiallyEmpty && firstQuestionTracked && (secondQuestionTracked || secondQuestion.includes('Quiz Complete'));

  recordTestResult(
    'Used questions should be properly tracked in cell D1',
    testPassed,
    testPassed ?
      `✓ Used questions properly tracked. Initial: ${initialUsedQuestions.length}, After first: ${usedQuestionsAfterFirst.length}, After second: ${usedQuestionsAfterSecond.length}` :
      `✗ Tracking failed. Initial empty: ${initiallyEmpty}, First tracked: ${firstQuestionTracked}, Second tracked: ${secondQuestionTracked}, Questions: ["${firstQuestion}", "${secondQuestion}"]`
  );

  return testPassed;
}//ok

/**
 * Run all unit tests and record results in the test results sheet
 */
function runAllTestsGit() {
  clearTestResults();

  testUsedQuestionsTracking();
}

