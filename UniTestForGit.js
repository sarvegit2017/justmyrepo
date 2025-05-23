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

//Test that verifies quiz ends early if no more unique questions are available Returns true if the test passes, false otherwise
function testQuizEndsEarlyWhenNoMoreUniqueQuestions() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const quizSheet = ss.getSheetByName('quiz');
  const datastoreSheet = ss.getSheetByName('datastore');

  const data = datastoreSheet.getDataRange().getValues();
  const categoryQuestionCounts = {};

  data.slice(1).forEach(row => {
    if (row[1]) {
      categoryQuestionCounts[row[1]] = (categoryQuestionCounts[row[1]] || 0) + 1;
    }
  });

  let testCategory = null;
  let questionCount = 0;
  for (const [category, count] of Object.entries(categoryQuestionCounts)) {
    if (count < 5 && count > 1) {
      testCategory = category;
      questionCount = count;
      break;
    }
  }

  if (!testCategory) {

    const validCategories = [...new Set(data.slice(1).map(row => row[1]).filter(Boolean))];
    testCategory = validCategories.length > 0 ? validCategories[0] : 'Rishis';


    const allQuestionsForCategory = data.filter((row, index) => index !== 0 && row[1] === testCategory);

    const preUsedQuestions = allQuestionsForCategory.slice(0, -2).map(row => row[2]);
    quizSheet.getRange('D1').setValue(JSON.stringify(preUsedQuestions));
    questionCount = 2;
  }

  quizSheet.getRange('A1').setValue(testCategory);
  quizSheet.getRange('B2').setValue(false);
  quizSheet.getRange('A4').setValue('');
  quizSheet.getRange('C1').setValue(0);
  quizSheet.getRange('B5').setValue(false);
  quizSheet.getRange('B6').setValue(false);

  const startQuizEvent = {
    source: ss,
    range: quizSheet.getRange('B2'),
    value: true,
    oldValue: false
  };
  handleCheckboxEdit(startQuizEvent);

  let questionsAnswered = 1; // First question is loaded when quiz starts
  let debugInfo = [`Started with ${questionCount} available questions`];

  for (let i = 1; i < 5; i++) {
    const rightClickEvent = {
      source: ss,
      range: quizSheet.getRange('B5'),
      value: true,
      oldValue: false
    };
    handleCheckboxEdit(rightClickEvent);

    const currentQuestion = quizSheet.getRange('A4').getValue();

    if (currentQuestion && currentQuestion.toString().includes('Quiz Complete')) {
      debugInfo.push(`Quiz ended at question ${questionsAnswered} with completion message`);
      break;
    } else if (currentQuestion && currentQuestion !== '') {
      questionsAnswered++;
      debugInfo.push(`Question ${questionsAnswered} shown`);
    }
  }

  const finalQuestionText = quizSheet.getRange('A4').getValue();
  const showsEarlyCompletionMessage = finalQuestionText && (
    finalQuestionText.toString().includes('Quiz Complete') ||
    finalQuestionText.toString().includes('No more unique questions')
  );
  const endedEarly = questionsAnswered < 5;
  const startQuizUnchecked = quizSheet.getRange('B2').getValue() === false;

  const testPassed = endedEarly && showsEarlyCompletionMessage && startQuizUnchecked;

  recordTestResult(
    'Quiz should end early when no more unique questions are available',
    testPassed,
    testPassed ?
      `✓ Quiz properly ended early after ${questionsAnswered} questions` :
      `✗ Early end failed. Questions answered: ${questionsAnswered}, Shows completion: ${showsEarlyCompletionMessage}, Start Quiz unchecked: ${startQuizUnchecked}. Debug: ${debugInfo.join(' | ')}`
  );

  return testPassed;
}

/**
 * Run all unit tests and record results in the test results sheet
 */
function runAllTestsGit() {
  clearTestResults();

  testQuizEndsEarlyWhenNoMoreUniqueQuestions();
}