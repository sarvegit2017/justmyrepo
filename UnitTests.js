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

// Test that verifies checkbox in B2 is unchecked when category in A1 is changed
function testCategoryClearsCheckbox() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const quizSheet = ss.getSheetByName('quiz');

  quizSheet.getRange('B2').setValue(true);

  const mockEvent = {
    source: ss,
    range: quizSheet.getRange('A1'),
    value: 'Some Category',
    oldValue: 'Previous Category'
  };

  handleCheckboxEdit(mockEvent);

  const checkboxValue = quizSheet.getRange('B2').getValue();
  const checkboxUnchecked = checkboxValue === false;

  recordTestResult(
    'When category is changed, checkbox in B2 should be cleared',
    checkboxUnchecked,
    checkboxUnchecked ?
      '✓ B2 checkbox was properly unchecked' :
      '✗ B2 checkbox was not unchecked'
  );

  return checkboxUnchecked;
}

// Test that verifies Right and Wrong checkboxes are disabled when Start Quiz is unchecked
function testRightWrongCheckboxesDisabledWhenQuizNotStarted() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const quizSheet = ss.getSheetByName('quiz');
  const datastoreSheet = ss.getSheetByName('datastore');

  const data = datastoreSheet.getRange('B2:B' + datastoreSheet.getLastRow()).getValues();
  const validCategories = [...new Set(data.flat().filter(Boolean))];
  const validCategory = validCategories.length > 0 ? validCategories[0] : 'Rishis';

  quizSheet.getRange('B2').setValue(false);
  quizSheet.getRange('A1').setValue(validCategory);

  const mockEvent = {
    source: ss,
    range: quizSheet.getRange('B2'),
    value: false,
    oldValue: true
  };

  handleCheckboxEdit(mockEvent);

  const rightProtected = isRangeProtected(quizSheet, 'B5');
  const wrongProtected = isRangeProtected(quizSheet, 'B6');

  const rightUnchecked = quizSheet.getRange('B5').getValue() === false;
  const wrongUnchecked = quizSheet.getRange('B6').getValue() === false;

  const testPassed = rightProtected && wrongProtected && rightUnchecked && wrongUnchecked;

  recordTestResult(
    'Right and Wrong checkboxes should be disabled when Start Quiz is unchecked',
    testPassed,
    testPassed ?
      '✓ Right and Wrong checkboxes are properly disabled and unchecked' :
      `✗ Right protected: ${rightProtected}, Wrong protected: ${wrongProtected}, Right unchecked: ${rightUnchecked}, Wrong unchecked: ${wrongUnchecked}`
  );

  return testPassed;
}

// Test that verifies Right and Wrong checkboxes are enabled when Start Quiz is checked
function testRightWrongCheckboxesEnabledWhenQuizStarted() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const quizSheet = ss.getSheetByName('quiz');
  const datastoreSheet = ss.getSheetByName('datastore');

  const data = datastoreSheet.getRange('B2:B' + datastoreSheet.getLastRow()).getValues();
  const validCategories = [...new Set(data.flat().filter(Boolean))];
  const testCategory = validCategories.length > 0 ? validCategories[0] : 'Rishis';

  quizSheet.getRange('A1').setValue(testCategory);
  quizSheet.getRange('B2').setValue(true);

  const mockEvent = {
    source: ss,
    range: quizSheet.getRange('B2'),
    value: true,
    oldValue: false
  };

  handleCheckboxEdit(mockEvent);

  const rightNotProtected = !isRangeProtected(quizSheet, 'B5');
  const wrongNotProtected = !isRangeProtected(quizSheet, 'B6');

  const questionLoaded = quizSheet.getRange('A4').getValue() !== '';

  const testPassed = rightNotProtected && wrongNotProtected && questionLoaded;

  recordTestResult(
    'Right and Wrong checkboxes should be enabled when Start Quiz is checked',
    testPassed,
    testPassed ?
      '✓ Right and Wrong checkboxes are properly enabled and question loaded' :
      `✗ Right enabled: ${rightNotProtected}, Wrong enabled: ${wrongNotProtected}, Question loaded: ${questionLoaded}`
  );

  return testPassed;
}

// Test that verifies Right and Wrong checkboxes auto-uncheck when clicked while Start Quiz is unchecked
function testRightWrongCheckboxesAutoUncheckWhenQuizNotStarted() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const quizSheet = ss.getSheetByName('quiz');
  const datastoreSheet = ss.getSheetByName('datastore');

  const data = datastoreSheet.getRange('B2:B' + datastoreSheet.getLastRow()).getValues();
  const validCategories = [...new Set(data.flat().filter(Boolean))];
  const validCategory = validCategories.length > 0 ? validCategories[0] : 'Rishis';

  quizSheet.getRange('B2').setValue(false);
  quizSheet.getRange('A1').setValue(validCategory);

  const mockEventRight = {
    source: ss,
    range: quizSheet.getRange('B5'),
    value: true,
    oldValue: false
  };

  handleCheckboxEdit(mockEventRight);

  const rightAutoUnchecked = quizSheet.getRange('B5').getValue() === false;

  const mockEventWrong = {
    source: ss,
    range: quizSheet.getRange('B6'),
    value: true,
    oldValue: false
  };

  handleCheckboxEdit(mockEventWrong);

  const wrongAutoUnchecked = quizSheet.getRange('B6').getValue() === false;

  const testPassed = rightAutoUnchecked && wrongAutoUnchecked;

  recordTestResult(
    'Right and Wrong checkboxes should auto-uncheck when clicked while Start Quiz is unchecked',
    testPassed,
    testPassed ?
      '✓ Right and Wrong checkboxes properly auto-unchecked' :
      `✗ Right auto-unchecked: ${rightAutoUnchecked}, Wrong auto-unchecked: ${wrongAutoUnchecked}`
  );

  return testPassed;
}

// Test that verifies quiz completes after 5 questions and shows completion message
function testQuizCompletesAfter5Questions() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const quizSheet = ss.getSheetByName('quiz');
  const datastoreSheet = ss.getSheetByName('datastore');

  const data = datastoreSheet.getRange('B2:B' + datastoreSheet.getLastRow()).getValues();
  const validCategories = [...new Set(data.flat().filter(Boolean))];
  const testCategory = validCategories.length > 0 ? validCategories[0] : 'Rishis';

  quizSheet.getRange('A1').setValue(testCategory);
  quizSheet.getRange('B2').setValue(false);
  quizSheet.getRange('A4').setValue('');
  quizSheet.getRange('C1').setValue(0);
  quizSheet.getRange('B5').setValue(false);
  quizSheet.getRange('B6').setValue(false);

  quizSheet.getRange('B2').setValue(true);

  const filtered = datastoreSheet.getDataRange().getValues().filter((row, index) => index !== 0 && row[1] === testCategory);
  if (filtered.length > 0) {
    const random = filtered[Math.floor(Math.random() * filtered.length)];
    quizSheet.getRange('A4').setValue(random[2]);
    quizSheet.getRange('C1').setValue(1);
    toggleRightWrongCheckboxes(quizSheet, true);
  }

  let debugInfo = [];
  debugInfo.push(`Initial state: Counter = ${quizSheet.getRange('C1').getValue()}, Question = "${quizSheet.getRange('A4').getValue()}"`);

  for (let i = 0; i < 4; i++) {
    const currentCount = getQuestionCounter(quizSheet);
    debugInfo.push(`Before question ${i + 2}: Counter = ${currentCount}`);

    showNextQuestion(quizSheet, ss);

    const newCount = getQuestionCounter(quizSheet);
    const questionText = quizSheet.getRange('A4').getValue();
    debugInfo.push(`After question ${i + 2}: Counter = ${newCount}, Question = "${questionText}"`);
  }

  const beforeFinalCount = getQuestionCounter(quizSheet);
  debugInfo.push(`Before final call: Counter = ${beforeFinalCount}`);

  showNextQuestion(quizSheet, ss);

  const finalQuestionText = quizSheet.getRange('A4').getValue();
  const finalCounter = quizSheet.getRange('C1').getValue();
  const startQuizUnchecked = quizSheet.getRange('B2').getValue() === false;
  const rightCheckboxDisabled = isRangeProtected(quizSheet, 'B5');
  const wrongCheckboxDisabled = isRangeProtected(quizSheet, 'B6');

  debugInfo.push(`Final state: Counter = ${finalCounter}, Question = "${finalQuestionText}", Start Quiz unchecked = ${startQuizUnchecked}`);

  const isCompletionMessage = finalQuestionText && finalQuestionText.toString().includes('Quiz Complete');
  const questionCounterReset = finalCounter === 0;

  const testPassed = isCompletionMessage && startQuizUnchecked && rightCheckboxDisabled && wrongCheckboxDisabled && questionCounterReset;

  recordTestResult(
    'Quiz should complete after 5 questions with completion message',
    testPassed,
    testPassed ?
      '✓ Quiz properly completed after 5 questions with all expected behaviors' :
      `✗ Debug info: ${debugInfo.join(' | ')} | Final checks: Completion message found: ${isCompletionMessage}, Counter reset: ${questionCounterReset}, Start Quiz unchecked: ${startQuizUnchecked}, Right disabled: ${rightCheckboxDisabled}, Wrong disabled: ${wrongCheckboxDisabled}`
  );

  return testPassed;
}

// Test that verifies question counter resets when category is changed
function testQuestionCounterResetsOnCategoryChange() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const quizSheet = ss.getSheetByName('quiz');
  const datastoreSheet = ss.getSheetByName('datastore');

  const data = datastoreSheet.getRange('B2:B' + datastoreSheet.getLastRow()).getValues();
  const validCategories = [...new Set(data.flat().filter(Boolean))];
  const category1 = validCategories.length > 0 ? validCategories[0] : 'Rishis';
  const category2 = validCategories.length > 1 ? validCategories[1] : validCategories[0];

  quizSheet.getRange('A1').setValue(category1);
  quizSheet.getRange('B2').setValue(true);

  const startQuizEvent = {
    source: ss,
    range: quizSheet.getRange('B2'),
    value: true,
    oldValue: false
  };
  handleCheckboxEdit(startQuizEvent);

  for (let i = 0; i < 2; i++) {
    const rightClickEvent = {
      source: ss,
      range: quizSheet.getRange('B5'),
      value: true,
      oldValue: false
    };
    handleCheckboxEdit(rightClickEvent);
  }

  const counterBefore = quizSheet.getRange('C1').getValue();

  const categoryChangeEvent = {
    source: ss,
    range: quizSheet.getRange('A1'),
    value: category2,
    oldValue: category1
  };
  handleCheckboxEdit(categoryChangeEvent);

  const counterAfter = quizSheet.getRange('C1').getValue();
  const counterReset = counterAfter === 0;

  const testPassed = counterBefore > 0 && counterReset;

  recordTestResult(
    'Question counter should reset when category is changed',
    testPassed,
    testPassed ?
      '✓ Question counter properly reset when category changed' :
      `✗ Counter before: ${counterBefore}, Counter after: ${counterAfter}, Reset: ${counterReset}`
  );

  return testPassed;
}

// Test that verifies questions are not repeated during a quiz session
function testQuestionsNotRepeated() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const quizSheet = ss.getSheetByName('quiz');
  const datastoreSheet = ss.getSheetByName('datastore');

  try {
    const data = datastoreSheet.getDataRange().getValues();
    const categoryQuestions = {};

    data.slice(1).forEach(row => {
      if (row[1] && row[2]) {
        if (!categoryQuestions[row[1]]) {
          categoryQuestions[row[1]] = [];
        }
        categoryQuestions[row[1]].push(row[2]);
      }
    });

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

    quizSheet.getRange('A1').setValue('');
    quizSheet.getRange('B2').setValue(false);
    quizSheet.getRange('A4').setValue('');
    quizSheet.getRange('C1').setValue('');
    quizSheet.getRange('D1').setValue('');
    quizSheet.getRange('B5').setValue(false);
    quizSheet.getRange('B6').setValue(false);

    const protections = quizSheet.getProtections(SpreadsheetApp.ProtectionType.RANGE);
    protections.forEach(protection => {
      const range = protection.getRange();
      if (range.getA1Notation() === 'B5' || range.getA1Notation() === 'B6') {
        protection.remove();
      }
    });

    const categoryEvent = createMockEditEvent(ss, quizSheet.getRange('A1'), testCategory, '');
    handleCheckboxEdit(categoryEvent);

    const startQuizEvent = createMockEditEvent(ss, quizSheet.getRange('B2'), true, false);
    handleCheckboxEdit(startQuizEvent);

    const usedQuestions = [];
    let hasRepeats = false;
    let debugInfo = [];

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

    for (let i = 1; i < 5; i++) {
      const rightClickEvent = createMockEditEvent(ss, quizSheet.getRange('B5'), true, false);
      handleCheckboxEdit(rightClickEvent);

      Utilities.sleep(100);

      const currentQuestion = quizSheet.getRange('A4').getValue();

      if (currentQuestion && currentQuestion.toString().includes('Quiz Complete')) {
        debugInfo.push(`Quiz completed at question ${i + 1}`);
        break;
      }

      if (currentQuestion && currentQuestion.toString().includes('No questions')) {
        debugInfo.push(`No more questions at question ${i + 1}`);
        break;
      }

      if (currentQuestion && currentQuestion !== '') {
        const questionStr = currentQuestion.toString();
        debugInfo.push(`Q${i + 1}: "${questionStr}"`);

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

// Test that verifies used questions list is reset when category changes
function testUsedQuestionsResetOnCategoryChange() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const quizSheet = ss.getSheetByName('quiz');
  const datastoreSheet = ss.getSheetByName('datastore');

  const data = datastoreSheet.getDataRange().getValues();
  const validCategories = [...new Set(data.slice(1).map(row => row[1]).filter(Boolean))];
  const category1 = validCategories.length > 0 ? validCategories[0] : 'Rishis';
  const category2 = validCategories.length > 1 ? validCategories[1] : validCategories[0];

  quizSheet.getRange('A1').setValue(category1);
  quizSheet.getRange('B2').setValue(true);

  const startQuizEvent = {
    source: ss,
    range: quizSheet.getRange('B2'),
    value: true,
    oldValue: false
  };
  handleCheckboxEdit(startQuizEvent);

  const rightClickEvent = {
    source: ss,
    range: quizSheet.getRange('B5'),
    value: true,
    oldValue: false
  };
  handleCheckboxEdit(rightClickEvent);

  const usedQuestionsBefore = getUsedQuestions(quizSheet);
  const hasUsedQuestionsBefore = usedQuestionsBefore.length > 0;

  const categoryChangeEvent = {
    source: ss,
    range: quizSheet.getRange('A1'),
    value: category2,
    oldValue: category1
  };
  handleCheckboxEdit(categoryChangeEvent);

  const usedQuestionsAfter = getUsedQuestions(quizSheet);
  const usedQuestionsReset = usedQuestionsAfter.length === 0;

  const testPassed = hasUsedQuestionsBefore && usedQuestionsReset;

  recordTestResult(
    'Used questions list should reset when category changes',
    testPassed,
    testPassed ?
      '✓ Used questions list properly reset when category changed' :
      `✗ Reset failed. Before: ${usedQuestionsBefore.length} questions, After: ${usedQuestionsAfter.length} questions`
  );

  return testPassed;
}

// Test that verifies used questions list is reset when quiz completes
function testUsedQuestionsResetOnQuizComplete() {
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

  quizSheet.getRange('B2').setValue(true);

  const filtered = datastoreSheet.getDataRange().getValues().filter((row, index) => index !== 0 && row[1] === testCategory);
  if (filtered.length > 0) {
    resetUsedQuestions(quizSheet);
    const random = filtered[Math.floor(Math.random() * filtered.length)];
    quizSheet.getRange('A4').setValue(random[2]);
    addUsedQuestion(quizSheet, random[2]);
    quizSheet.getRange('C1').setValue(1);
    toggleRightWrongCheckboxes(quizSheet, true);
  }

  for (let i = 0; i < 4; i++) {
    showNextQuestion(quizSheet, ss);
  }

  const usedQuestionsBeforeComplete = getUsedQuestions(quizSheet);
  const hasUsedQuestionsBeforeComplete = usedQuestionsBeforeComplete.length > 0;

  showNextQuestion(quizSheet, ss);

  const usedQuestionsAfterComplete = getUsedQuestions(quizSheet);
  const usedQuestionsResetAfterComplete = usedQuestionsAfterComplete.length === 0;

  const finalQuestionText = quizSheet.getRange('A4').getValue();
  const showsCompletionMessage = finalQuestionText && finalQuestionText.toString().includes('Quiz Complete');

  const testPassed = hasUsedQuestionsBeforeComplete && usedQuestionsResetAfterComplete && showsCompletionMessage;

  recordTestResult(
    'Used questions list should reset when quiz completes',
    testPassed,
    testPassed ?
      '✓ Used questions list properly reset when quiz completed' :
      `✗ Reset failed. Before complete: ${usedQuestionsBeforeComplete.length} questions, After complete: ${usedQuestionsAfterComplete.length} questions, Shows completion: ${showsCompletionMessage}`
  );

  return testPassed;
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
function runAllTests() {
  clearTestResults();

  testCategoryClearsQuestionCell();
  testCategoryClearsCheckbox();
  testRightWrongCheckboxesDisabledWhenQuizNotStarted();
  testRightWrongCheckboxesEnabledWhenQuizStarted();
  testRightWrongCheckboxesAutoUncheckWhenQuizNotStarted();
  testQuizCompletesAfter5Questions();
  testQuestionCounterResetsOnCategoryChange();
  testQuestionsNotRepeated();
  testUsedQuestionsTracking();
  testUsedQuestionsResetOnCategoryChange();
  testUsedQuestionsResetOnQuizComplete();
  testQuizEndsEarlyWhenNoMoreUniqueQuestions()
}

