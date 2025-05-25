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
    'testCategoryClearsCheckbox',
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
    'testRightWrongCheckboxesDisabledWhenQuizNotStarted',
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
    'testRightWrongCheckboxesEnabledWhenQuizStarted',
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
    'testRightWrongCheckboxesAutoUncheckWhenQuizNotStarted',
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
    'testQuizCompletesAfter5Questions',
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
    'testQuestionCounterResetsOnCategoryChange',
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
      'testQuestionsNotRepeated',
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

  // Clear and setup initial state
  quizSheet.getRange('A1').setValue(testCategory);
  quizSheet.getRange('B2').setValue(false);
  quizSheet.getRange('A4').setValue('');
  quizSheet.getRange('C1').setValue(0);
  quizSheet.getRange('D1').setValue('');
  quizSheet.getRange('B5').setValue(false);
  quizSheet.getRange('B6').setValue(false);

  const initialUsedQuestions = getUsedQuestions(quizSheet);
  const initiallyEmpty = Array.isArray(initialUsedQuestions) && initialUsedQuestions.length === 0;

  // Create proper mock event with getA1Notation method
  const startQuizRange = quizSheet.getRange('B2');
  const startQuizEvent = {
    source: ss,
    range: {
      getValue: () => true,
      setValue: (value) => startQuizRange.setValue(value),
      getA1Notation: () => 'B2'
    }
  };

  handleCheckboxEdit(startQuizEvent);

  const firstQuestion = quizSheet.getRange('A4').getValue();
  const usedQuestionsAfterFirst = getUsedQuestions(quizSheet);
  const firstQuestionTracked = firstQuestion && firstQuestion !== '' && usedQuestionsAfterFirst.includes(firstQuestion);

  // Create proper mock event for right checkbox
  const rightCheckboxRange = quizSheet.getRange('B5');
  const rightClickEvent = {
    source: ss,
    range: {
      getValue: () => true,
      setValue: (value) => rightCheckboxRange.setValue(value),
      getA1Notation: () => 'B5'
    }
  };

  handleCheckboxEdit(rightClickEvent);

  const secondQuestion = quizSheet.getRange('A4').getValue();
  const usedQuestionsAfterSecond = getUsedQuestions(quizSheet);
  const secondQuestionTracked = !secondQuestion.includes('Quiz Complete') && usedQuestionsAfterSecond.includes(secondQuestion);
  const bothQuestionsTracked = usedQuestionsAfterSecond.length >= 2 || secondQuestion.includes('Quiz Complete');

  const testPassed = initiallyEmpty && firstQuestionTracked && (secondQuestionTracked || secondQuestion.includes('Quiz Complete') || usedQuestionsAfterSecond.length >= 2);

  recordTestResult(
    'testUsedQuestionsTracking',
    'Used questions should be properly tracked in cell D1',
    testPassed,
    testPassed ?
      `✓ Used questions properly tracked. Initial: ${initialUsedQuestions.length}, After first: ${usedQuestionsAfterFirst.length}, After second: ${usedQuestionsAfterSecond.length}` :
      `✗ Tracking failed. Initial empty: ${initiallyEmpty}, First tracked: ${firstQuestionTracked}, Second tracked: ${secondQuestionTracked}, Questions: ["${firstQuestion}", "${secondQuestion}"], Used after first: ${JSON.stringify(usedQuestionsAfterFirst)}, Used after second: ${JSON.stringify(usedQuestionsAfterSecond)}`
  );

  return testPassed;
}


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
    'testUsedQuestionsResetOnCategoryChange',
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
    'testUsedQuestionsResetOnQuizComplete',
    'Used questions list should reset when quiz completes',
    testPassed,
    testPassed ?
      '✓ Used questions list properly reset when quiz completed' :
      `✗ Reset failed. Before complete: ${usedQuestionsBeforeComplete.length} questions, After complete: ${usedQuestionsAfterComplete.length} questions, Shows completion: ${showsCompletionMessage}`
  );

  return testPassed;
}

// Test that Show Answer checkbox is disabled when quiz is not started
function testShowAnswerCheckboxDisabledWhenQuizNotStarted() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const quizSheet = ss.getSheetByName('quiz');

  // Reset quiz state
  quizSheet.getRange('B2').setValue(false); // Uncheck Start Quiz
  quizSheet.getRange('B7').setValue(false); // Uncheck Show Answer
  quizSheet.getRange('A4').setValue(''); // Clear question
  quizSheet.getRange('A8').setValue(''); // Clear answer

  // Trigger the edit event to disable checkboxes
  const stopQuizEvent = {
    source: ss,
    range: quizSheet.getRange('B2'),
    value: false,
    oldValue: true
  };
  handleCheckboxEdit(stopQuizEvent);

  // Check if Show Answer checkbox is protected (disabled)
  const isShowAnswerProtected = isRangeProtected(quizSheet, 'B7');

  recordTestResult(
    'testShowAnswerCheckboxDisabledWhenQuizNotStarted',
    'Show Answer checkbox should be disabled when quiz is not started',
    isShowAnswerProtected,
    isShowAnswerProtected ?
      '✓ Show Answer checkbox is properly disabled when quiz is not started' :
      '✗ Show Answer checkbox is not disabled when quiz is not started'
  );

  return isShowAnswerProtected;
}

// Test that Show Answer checkbox shows the correct answer when checked
function testShowAnswerDisplaysCorrectAnswer() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const quizSheet = ss.getSheetByName('quiz');
  const datastoreSheet = ss.getSheetByName('datastore');

  // Get datastore data and find a valid category with questions
  const data = datastoreSheet.getDataRange().getValues();

  if (data.length <= 1) {
    recordTestResult(
      'Show Answer checkbox should display the correct answer when checked',
      false,
      '✗ No data found in datastore sheet'
    );
    return false;
  }

  // Find a category that actually has questions
  let testCategory = null;
  let testQuestionRow = null;

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (row[1] && row[2] && row[3]) { // Category, Question, and Answer all exist
      testCategory = row[1];
      testQuestionRow = row;
      break;
    }
  }

  if (!testCategory || !testQuestionRow) {
    recordTestResult(
      'Show Answer checkbox should display the correct answer when checked',
      false,
      '✗ No valid category with question and answer found in datastore'
    );
    return false;
  }

  // Clear any existing state completely
  quizSheet.getRange('A1').setValue(''); // Clear category first
  quizSheet.getRange('B2').setValue(false);
  quizSheet.getRange('B7').setValue(false);
  quizSheet.getRange('A4').setValue(''); // Clear question cell
  quizSheet.getRange('A8').setValue(''); // Clear answer cell
  quizSheet.getRange('C1').setValue(0); // Reset counter
  quizSheet.getRange('D1').setValue(''); // Clear used questions

  // Wait a moment to ensure state is cleared
  Utilities.sleep(100);

  // Set category
  quizSheet.getRange('A1').setValue(testCategory);

  // Wait a moment after setting category
  Utilities.sleep(100);

  // Start quiz by setting checkbox to true
  quizSheet.getRange('B2').setValue(true);

  // Trigger the edit event manually
  const startQuizEvent = {
    source: ss,
    range: quizSheet.getRange('B2'),
    value: true,
    oldValue: false
  };
  handleCheckboxEdit(startQuizEvent);

  // Get the current question after quiz starts
  const currentQuestion = quizSheet.getRange('A4').getValue();

  // Debugging: Check if question was loaded
  if (!currentQuestion || currentQuestion === '') {
    // Try to debug by checking what's in the datastore for this category
    const categoryData = data.filter((row, index) => index !== 0 && row[1] === testCategory);

    recordTestResult(
      'testShowAnswerCheckboxDisabledWhenQuizNotStarted',
      'Show Answer checkbox should display the correct answer when checked',
      false,
      `✗ No question loaded after starting quiz. Category: "${testCategory}", Questions in category: ${categoryData.length}, First question in category: "${categoryData.length > 0 ? categoryData[0][2] : 'none'}"`
    );
    return false;
  }

  // Find the expected answer from datastore for the current question
  let expectedAnswer = '';

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (row[1] === testCategory && row[2] === currentQuestion) {
      expectedAnswer = row[3] || '';
      break;
    }
  }

  // Check Show Answer checkbox
  quizSheet.getRange('B7').setValue(true);
  const showAnswerEvent = {
    source: ss,
    range: quizSheet.getRange('B7'),
    value: true,
    oldValue: false
  };
  handleCheckboxEdit(showAnswerEvent);

  // Get the displayed answer
  const displayedAnswer = quizSheet.getRange('A8').getValue();

  // Test passes if answer is displayed and matches expected answer
  const testPassed = displayedAnswer !== '' && displayedAnswer === expectedAnswer;

  recordTestResult(
    'testShowAnswerDisplaysCorrectAnswer',
    'Show Answer checkbox should display the correct answer when checked',
    testPassed,
    testPassed ?
      '✓ Show Answer checkbox correctly displays the answer' :
      `✗ Show Answer failed. Question: "${currentQuestion}", Expected: "${expectedAnswer}", Got: "${displayedAnswer}"`
  );

  return testPassed;
}

// Test that Show Answer checkbox hides the answer when unchecked
function testShowAnswerHidesAnswerWhenUnchecked() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const quizSheet = ss.getSheetByName('quiz');
  const datastoreSheet = ss.getSheetByName('datastore');

  // Get datastore data and find a valid category with questions
  const data = datastoreSheet.getDataRange().getValues();

  if (data.length <= 1) {
    recordTestResult(
      'Show Answer checkbox should hide the answer when unchecked',
      false,
      '✗ No data found in datastore sheet'
    );
    return false;
  }

  // Find a category that actually has questions
  let testCategory = null;

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (row[1] && row[2] && row[3]) { // Category, Question, and Answer all exist
      testCategory = row[1];
      break;
    }
  }

  if (!testCategory) {
    recordTestResult(
      'Show Answer checkbox should hide the answer when unchecked',
      false,
      '✗ No valid category with question and answer found in datastore'
    );
    return false;
  }

  // Clear any existing state completely
  quizSheet.getRange('A1').setValue(''); // Clear category first  
  quizSheet.getRange('B2').setValue(false);
  quizSheet.getRange('B7').setValue(false);
  quizSheet.getRange('A4').setValue(''); // Clear question cell
  quizSheet.getRange('A8').setValue(''); // Clear answer cell
  quizSheet.getRange('C1').setValue(0); // Reset counter
  quizSheet.getRange('D1').setValue(''); // Clear used questions

  // Wait a moment to ensure state is cleared
  Utilities.sleep(100);

  // Set category
  quizSheet.getRange('A1').setValue(testCategory);

  // Wait a moment after setting category
  Utilities.sleep(100);

  // Start quiz
  quizSheet.getRange('B2').setValue(true);
  const startQuizEvent = {
    source: ss,
    range: quizSheet.getRange('B2'),
    value: true,
    oldValue: false
  };
  handleCheckboxEdit(startQuizEvent);

  // Verify a question was loaded
  const currentQuestion = quizSheet.getRange('A4').getValue();
  if (!currentQuestion || currentQuestion === '') {
    recordTestResult(
      'Show Answer checkbox should hide the answer when unchecked',
      false,
      `✗ No question loaded after starting quiz with category: "${testCategory}"`
    );
    return false;
  }

  // Check Show Answer checkbox first
  quizSheet.getRange('B7').setValue(true);
  const showAnswerEvent = {
    source: ss,
    range: quizSheet.getRange('B7'),
    value: true,
    oldValue: false
  };
  handleCheckboxEdit(showAnswerEvent);

  // Verify answer is displayed
  const answerWhenChecked = quizSheet.getRange('A8').getValue();

  // Uncheck Show Answer checkbox
  quizSheet.getRange('B7').setValue(false);
  const hideAnswerEvent = {
    source: ss,
    range: quizSheet.getRange('B7'),
    value: false,
    oldValue: true
  };
  handleCheckboxEdit(hideAnswerEvent);

  // Get the answer after unchecking
  const answerWhenUnchecked = quizSheet.getRange('A8').getValue();

  const testPassed = answerWhenChecked !== '' && answerWhenUnchecked === '';

  recordTestResult(
    'testShowAnswerHidesAnswerWhenUnchecked',
    'Show Answer checkbox should hide the answer when unchecked',
    testPassed,
    testPassed ?
      '✓ Show Answer checkbox correctly hides the answer when unchecked' :
      `✗ Show Answer hide failed. Answer when checked: "${answerWhenChecked}", Answer when unchecked: "${answerWhenUnchecked}"`
  );

  return testPassed;
}
// Test that verifies score display labels are correctly initialized
function testScoreDisplayInitialization() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const quizSheet = ss.getSheetByName('quiz');

  // Run setup to ensure labels are set
  setupQuizSheet();

  const rightLabel = quizSheet.getRange('A9').getValue();
  const wrongLabel = quizSheet.getRange('A10').getValue();

  const labelsCorrect = (rightLabel === 'Right Answers:' && wrongLabel === 'Wrong Answers:');

  recordTestResult(
    'testScoreDisplayInitialization',
    'Score display labels should be correctly initialized',
    labelsCorrect,
    labelsCorrect ?
      '✓ Labels correctly set: A9="Right Answers:", A10="Wrong Answers:"' :
      `✗ Labels incorrect: A9="${rightLabel}", A10="${wrongLabel}"`
  );

  return labelsCorrect;
}

// Test that verifies score display initial values are 0
function testScoreDisplayInitialValues() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const quizSheet = ss.getSheetByName('quiz');

  // Run setup to ensure initial values are set
  setupQuizSheet();

  const rightCount = quizSheet.getRange('B9').getValue();
  const wrongCount = quizSheet.getRange('B10').getValue();

  const valuesCorrect = (rightCount === 0 && wrongCount === 0);

  recordTestResult(
    'testScoreDisplayInitialValues',
    'Score display initial values should be 0',
    valuesCorrect,
    valuesCorrect ?
      '✓ Initial values correct: B9=0, B10=0' :
      `✗ Initial values incorrect: B9=${rightCount}, B10=${wrongCount}`
  );

  return valuesCorrect;
}

// Test score reset when category changes
function testScoreResetOnCategoryChange() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const quizSheet = ss.getSheetByName('quiz');

  // Setup initial state with some scores
  quizSheet.getRange('B9').setValue(5); // Set right count to 5
  quizSheet.getRange('B10').setValue(3); // Set wrong count to 3

  // Mock category change
  const mockEvent = {
    source: ss,
    range: quizSheet.getRange('A1'),
    value: 'New Category',
    oldValue: 'Old Category'
  };

  handleCheckboxEdit(mockEvent);

  const rightCount = quizSheet.getRange('B9').getValue();
  const wrongCount = quizSheet.getRange('B10').getValue();

  const scoresReset = (rightCount === 0 && wrongCount === 0);

  recordTestResult(
    'testScoreResetOnCategoryChange',
    'Scores should reset to 0 when category changes',
    scoresReset,
    scoresReset ?
      '✓ Both scores reset to 0 after category change' :
      `✗ Scores not reset: Right=${rightCount}, Wrong=${wrongCount}`
  );

  return scoresReset;
}

// Test score reset when quiz starts
function testScoreResetOnQuizStart() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const quizSheet = ss.getSheetByName('quiz');

  // Setup initial state with some scores
  quizSheet.getRange('A1').setValue('Test Category');
  quizSheet.getRange('B9').setValue(7); // Set right count to 7
  quizSheet.getRange('B10').setValue(4); // Set wrong count to 4

  // Mock Start Quiz checkbox being checked
  const mockEvent = {
    source: ss,
    range: quizSheet.getRange('B2'),
    value: true,
    oldValue: false
  };

  handleCheckboxEdit(mockEvent);

  const rightCount = quizSheet.getRange('B9').getValue();
  const wrongCount = quizSheet.getRange('B10').getValue();

  const scoresReset = (rightCount === 0 && wrongCount === 0);

  recordTestResult(
    'testScoreResetOnQuizStart',
    'Scores should reset to 0 when quiz starts',
    scoresReset,
    scoresReset ?
      '✓ Both scores reset to 0 when quiz started' :
      `✗ Scores not reset on quiz start: Right=${rightCount}, Wrong=${wrongCount}`
  );

  return scoresReset;
}

// Test score reset when quiz stops
function testScoreResetOnQuizStop() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const quizSheet = ss.getSheetByName('quiz');

  // Setup initial state with some scores and active quiz
  quizSheet.getRange('A1').setValue('Test Category');
  quizSheet.getRange('B2').setValue(true); // Quiz is active
  quizSheet.getRange('B9').setValue(6); // Set right count to 6
  quizSheet.getRange('B10').setValue(2); // Set wrong count to 2

  // Mock Start Quiz checkbox being unchecked (stop quiz)
  const mockEvent = {
    source: ss,
    range: quizSheet.getRange('B2'),
    value: false,
    oldValue: true
  };

  handleCheckboxEdit(mockEvent);

  const rightCount = quizSheet.getRange('B9').getValue();
  const wrongCount = quizSheet.getRange('B10').getValue();

  const scoresReset = (rightCount === 0 && wrongCount === 0);

  recordTestResult(
    'testScoreResetOnQuizStop',
    'Scores should reset to 0 when quiz stops',
    scoresReset,
    scoresReset ?
      '✓ Both scores reset to 0 when quiz stopped' :
      `✗ Scores not reset on quiz stop: Right=${rightCount}, Wrong=${wrongCount}`
  );

  return scoresReset;
}

// Test that scores don't increment when quiz is not started
function testNoIncrementWhenQuizNotStarted() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const quizSheet = ss.getSheetByName('quiz');

  // Setup initial state without starting quiz
  quizSheet.getRange('A1').setValue('Test Category');
  quizSheet.getRange('B2').setValue(false); // Quiz not started
  quizSheet.getRange('B9').setValue(0); // Reset counts
  quizSheet.getRange('B10').setValue(0);

  // Try to check Right checkbox
  const rightMockEvent = {
    source: ss,
    range: quizSheet.getRange('B5'),
    value: true,
    oldValue: false
  };

  handleCheckboxEdit(rightMockEvent);

  const rightCountAfterRight = quizSheet.getRange('B9').getValue();

  // Try to check Wrong checkbox
  const wrongMockEvent = {
    source: ss,
    range: quizSheet.getRange('B6'),
    value: true,
    oldValue: false
  };

  handleCheckboxEdit(wrongMockEvent);

  const rightCountAfterWrong = quizSheet.getRange('B9').getValue();
  const wrongCountAfterWrong = quizSheet.getRange('B10').getValue();

  const countsUnchanged = (rightCountAfterRight === 0 && rightCountAfterWrong === 0 && wrongCountAfterWrong === 0);

  recordTestResult(
    'testNoIncrementWhenQuizNotStarted',
    'Scores should not increment when quiz is not started',
    countsUnchanged,
    countsUnchanged ?
      '✓ Scores remained at 0 when checkboxes clicked without active quiz' :
      `✗ Scores incremented without active quiz: Right=${rightCountAfterWrong}, Wrong=${wrongCountAfterWrong}`
  );

  return countsUnchanged;
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

  // Enable the checkboxes by removing protection (simulate quiz started state)
  toggleRightWrongCheckboxes(quizSheet, true);

  const initialRightCount = quizSheet.getRange('B9').getValue();

  // Create proper mock event with the checkbox value set
  const checkboxRange = quizSheet.getRange('B5');
  checkboxRange.setValue(true); // Set the checkbox value first

  const mockEvent = {
    source: ss,
    range: checkboxRange,
    value: true,
    oldValue: false
  };

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

  // Enable the checkboxes by removing protection (simulate quiz started state)
  toggleRightWrongCheckboxes(quizSheet, true);

  const initialWrongCount = quizSheet.getRange('B10').getValue();

  // Create proper mock event with the checkbox value set
  const checkboxRange = quizSheet.getRange('B6');
  checkboxRange.setValue(true); // Set the checkbox value first

  const mockEvent = {
    source: ss,
    range: checkboxRange,
    value: true,
    oldValue: false
  };

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

  // Enable the checkboxes by removing protection (simulate quiz started state)
  toggleRightWrongCheckboxes(quizSheet, true);

  let allIncrementsCorrect = true;
  let details = '';

  // Check Right checkbox 3 times
  for (let i = 1; i <= 3; i++) {
    const checkboxRange = quizSheet.getRange('B5');
    checkboxRange.setValue(true); // Set the checkbox value first

    const mockEvent = {
      source: ss,
      range: checkboxRange,
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

  // Enable the checkboxes by removing protection (simulate quiz started state)
  toggleRightWrongCheckboxes(quizSheet, true);

  let allIncrementsCorrect = true;
  let details = '';

  // Check Wrong checkbox 3 times
  for (let i = 1; i <= 3; i++) {
    const checkboxRange = quizSheet.getRange('B6');
    checkboxRange.setValue(true); // Set the checkbox value first

    const mockEvent = {
      source: ss,
      range: checkboxRange,
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

  // Enable the checkboxes by removing protection (simulate quiz started state)
  toggleRightWrongCheckboxes(quizSheet, true);

  const rightCountBefore = quizSheet.getRange('B9').getValue();
  const wrongCountBefore = quizSheet.getRange('B10').getValue();

  // Simulate answering correctly (this would typically trigger next question)
  const checkboxRange = quizSheet.getRange('B5');
  checkboxRange.setValue(true); // Set the checkbox value first
  
  const mockEvent = {
    source: ss,
    range: checkboxRange,
    value: true,
    oldValue: false
  };

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
function runAllTests() {
  clearTestResults();

  /*testCategoryClearsQuestionCell();
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
  testShowAnswerCheckboxDisabledWhenQuizNotStarted();
  testShowAnswerDisplaysCorrectAnswer();
  testShowAnswerHidesAnswerWhenUnchecked();
  testScoreDisplayInitialization();
  testScoreDisplayInitialValues();
  testScoreResetOnCategoryChange();
  testScoreResetOnQuizStart();
  testScoreResetOnQuizStop();
  testNoIncrementWhenQuizNotStarted();
  testRightAnswerIncrement();
  testWrongAnswerIncrement();
  testMultipleRightAnswers();
  testMultipleWrongAnswers();
  testScorePreservationDuringQuizProgress();*/
  testMixedAnswers();
  


}

