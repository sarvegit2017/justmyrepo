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

// Helper function to setup test data in datastore sheet
function setupTestDatastore() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let datastoreSheet = ss.getSheetByName('datastore');

  if (!datastoreSheet) {
    datastoreSheet = ss.insertSheet('datastore');
  }

  // Clear existing data and setup test data
  datastoreSheet.clear();
  datastoreSheet.getRange('A1:D6').setValues([
    ['SL#', 'Category', 'Questions', 'Answers'],
    [1, 'Test Category', 'Question 1', 'Answer 1'],
    [2, 'Test Category', 'Question 2', 'Answer 2'],
    [3, 'Test Category', 'Question 3', 'Answer 3'],
    [4, 'Other Category', 'Question 4', 'Answer 4'],
    [5, 'Test Category', 'Question 5', 'Answer 5']
  ]);
}

// Helper function to reset quiz sheet to clean state
function resetQuizSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const quizSheet = ss.getSheetByName('quiz');

  // Clear all values and reset state
  quizSheet.getRange('A1').setValue('Test Category');
  quizSheet.getRange('B2').setValue(false); // Start Quiz
  quizSheet.getRange('B3').setValue(false); // Retry Wrong Questions
  quizSheet.getRange('A4').setValue(''); // Question
  quizSheet.getRange('B5').setValue(false); // Right
  quizSheet.getRange('B6').setValue(false); // Wrong
  quizSheet.getRange('B7').setValue(false); // Show Answer
  quizSheet.getRange('A8').setValue(''); // Answer
  quizSheet.getRange('B9').setValue(0); // Right count
  quizSheet.getRange('B10').setValue(0); // Wrong count

  // Reset counters and lists
  resetQuestionCounter(quizSheet);
  resetUsedQuestions(quizSheet);
  resetWrongQuestions(quizSheet);

  // Remove all protections
  toggleRightWrongCheckboxes(quizSheet, false);
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



// 7. State Management Tests
function testWrongQuestionsListPersistence() {
  setupTestDatastore();
  resetQuizSheet();

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const quizSheet = ss.getSheetByName('quiz');

  // Add questions to wrong questions list
  const testQuestions = ['Question 1', 'Question 2', 'Question 3'];
  setWrongQuestions(quizSheet, testQuestions);

  // Start and stop normal quiz (should not affect wrong questions list)
  quizSheet.getRange('B2').setValue(true);
  const startEvent = createMockEditEvent(ss, quizSheet.getRange('B2'), true, false);
  handleCheckboxEdit(startEvent);

  quizSheet.getRange('B2').setValue(false);
  const stopEvent = createMockEditEvent(ss, quizSheet.getRange('B2'), false, true);
  handleCheckboxEdit(stopEvent);

  const wrongQuestionsAfter = getWrongQuestions(quizSheet);

  const passed = (JSON.stringify(testQuestions.sort()) === JSON.stringify(wrongQuestionsAfter.sort()));

  recordTestResult(
    'testWrongQuestionsListPersistence',
    'Wrong questions list should persist across quiz sessions',
    passed,
    `Original: [${testQuestions.join(', ')}]. After quiz cycle: [${wrongQuestionsAfter.join(', ')}]`
  );

  return passed;
}

function testWrongQuestionsResetInNormalMode() {
  setupTestDatastore();
  resetQuizSheet();

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const quizSheet = ss.getSheetByName('quiz');

  // Add questions to wrong questions list
  setWrongQuestions(quizSheet, ['Question 1', 'Question 2']);
  const wrongQuestionsBefore = getWrongQuestions(quizSheet);

  // Start normal quiz (not retry mode)
  quizSheet.getRange('B3').setValue(false); // Ensure retry mode is off
  quizSheet.getRange('B2').setValue(true);
  const startEvent = createMockEditEvent(ss, quizSheet.getRange('B2'), true, false);
  handleCheckboxEdit(startEvent);

  const wrongQuestionsAfter = getWrongQuestions(quizSheet);

  const passed = (wrongQuestionsBefore.length > 0 && wrongQuestionsAfter.length === 0);

  recordTestResult(
    'testWrongQuestionsResetInNormalMode',
    'Wrong questions list should be reset when starting normal quiz',
    passed,
    `Before normal quiz: ${wrongQuestionsBefore.length} questions. After: ${wrongQuestionsAfter.length} questions`
  );

  return passed;
}

// 1. Basic Retry Mode Activation Tests
function testRetryModeStopsRunningQuiz() {
  setupTestDatastore();
  resetQuizSheet();

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const quizSheet = ss.getSheetByName('quiz');

  // Start a quiz first
  quizSheet.getRange('B2').setValue(true);
  const startEvent = createMockEditEvent(ss, quizSheet.getRange('B2'), true, false);
  handleCheckboxEdit(startEvent);

  const questionBefore = quizSheet.getRange('A4').getValue();
  const quizRunning = quizSheet.getRange('B2').getValue();

  // Now check Retry Wrong Questions
  quizSheet.getRange('B3').setValue(true);
  const retryEvent = createMockEditEvent(ss, quizSheet.getRange('B3'), true, false);
  handleCheckboxEdit(retryEvent);

  const questionAfter = quizSheet.getRange('A4').getValue();
  const quizStoppedAfter = quizSheet.getRange('B2').getValue();

  const passed = (questionBefore !== '' && quizRunning === true && questionAfter === '' && quizStoppedAfter === false);

  recordTestResult(
    'testRetryModeStopsRunningQuiz',
    'Checking Retry Wrong Questions should stop running quiz',
    passed,
    `Before: Quiz running=${quizRunning}, Question="${questionBefore}". After: Quiz running=${quizStoppedAfter}, Question="${questionAfter}"`
  );

  return passed;
}

function testRetryModeToggleWhenQuizNotRunning() {
  resetQuizSheet();

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const quizSheet = ss.getSheetByName('quiz');

  // Check retry checkbox when quiz is not running
  quizSheet.getRange('B3').setValue(true);
  const checkEvent = createMockEditEvent(ss, quizSheet.getRange('B3'), true, false);
  handleCheckboxEdit(checkEvent);

  const checkedState = quizSheet.getRange('B3').getValue();

  // Uncheck retry checkbox
  quizSheet.getRange('B3').setValue(false);
  const uncheckEvent = createMockEditEvent(ss, quizSheet.getRange('B3'), false, true);
  handleCheckboxEdit(uncheckEvent);

  const uncheckedState = quizSheet.getRange('B3').getValue();

  const passed = (checkedState === true && uncheckedState === false);

  recordTestResult(
    'testRetryModeToggleWhenQuizNotRunning',
    'Retry Wrong Questions should be toggleable when quiz is not running',
    passed,
    `Checked state: ${checkedState}, Unchecked state: ${uncheckedState}`
  );

  return passed;
}

// 2. Wrong Questions Collection Tests
function testWrongAnswerAddsToList() {
  setupTestDatastore();
  resetQuizSheet();

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const quizSheet = ss.getSheetByName('quiz');

  // Start quiz
  quizSheet.getRange('B2').setValue(true);
  const startEvent = createMockEditEvent(ss, quizSheet.getRange('B2'), true, false);
  handleCheckboxEdit(startEvent);

  const currentQuestion = quizSheet.getRange('A4').getValue();
  const wrongQuestionsBefore = getWrongQuestions(quizSheet);

  // Answer wrong
  toggleRightWrongCheckboxes(quizSheet, true);
  quizSheet.getRange('B6').setValue(true);
  const wrongEvent = createMockEditEvent(ss, quizSheet.getRange('B6'), true, false);
  handleCheckboxEdit(wrongEvent);

  const wrongQuestionsAfter = getWrongQuestions(quizSheet);

  const passed = (wrongQuestionsBefore.length === 0 &&
    wrongQuestionsAfter.length === 1 &&
    wrongQuestionsAfter.includes(currentQuestion));

  recordTestResult(
    'testWrongAnswerAddsToList',
    'Answering wrong should add question to wrong questions list',
    passed,
    `Question: "${currentQuestion}". Before: ${wrongQuestionsBefore.length} questions. After: ${wrongQuestionsAfter.length} questions. Contains question: ${wrongQuestionsAfter.includes(currentQuestion)}`
  );

  return passed;
}

function testRightAnswerDoesNotAddToList() {
  setupTestDatastore();
  resetQuizSheet();

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const quizSheet = ss.getSheetByName('quiz');

  // Start quiz
  quizSheet.getRange('B2').setValue(true);
  const startEvent = createMockEditEvent(ss, quizSheet.getRange('B2'), true, false);
  handleCheckboxEdit(startEvent);

  const currentQuestion = quizSheet.getRange('A4').getValue();
  const wrongQuestionsBefore = getWrongQuestions(quizSheet);

  // Answer right
  toggleRightWrongCheckboxes(quizSheet, true);
  quizSheet.getRange('B5').setValue(true);
  const rightEvent = createMockEditEvent(ss, quizSheet.getRange('B5'), true, false);
  handleCheckboxEdit(rightEvent);

  const wrongQuestionsAfter = getWrongQuestions(quizSheet);

  const passed = (wrongQuestionsBefore.length === 0 &&
    wrongQuestionsAfter.length === 0);

  recordTestResult(
    'testRightAnswerDoesNotAddToList',
    'Answering right should not add question to wrong questions list',
    passed,
    `Question: "${currentQuestion}". Before: ${wrongQuestionsBefore.length} questions. After: ${wrongQuestionsAfter.length} questions.`
  );

  return passed;
}

function testNoDuplicateWrongQuestions() {
  setupTestDatastore();
  resetQuizSheet();

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const quizSheet = ss.getSheetByName('quiz');

  const testQuestion = "Test Question";

  // Manually set a question and add it to wrong questions multiple times
  quizSheet.getRange('A4').setValue(testQuestion);

  addWrongQuestion(quizSheet, testQuestion);
  addWrongQuestion(quizSheet, testQuestion);
  addWrongQuestion(quizSheet, testQuestion);

  const wrongQuestions = getWrongQuestions(quizSheet);
  const uniqueCount = [...new Set(wrongQuestions)].length;

  const passed = (wrongQuestions.length === 1 && uniqueCount === 1);

  recordTestResult(
    'testNoDuplicateWrongQuestions',
    'Same wrong question should not be added multiple times',
    passed,
    `Added same question 3 times. List length: ${wrongQuestions.length}, Unique count: ${uniqueCount}`
  );

  return passed;
}

// 3. Retry Mode Question Selection Tests
function testRetryModeShowsOnlyWrongQuestions() {
  setupTestDatastore();
  resetQuizSheet();

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const quizSheet = ss.getSheetByName('quiz');

  // Manually add some wrong questions
  setWrongQuestions(quizSheet, ['Question 1', 'Question 3']);

  // Check retry mode and start quiz
  quizSheet.getRange('B3').setValue(true);
  quizSheet.getRange('B2').setValue(true);
  const startEvent = createMockEditEvent(ss, quizSheet.getRange('B2'), true, false);
  handleCheckboxEdit(startEvent);

  const currentQuestion = quizSheet.getRange('A4').getValue();
  const wrongQuestions = getWrongQuestions(quizSheet);

  const passed = wrongQuestions.includes(currentQuestion);

  recordTestResult(
    'testRetryModeShowsOnlyWrongQuestions',
    'Retry mode should only show questions from wrong questions list',
    passed,
    `Current question: "${currentQuestion}". Wrong questions list: [${wrongQuestions.join(', ')}]. Question in list: ${passed}`
  );

  return passed;
}

function testRetryModeEmptyListMessage() {
  setupTestDatastore();
  resetQuizSheet();

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const quizSheet = ss.getSheetByName('quiz');

  // Ensure wrong questions list is empty
  resetWrongQuestions(quizSheet);

  // Check retry mode and start quiz
  quizSheet.getRange('B3').setValue(true);
  quizSheet.getRange('B2').setValue(true);
  const startEvent = createMockEditEvent(ss, quizSheet.getRange('B2'), true, false);
  handleCheckboxEdit(startEvent);

  const currentQuestion = quizSheet.getRange('A4').getValue();
  const expectedMessage = "No wrong questions available to retry for this category.";

  const passed = (currentQuestion === expectedMessage);

  recordTestResult(
    'testRetryModeEmptyListMessage',
    'Retry mode with empty wrong questions list should show appropriate message',
    passed,
    `Expected: "${expectedMessage}". Got: "${currentQuestion}"`
  );

  return passed;
}

// 4. Right Answer in Retry Mode Tests
function testRightAnswerInRetryModeRemovesFromList() {
  setupTestDatastore();
  resetQuizSheet();

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const quizSheet = ss.getSheetByName('quiz');

  // Set up wrong questions list
  setWrongQuestions(quizSheet, ['Question 1', 'Question 2', 'Question 3']);

  // Start retry mode
  quizSheet.getRange('B3').setValue(true);
  quizSheet.getRange('B2').setValue(true);
  const startEvent = createMockEditEvent(ss, quizSheet.getRange('B2'), true, false);
  handleCheckboxEdit(startEvent);

  const currentQuestion = quizSheet.getRange('A4').getValue();
  const wrongQuestionsBefore = getWrongQuestions(quizSheet);

  // Answer right
  toggleRightWrongCheckboxes(quizSheet, true);
  quizSheet.getRange('B5').setValue(true);
  const rightEvent = createMockEditEvent(ss, quizSheet.getRange('B5'), true, false);
  handleCheckboxEdit(rightEvent);

  const wrongQuestionsAfter = getWrongQuestions(quizSheet);

  const passed = (wrongQuestionsBefore.includes(currentQuestion) &&
    !wrongQuestionsAfter.includes(currentQuestion) &&
    wrongQuestionsAfter.length === wrongQuestionsBefore.length - 1);

  recordTestResult(
    'testRightAnswerInRetryModeRemovesFromList',
    'Right answer in retry mode should remove question from wrong questions list',
    passed,
    `Question: "${currentQuestion}". Before: ${wrongQuestionsBefore.length} questions. After: ${wrongQuestionsAfter.length} questions. Removed: ${!wrongQuestionsAfter.includes(currentQuestion)}`
  );

  return passed;
}

function testRightAnswerInRetryModeIncrementsCounter() {
  setupTestDatastore();
  resetQuizSheet();

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const quizSheet = ss.getSheetByName('quiz');

  // Set up wrong questions list
  setWrongQuestions(quizSheet, ['Question 1', 'Question 2']);

  // Start retry mode
  quizSheet.getRange('B3').setValue(true);
  quizSheet.getRange('B2').setValue(true);
  const startEvent = createMockEditEvent(ss, quizSheet.getRange('B2'), true, false);
  handleCheckboxEdit(startEvent);

  const rightCountBefore = getRightAnswersCount(quizSheet);

  // Answer right
  toggleRightWrongCheckboxes(quizSheet, true);
  quizSheet.getRange('B5').setValue(true);
  const rightEvent = createMockEditEvent(ss, quizSheet.getRange('B5'), true, false);
  handleCheckboxEdit(rightEvent);

  const rightCountAfter = getRightAnswersCount(quizSheet);

  const passed = (rightCountAfter === rightCountBefore + 1);

  recordTestResult(
    'testRightAnswerInRetryModeIncrementsCounter',
    'Right answer in retry mode should increment right answer counter',
    passed,
    `Right count before: ${rightCountBefore}. Right count after: ${rightCountAfter}`
  );

  return passed;
}

// 5. Wrong Answer in Retry Mode Tests
function testWrongAnswerInRetryModeKeepsInList() {
  setupTestDatastore();
  resetQuizSheet();

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const quizSheet = ss.getSheetByName('quiz');

  // Set up wrong questions list
  setWrongQuestions(quizSheet, ['Question 1', 'Question 2']);

  // Start retry mode
  quizSheet.getRange('B3').setValue(true);
  quizSheet.getRange('B2').setValue(true);
  const startEvent = createMockEditEvent(ss, quizSheet.getRange('B2'), true, false);
  handleCheckboxEdit(startEvent);

  const currentQuestion = quizSheet.getRange('A4').getValue();
  const wrongQuestionsBefore = getWrongQuestions(quizSheet);

  // Answer wrong
  toggleRightWrongCheckboxes(quizSheet, true);
  quizSheet.getRange('B6').setValue(true);
  const wrongEvent = createMockEditEvent(ss, quizSheet.getRange('B6'), true, false);
  handleCheckboxEdit(wrongEvent);

  const wrongQuestionsAfter = getWrongQuestions(quizSheet);

  const passed = (wrongQuestionsBefore.includes(currentQuestion) &&
    wrongQuestionsAfter.includes(currentQuestion) &&
    wrongQuestionsAfter.length === wrongQuestionsBefore.length);

  recordTestResult(
    'testWrongAnswerInRetryModeKeepsInList',
    'Wrong answer in retry mode should keep question in wrong questions list',
    passed,
    `Question: "${currentQuestion}". Before: ${wrongQuestionsBefore.length} questions. After: ${wrongQuestionsAfter.length} questions. Still in list: ${wrongQuestionsAfter.includes(currentQuestion)}`
  );

  return passed;
}

function testWrongAnswerInRetryModeIncrementsCounter() {
  setupTestDatastore();
  resetQuizSheet();

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const quizSheet = ss.getSheetByName('quiz');

  // Set up wrong questions list
  setWrongQuestions(quizSheet, ['Question 1', 'Question 2']);

  // Start retry mode
  quizSheet.getRange('B3').setValue(true);
  quizSheet.getRange('B2').setValue(true);
  const startEvent = createMockEditEvent(ss, quizSheet.getRange('B2'), true, false);
  handleCheckboxEdit(startEvent);

  const wrongCountBefore = getWrongAnswersCount(quizSheet);

  // Answer wrong
  toggleRightWrongCheckboxes(quizSheet, true);
  quizSheet.getRange('B6').setValue(true);
  const wrongEvent = createMockEditEvent(ss, quizSheet.getRange('B6'), true, false);
  handleCheckboxEdit(wrongEvent);

  const wrongCountAfter = getWrongAnswersCount(quizSheet);

  const passed = (wrongCountAfter === wrongCountBefore + 1);

  recordTestResult(
    'testWrongAnswerInRetryModeIncrementsCounter',
    'Wrong answer in retry mode should increment wrong answer counter',
    passed,
    `Wrong count before: ${wrongCountBefore}. Wrong count after: ${wrongCountAfter}`
  );

  return passed;
}

// 6. Retry Mode Completion Tests
function testRetryModeCompletionMessage() {
  setupTestDatastore();
  resetQuizSheet();

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const quizSheet = ss.getSheetByName('quiz');

  // Set up wrong questions list with only one question
  setWrongQuestions(quizSheet, ['Question 1']);

  // Start retry mode
  quizSheet.getRange('B3').setValue(true);
  quizSheet.getRange('B2').setValue(true);
  const startEvent = createMockEditEvent(ss, quizSheet.getRange('B2'), true, false);
  handleCheckboxEdit(startEvent);

  // Answer the question right (should remove it from wrong questions list)
  toggleRightWrongCheckboxes(quizSheet, true);
  quizSheet.getRange('B5').setValue(true);
  const rightEvent = createMockEditEvent(ss, quizSheet.getRange('B5'), true, false);
  handleCheckboxEdit(rightEvent);

  const finalQuestion = quizSheet.getRange('A4').getValue();
  const expectedMessage = "Quiz Complete! All wrong questions have been answered correctly.";
  const quizStopped = !quizSheet.getRange('B2').getValue();

  const passed = (finalQuestion === expectedMessage && quizStopped);

  recordTestResult(
    'testRetryModeCompletionMessage',
    'Retry mode should show completion message when all wrong questions answered correctly',
    passed,
    `Expected: "${expectedMessage}". Got: "${finalQuestion}". Quiz stopped: ${quizStopped}`
  );

  return passed;
}

function testRetryModeCompletionStopsQuiz() {
  setupTestDatastore();
  resetQuizSheet();

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const quizSheet = ss.getSheetByName('quiz');

  // Set up wrong questions list with only one question
  setWrongQuestions(quizSheet, ['Question 1']);

  // Start retry mode
  quizSheet.getRange('B3').setValue(true);
  quizSheet.getRange('B2').setValue(true);
  const startEvent = createMockEditEvent(ss, quizSheet.getRange('B2'), true, false);
  handleCheckboxEdit(startEvent);

  const quizRunningBefore = quizSheet.getRange('B2').getValue();

  // Answer the question right (should complete the quiz)
  toggleRightWrongCheckboxes(quizSheet, true);
  quizSheet.getRange('B5').setValue(true);
  const rightEvent = createMockEditEvent(ss, quizSheet.getRange('B5'), true, false);
  handleCheckboxEdit(rightEvent);

  const quizRunningAfter = quizSheet.getRange('B2').getValue();
  const checkboxesDisabled = isRangeProtected(quizSheet, 'B5') && isRangeProtected(quizSheet, 'B6');

  const passed = (quizRunningBefore === true && quizRunningAfter === false && checkboxesDisabled);

  recordTestResult(
    'testRetryModeCompletionStopsQuiz',
    'Retry mode completion should stop quiz and disable checkboxes',
    passed,
    `Quiz before: ${quizRunningBefore}. Quiz after: ${quizRunningAfter}. Checkboxes disabled: ${checkboxesDisabled}`
  );

  return passed;
}

/**
 * Run all unit tests and record results in the test results sheet
 */
function runAllTestsGit() {
  clearTestResults();


  // Score Increment Tests

  testWrongQuestionsListPersistence();
  testRetryModeStopsRunningQuiz();
  testRetryModeToggleWhenQuizNotRunning();
  testWrongAnswerAddsToList();
  testRightAnswerDoesNotAddToList();
  testNoDuplicateWrongQuestions();
  testRetryModeShowsOnlyWrongQuestions();
  testRetryModeEmptyListMessage();
  testRightAnswerInRetryModeRemovesFromList();
  testRightAnswerInRetryModeIncrementsCounter();
  testWrongAnswerInRetryModeKeepsInList();
  testWrongAnswerInRetryModeIncrementsCounter();
  testRetryModeCompletionMessage();
  testRetryModeCompletionStopsQuiz();
  testWrongQuestionsResetInNormalMode();

  

 

}