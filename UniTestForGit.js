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
    'Show Answer checkbox should be disabled when quiz is not started',
    isShowAnswerProtected,
    isShowAnswerProtected ? 
      '✓ Show Answer checkbox is properly disabled when quiz is not started' :
      '✗ Show Answer checkbox is not disabled when quiz is not started'
  );
  
  return isShowAnswerProtected;
}

// Test that Show Answer checkbox is enabled when quiz is started
function testShowAnswerCheckboxEnabledWhenQuizStarted() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const quizSheet = ss.getSheetByName('quiz');
  const datastoreSheet = ss.getSheetByName('datastore');
  
  // Get a valid category
  const data = datastoreSheet.getDataRange().getValues();
  const validCategories = [...new Set(data.slice(1).map(row => row[1]).filter(Boolean))];
  const testCategory = validCategories[0];
  
  // Setup quiz
  quizSheet.getRange('A1').setValue(testCategory);
  quizSheet.getRange('B2').setValue(false);
  quizSheet.getRange('B7').setValue(false);
  
  // Start quiz
  const startQuizEvent = {
    source: ss,
    range: quizSheet.getRange('B2'),
    value: true,
    oldValue: false
  };
  handleCheckboxEdit(startQuizEvent);
  
  // Check if Show Answer checkbox is not protected (enabled)
  const isShowAnswerProtected = isRangeProtected(quizSheet, 'B7');
  
  recordTestResult(
    'Show Answer checkbox should be enabled when quiz is started',
    !isShowAnswerProtected,
    !isShowAnswerProtected ? 
      '✓ Show Answer checkbox is properly enabled when quiz is started' :
      '✗ Show Answer checkbox is still disabled when quiz is started'
  );
  
  return !isShowAnswerProtected;
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
  
  // Get a valid category
  const data = datastoreSheet.getDataRange().getValues();
  const validCategories = [...new Set(data.slice(1).map(row => row[1]).filter(Boolean))];
  const testCategory = validCategories[0];
  
  // Setup quiz
  quizSheet.getRange('A1').setValue(testCategory);
  quizSheet.getRange('B2').setValue(false);
  quizSheet.getRange('B7').setValue(false);
  
  // Start quiz
  const startQuizEvent = {
    source: ss,
    range: quizSheet.getRange('B2'),
    value: true,
    oldValue: false
  };
  handleCheckboxEdit(startQuizEvent);
  
  // Check Show Answer checkbox first
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
    'Show Answer checkbox should hide the answer when unchecked',
    testPassed,
    testPassed ? 
      '✓ Show Answer checkbox correctly hides the answer when unchecked' :
      `✗ Show Answer hide failed. Answer when checked: "${answerWhenChecked}", Answer when unchecked: "${answerWhenUnchecked}"`
  );
  
  return testPassed;
}

// Test that Show Answer updates with new question when moving to next question
function testShowAnswerUpdatesWithNextQuestion() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const quizSheet = ss.getSheetByName('quiz');
  const datastoreSheet = ss.getSheetByName('datastore');
  
  // Get a category with multiple questions
  const data = datastoreSheet.getDataRange().getValues();
  const categoryQuestionCounts = {};
  
  data.slice(1).forEach(row => {
    if (row[1]) {
      categoryQuestionCounts[row[1]] = (categoryQuestionCounts[row[1]] || 0) + 1;
    }
  });
  
  let testCategory = null;
  for (const [category, count] of Object.entries(categoryQuestionCounts)) {
    if (count >= 2) {
      testCategory = category;
      break;
    }
  }
  
  if (!testCategory) {
    recordTestResult(
      'Show Answer should update with new question when moving to next question',
      false,
      '✗ No category with multiple questions found for testing'
    );
    return false;
  }
  
  // Setup quiz
  quizSheet.getRange('A1').setValue(testCategory);
  quizSheet.getRange('B2').setValue(false);
  quizSheet.getRange('B7').setValue(false);
  quizSheet.getRange('C1').setValue(0); // Reset counter
  quizSheet.getRange('D1').setValue(''); // Clear used questions
  
  // Start quiz
  const startQuizEvent = {
    source: ss,
    range: quizSheet.getRange('B2'),
    value: true,
    oldValue: false
  };
  handleCheckboxEdit(startQuizEvent);
  
  // Check Show Answer checkbox
  const showAnswerEvent = {
    source: ss,
    range: quizSheet.getRange('B7'),
    value: true,
    oldValue: false
  };
  handleCheckboxEdit(showAnswerEvent);
  
  // Get first question and answer
  const firstQuestion = quizSheet.getRange('A4').getValue();
  const firstAnswer = quizSheet.getRange('A8').getValue();
  
  // Move to next question by clicking Right
  const rightClickEvent = {
    source: ss,
    range: quizSheet.getRange('B5'),
    value: true,
    oldValue: false
  };
  handleCheckboxEdit(rightClickEvent);
  
  // Get second question and answer
  const secondQuestion = quizSheet.getRange('A4').getValue();
  const secondAnswer = quizSheet.getRange('A8').getValue();
  
  const testPassed = firstQuestion !== secondQuestion && 
                     firstAnswer !== secondAnswer && 
                     firstAnswer !== '' && 
                     secondAnswer !== '' &&
                     !secondQuestion.includes('Quiz Complete');
  
  recordTestResult(
    'Show Answer should update with new question when moving to next question',
    testPassed,
    testPassed ? 
      '✓ Show Answer correctly updates with new question' :
      `✗ Show Answer update failed. First Q: "${firstQuestion}", Second Q: "${secondQuestion}", First A: "${firstAnswer}", Second A: "${secondAnswer}"`
  );
  
  return testPassed;
}

// Test that Show Answer is cleared when quiz ends
function testShowAnswerClearedWhenQuizEnds() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const quizSheet = ss.getSheetByName('quiz');
  const datastoreSheet = ss.getSheetByName('datastore');
  
  // Get a valid category
  const data = datastoreSheet.getDataRange().getValues();
  const validCategories = [...new Set(data.slice(1).map(row => row[1]).filter(Boolean))];
  const testCategory = validCategories[0];
  
  // Setup quiz
  quizSheet.getRange('A1').setValue(testCategory);
  quizSheet.getRange('B2').setValue(false);
  quizSheet.getRange('B7').setValue(false);
  quizSheet.getRange('C1').setValue(0); // Reset counter
  quizSheet.getRange('D1').setValue(''); // Clear used questions
  
  // Start quiz
  const startQuizEvent = {
    source: ss,
    range: quizSheet.getRange('B2'),
    value: true,
    oldValue: false
  };
  handleCheckboxEdit(startQuizEvent);
  
  // Check Show Answer checkbox
  const showAnswerEvent = {
    source: ss,
    range: quizSheet.getRange('B7'),
    value: true,
    oldValue: false
  };
  handleCheckboxEdit(showAnswerEvent);
  
  // Verify answer is displayed
  const answerBeforeStop = quizSheet.getRange('A8').getValue();
  
  // Stop quiz
  const stopQuizEvent = {
    source: ss,
    range: quizSheet.getRange('B2'),
    value: false,
    oldValue: true
  };
  handleCheckboxEdit(stopQuizEvent);
  
  // Check answer after stopping quiz
  const answerAfterStop = quizSheet.getRange('A8').getValue();
  
  const testPassed = answerBeforeStop !== '' && answerAfterStop === '';
  
  recordTestResult(
    'Show Answer should be cleared when quiz ends',
    testPassed,
    testPassed ? 
      '✓ Show Answer is correctly cleared when quiz ends' :
      `✗ Show Answer clear failed. Before stop: "${answerBeforeStop}", After stop: "${answerAfterStop}"`
  );
  
  return testPassed;
}

/**
 * Run all unit tests and record results in the test results sheet
 */
function runAllTestsGit() {
  clearTestResults();

  // Existing tests
  testQuizEndsEarlyWhenNoMoreUniqueQuestions();
  
  // New Show Answer tests
  testShowAnswerCheckboxDisabledWhenQuizNotStarted();
  testShowAnswerCheckboxEnabledWhenQuizStarted();
  testShowAnswerDisplaysCorrectAnswer();
  testShowAnswerHidesAnswerWhenUnchecked();
  testShowAnswerUpdatesWithNextQuestion();
  testShowAnswerClearedWhenQuizEnds();
}