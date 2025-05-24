function setupQuizSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const datastoreSheet = ss.getSheetByName('datastore');
  const quizSheet = ss.getSheetByName('quiz');

  // === 1. Set Category Dropdown in A1 ===
  const data = datastoreSheet.getRange('B2:B' + datastoreSheet.getLastRow()).getValues();
  const uniqueCategories = [...new Set(data.flat().filter(Boolean))];

  const dropdownRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(uniqueCategories, true)
    .setAllowInvalid(false)
    .build();

  quizSheet.getRange('A1').setDataValidation(dropdownRule);

  // === 2. Insert "Start Quiz" label and checkbox ===
  quizSheet.getRange('A2').setValue('Start Quiz');
  quizSheet.getRange('B2').insertCheckboxes();

  // === 3. Add Right and Wrong checkboxes ===
  quizSheet.getRange('A5').setValue('Right');
  quizSheet.getRange('B5').insertCheckboxes();
  quizSheet.getRange('A6').setValue('Wrong');
  quizSheet.getRange('B6').insertCheckboxes();

  // === 4. Add Show Answer checkbox ===
  quizSheet.getRange('A7').setValue('Show Answer');
  quizSheet.getRange('B7').insertCheckboxes();

  // === 5. Add Score Display Section ===
  quizSheet.getRange('A9').setValue('Right Answers:');
  quizSheet.getRange('B9').setValue(0);
  quizSheet.getRange('A10').setValue('Wrong Answers:');
  quizSheet.getRange('B10').setValue(0);

  // === 6. Initially disable Right, Wrong, and Show Answer checkboxes ===
  quizSheet.getRange('B5').protect().setDescription('Right checkbox - disabled until quiz starts');
  quizSheet.getRange('B6').protect().setDescription('Wrong checkbox - disabled until quiz starts');
  quizSheet.getRange('B7').protect().setDescription('Show Answer checkbox - disabled until quiz starts');

  // === 7. Add onEdit trigger (only if not already set) ===
  const triggers = ScriptApp.getProjectTriggers();
  const hasOnEdit = triggers.some(trigger => trigger.getHandlerFunction() === 'handleCheckboxEdit');

  if (!hasOnEdit) {
    ScriptApp.newTrigger('handleCheckboxEdit')
      .forSpreadsheet(ss)
      .onEdit()
      .create();
  }
}

// === Helper function to enable/disable Right, Wrong, and Show Answer checkboxes ===
function toggleRightWrongCheckboxes(sheet, enable) {
  const rightCheckbox = sheet.getRange('B5');
  const wrongCheckbox = sheet.getRange('B6');
  const showAnswerCheckbox = sheet.getRange('B7');
  
  if (enable) {
    // Remove protection to enable checkboxes
    const protections = sheet.getProtections(SpreadsheetApp.ProtectionType.RANGE);
    protections.forEach(protection => {
      const range = protection.getRange();
      if (range.getA1Notation() === 'B5' || range.getA1Notation() === 'B6' || range.getA1Notation() === 'B7') {
        protection.remove();
      }
    });
    // Don't clear checkbox values when enabling - preserve existing state
  } else {
    // Add protection to disable checkboxes
    rightCheckbox.protect().setDescription('Right checkbox - disabled until quiz starts');
    wrongCheckbox.protect().setDescription('Wrong checkbox - disabled until quiz starts');
    showAnswerCheckbox.protect().setDescription('Show Answer checkbox - disabled until quiz starts');
    // Clear checkbox values when disabling
    rightCheckbox.setValue(false);
    wrongCheckbox.setValue(false);
    showAnswerCheckbox.setValue(false);
  }
}

// === Helper function to get or initialize question counter ===
function getQuestionCounter(sheet) {
  const counterCell = sheet.getRange('C1');
  const counter = counterCell.getValue();
  return (typeof counter === 'number' && counter >= 0) ? counter : 0;
}

// === Helper function to set question counter ===
function setQuestionCounter(sheet, count) {
  sheet.getRange('C1').setValue(count);
}

// === Helper function to reset question counter ===
function resetQuestionCounter(sheet) {
  setQuestionCounter(sheet, 0);
}

// === Helper function to get used questions list ===
function getUsedQuestions(sheet) {
  const usedQuestionsCell = sheet.getRange('D1');
  const usedQuestionsStr = usedQuestionsCell.getValue();
  if (!usedQuestionsStr || usedQuestionsStr === '') {
    return [];
  }
  try {
    return JSON.parse(usedQuestionsStr);
  } catch (e) {
    return [];
  }
}

// === Helper function to set used questions list ===
function setUsedQuestions(sheet, usedQuestions) {
  const usedQuestionsCell = sheet.getRange('D1');
  usedQuestionsCell.setValue(JSON.stringify(usedQuestions));
}

// === Helper function to reset used questions list ===
function resetUsedQuestions(sheet) {
  setUsedQuestions(sheet, []);
}

// === Helper function to add a question to used questions list ===
function addUsedQuestion(sheet, questionText) {
  const usedQuestions = getUsedQuestions(sheet);
  if (!usedQuestions.includes(questionText)) {
    usedQuestions.push(questionText);
    setUsedQuestions(sheet, usedQuestions);
  }
}

// === Helper function to get right answers count ===
function getRightAnswersCount(sheet) {
  const rightCountCell = sheet.getRange('B9');
  const count = rightCountCell.getValue();
  return (typeof count === 'number' && count >= 0) ? count : 0;
}

// === Helper function to set right answers count ===
function setRightAnswersCount(sheet, count) {
  sheet.getRange('B9').setValue(count);
}

// === Helper function to get wrong answers count ===
function getWrongAnswersCount(sheet) {
  const wrongCountCell = sheet.getRange('B10');
  const count = wrongCountCell.getValue();
  return (typeof count === 'number' && count >= 0) ? count : 0;
}

// === Helper function to set wrong answers count ===
function setWrongAnswersCount(sheet, count) {
  sheet.getRange('B10').setValue(count);
}

// === Helper function to increment right answers count ===
function incrementRightAnswers(sheet) {
  const currentCount = getRightAnswersCount(sheet);
  setRightAnswersCount(sheet, currentCount + 1);
}

// === Helper function to increment wrong answers count ===
function incrementWrongAnswers(sheet) {
  const currentCount = getWrongAnswersCount(sheet);
  setWrongAnswersCount(sheet, currentCount + 1);
}

// === Helper function to reset answer counts ===
function resetAnswerCounts(sheet) {
  setRightAnswersCount(sheet, 0);
  setWrongAnswersCount(sheet, 0);
}

// === Helper function to get current question's answer ===
function getCurrentQuestionAnswer(sheet, spreadsheet) {
  const currentQuestion = sheet.getRange('A4').getValue();
  if (!currentQuestion || currentQuestion === '') {
    return '';
  }
  
  const datastore = spreadsheet.getSheetByName('datastore');
  const data = datastore.getDataRange().getValues();
  
  // Find the row with the current question
  const questionRow = data.find((row, index) => index !== 0 && row[2] === currentQuestion);
  
  if (questionRow) {
    return questionRow[3]; // Answer is in column D (index 3)
  }
  
  return '';
}

// === Helper function to show/hide answer ===
function toggleAnswer(sheet, spreadsheet, show) {
  const answerCell = sheet.getRange('A8');
  
  if (show) {
    const answer = getCurrentQuestionAnswer(sheet, spreadsheet);
    answerCell.setValue(answer);
  } else {
    answerCell.setValue('');
  }
}

// === Respond to checkbox and category edits ===
function handleCheckboxEdit(e) {
  if (!e) return;

  const sheet = e.source.getSheetByName('quiz');
  const range = e.range;
  const cell = range.getA1Notation();

  // === If category changed in A1 ===
  if (sheet.getName() === 'quiz' && cell === 'A1') {
    sheet.getRange('A4').setValue('');     // Clear question
    sheet.getRange('A8').setValue('');     // Clear answer
    sheet.getRange('B2').setValue(false);  // Uncheck Start Quiz checkbox
    resetQuestionCounter(sheet);           // Reset question counter
    resetUsedQuestions(sheet);             // Reset used questions list
    resetAnswerCounts(sheet);              // Reset answer counts
    toggleRightWrongCheckboxes(sheet, false); // Disable Right/Wrong/Show Answer checkboxes
    return;
  }

  // === If Start Quiz checkbox in B2 is checked ===
  if (sheet.getName() === 'quiz' && cell === 'B2') {
    const isChecked = range.getValue();
    const category = sheet.getRange('A1').getValue();
    const questionCell = sheet.getRange('A4');
    const answerCell = sheet.getRange('A8');

    if (isChecked && category) {
      const datastore = e.source.getSheetByName('datastore');
      const data = datastore.getDataRange().getValues();
      const filtered = data.filter((row, index) => index !== 0 && row[1] === category);

      if (filtered.length > 0) {
        // Reset used questions and answer counts when starting a new quiz
        resetUsedQuestions(sheet);
        resetAnswerCounts(sheet);
        
        const usedQuestions = getUsedQuestions(sheet);
        const availableQuestions = filtered.filter(row => !usedQuestions.includes(row[2]));
        
        if (availableQuestions.length > 0) {
          const random = availableQuestions[Math.floor(Math.random() * availableQuestions.length)];
          questionCell.setValue(random[2]); // Question in column C
          addUsedQuestion(sheet, random[2]); // Track this question as used
          setQuestionCounter(sheet, 1); // Start with question 1
          // Enable Right/Wrong/Show Answer checkboxes when quiz starts
          toggleRightWrongCheckboxes(sheet, true);
          
          // Check if Show Answer checkbox is already checked and display answer if so
          const showAnswerChecked = sheet.getRange('B7').getValue();
          if (showAnswerChecked) {
            toggleAnswer(sheet, e.source, true);
          }
        } else {
          questionCell.setValue("No more questions available for this category.");
          resetQuestionCounter(sheet);
          resetUsedQuestions(sheet);
          // Disable checkboxes if no questions available
          toggleRightWrongCheckboxes(sheet, false);
        }
      } else {
        questionCell.setValue("No questions available for this category.");
        resetQuestionCounter(sheet);
        resetUsedQuestions(sheet);
        // Disable checkboxes if no questions available
        toggleRightWrongCheckboxes(sheet, false);
      }
    } else {
      questionCell.setValue('');
      answerCell.setValue(''); // Clear answer when quiz stops
      resetQuestionCounter(sheet);
      resetUsedQuestions(sheet);
      resetAnswerCounts(sheet); // Reset answer counts when quiz stops
      // Disable Right/Wrong/Show Answer checkboxes when quiz stops
      toggleRightWrongCheckboxes(sheet, false);
    }
  }

  // === If Right checkbox in B5 is checked ===
  if (sheet.getName() === 'quiz' && cell === 'B5') {
    const isChecked = range.getValue();
    const startQuizChecked = sheet.getRange('B2').getValue();
    
    // Only proceed if Start Quiz is checked
    if (isChecked && startQuizChecked) {
      // Increment right answers count
      incrementRightAnswers(sheet);
      // Uncheck Wrong checkbox
      sheet.getRange('B6').setValue(false);
      // Show next question
      showNextQuestion(sheet, e.source);
    } else if (isChecked && !startQuizChecked) {
      // If Start Quiz is not checked, uncheck this checkbox
      sheet.getRange('B5').setValue(false);
    }
  }

  // === If Wrong checkbox in B6 is checked ===
  if (sheet.getName() === 'quiz' && cell === 'B6') {
    const isChecked = range.getValue();
    const startQuizChecked = sheet.getRange('B2').getValue();
    
    // Only proceed if Start Quiz is checked
    if (isChecked && startQuizChecked) {
      // Increment wrong answers count
      incrementWrongAnswers(sheet);
      // Uncheck Right checkbox
      sheet.getRange('B5').setValue(false);
      // Show next question
      showNextQuestion(sheet, e.source);
    } else if (isChecked && !startQuizChecked) {
      // If Start Quiz is not checked, uncheck this checkbox
      sheet.getRange('B6').setValue(false);
    }
  }

  // === If Show Answer checkbox in B7 is checked/unchecked ===
  if (sheet.getName() === 'quiz' && cell === 'B7') {
    const isChecked = range.getValue();
    const startQuizChecked = sheet.getRange('B2').getValue();
    
    // Only proceed if Start Quiz is checked
    if (startQuizChecked) {
      // Show or hide answer based on checkbox state
      toggleAnswer(sheet, e.source, isChecked);
    } else if (isChecked) {
      // If Start Quiz is not checked, uncheck this checkbox
      sheet.getRange('B7').setValue(false);
    }
  }
}

// === Helper function to show next random question ===
function showNextQuestion(sheet, spreadsheet) {
  const category = sheet.getRange('A1').getValue();
  const questionCell = sheet.getRange('A4');
  const answerCell = sheet.getRange('A8');
  const showAnswerCheckbox = sheet.getRange('B7');
  const currentCount = getQuestionCounter(sheet);

  // Check if we've reached 5 questions
  if (currentCount >= 5) {
    questionCell.setValue("Quiz Complete! You have answered 5 questions.");
    answerCell.setValue(''); // Clear answer when quiz completes
    // Disable checkboxes and stop quiz
    toggleRightWrongCheckboxes(sheet, false);
    sheet.getRange('B2').setValue(false); // Uncheck Start Quiz
    resetQuestionCounter(sheet);
    resetUsedQuestions(sheet); // Reset used questions when quiz completes
    return;
  }

  if (category) {
    const datastore = spreadsheet.getSheetByName('datastore');
    const data = datastore.getDataRange().getValues();
    const filtered = data.filter((row, index) => index !== 0 && row[1] === category);

    if (filtered.length > 0) {
      const usedQuestions = getUsedQuestions(sheet);
      const availableQuestions = filtered.filter(row => !usedQuestions.includes(row[2]));
      
      if (availableQuestions.length > 0) {
        const random = availableQuestions[Math.floor(Math.random() * availableQuestions.length)];
        questionCell.setValue(random[2]); // Question in column C
        addUsedQuestion(sheet, random[2]); // Track this question as used
        
        // Increment question counter
        setQuestionCounter(sheet, currentCount + 1);
        
        // Clear both Right/Wrong checkboxes for the next question
        sheet.getRange('B5').setValue(false);
        sheet.getRange('B6').setValue(false);
        
        // Handle Show Answer checkbox - if checked, show new answer; if unchecked, clear answer
        const showAnswerChecked = showAnswerCheckbox.getValue();
        if (showAnswerChecked) {
          toggleAnswer(sheet, spreadsheet, true);
        } else {
          answerCell.setValue('');
        }
      } else {
        // No more unique questions available, end quiz early
        questionCell.setValue("Quiz Complete! No more unique questions available for this category.");
        answerCell.setValue(''); // Clear answer when quiz completes
        // Disable checkboxes and stop quiz
        toggleRightWrongCheckboxes(sheet, false);
        sheet.getRange('B2').setValue(false); // Uncheck Start Quiz
        resetQuestionCounter(sheet);
        resetUsedQuestions(sheet);
      }
    } else {
      questionCell.setValue("No questions available for this category.");
      answerCell.setValue(''); // Clear answer
      resetQuestionCounter(sheet);
      resetUsedQuestions(sheet);
    }
  }
}