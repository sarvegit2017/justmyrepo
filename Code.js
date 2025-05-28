function setupQuizSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const datastoreSheet = ss.getSheetByName('datastore');
  const quizSheet = ss.getSheetByName('quiz');
  // === NEW: Set up the tracker sheet for wrong answers ===
  setupWrongAnswersTrackerSheet();
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
  // === 3. Add Retry Wrong Questions checkbox ===
  quizSheet.getRange('A3').setValue('Retry Wrong Questions');
  quizSheet.getRange('B3').insertCheckboxes();
  // === 4. Add Right and Wrong checkboxes ===
  quizSheet.getRange('A5').setValue('Right');
  quizSheet.getRange('B5').insertCheckboxes();
  quizSheet.getRange('A6').setValue('Wrong');
  quizSheet.getRange('B6').insertCheckboxes();
  // === 5. Add Show Answer checkbox ===
  quizSheet.getRange('A7').setValue('Show Answer');
  quizSheet.getRange('B7').insertCheckboxes();
  // === 6. Add Score Display Section ===
  quizSheet.getRange('A9').setValue('Right Answers:');
  quizSheet.getRange('B9').setValue(0);
  quizSheet.getRange('A10').setValue('Wrong Answers:');
  quizSheet.getRange('B10').setValue(0);
  // === 7. Initially disable Right, Wrong, Show Answer, and Retry Wrong Questions checkboxes ===
  quizSheet.getRange('B3').protect().setDescription('Retry Wrong Questions checkbox - disabled until quiz starts');
  quizSheet.getRange('B5').protect().setDescription('Right checkbox - disabled until quiz starts');
  quizSheet.getRange('B6').protect().setDescription('Wrong checkbox - disabled until quiz starts');
  quizSheet.getRange('B7').protect().setDescription('Show Answer checkbox - disabled until quiz starts');

  // === 8. Add onEdit trigger (only if not already set) ===
  const triggers = ScriptApp.getProjectTriggers();
  const hasOnEdit = triggers.some(trigger => trigger.getHandlerFunction() === 'handleCheckboxEdit');

  if (!hasOnEdit) {
    ScriptApp.newTrigger('handleCheckboxEdit')
      .forSpreadsheet(ss)
      .onEdit()
      .create();
  }
}

// === NEW: Helper function to set up the WrongAnswersTracker sheet ===
function setupWrongAnswersTrackerSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let trackerSheet = ss.getSheetByName('WrongAnswersTracker');
  if (!trackerSheet) {
    trackerSheet = ss.insertSheet('WrongAnswersTracker');
    trackerSheet.getRange('A1:C1').setValues([['Question', 'Category', 'Wrong Count']]).setFontWeight('bold');
    trackerSheet.setColumnWidth(1, 400);
    // Set width for Question column
    trackerSheet.setColumnWidth(2, 150);
    // Set width for Category column
    trackerSheet.setColumnWidth(3, 100);
    // Set width for Wrong Count column
  }
}

// === NEW: Helper function to update the wrong answers tracker ===
function updateWrongAnswersTracker(questionText, category) {
  if (!questionText || !category) return;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const trackerSheet = ss.getSheetByName('WrongAnswersTracker');
  const data = trackerSheet.getDataRange().getValues();
  // Start searching from row 2 (index 1) to skip header
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === questionText) { // Question found in column A
      const currentCount = data[i][2] || 0; // Get current count from column C
      trackerSheet.getRange(i + 1, 3).setValue(currentCount + 1);
      // Increment count
      return; // Exit after updating
    }
  }

  // If question is not found, append it as a new row
  trackerSheet.appendRow([questionText, category, 1]);
}

// === NEW: Helper function to decrement the wrong answers tracker ===
function decrementWrongAnswersTracker(questionText) {
  if (!questionText) return;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const trackerSheet = ss.getSheetByName('WrongAnswersTracker');
  if (!trackerSheet) return;

  const data = trackerSheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === questionText) {
      const currentCount = data[i][2] || 0;
      if (currentCount > 1) {
        trackerSheet.getRange(i + 1, 3).setValue(currentCount - 1);
      } else if (currentCount === 1) {
        trackerSheet.deleteRow(i + 1); // Delete the row if count becomes 0
      }
      return;
    }
  }
}


// === Helper function to enable/disable Right, Wrong, Show Answer, and Retry Wrong Questions checkboxes ===
function toggleRightWrongCheckboxes(sheet, enable) {
  const rightCheckbox = sheet.getRange('B5');
  const wrongCheckbox = sheet.getRange('B6');
  const showAnswerCheckbox = sheet.getRange('B7');
  const retryWrongCheckbox = sheet.getRange('B3');
  if (enable) {
    // Remove protection to enable checkboxes
    const protections = sheet.getProtections(SpreadsheetApp.ProtectionType.RANGE);
    protections.forEach(protection => {
      const range = protection.getA1Notation();
      if (range === 'B5' || range === 'B6' || range === 'B7' || range === 'B3') {
        protection.remove();
      }
    });
    // Don't clear checkbox values when enabling - preserve existing state
  } else {
    // Add protection to disable checkboxes
    rightCheckbox.protect().setDescription('Right checkbox - disabled until quiz starts');
    wrongCheckbox.protect().setDescription('Wrong checkbox - disabled until quiz starts');
    showAnswerCheckbox.protect().setDescription('Show Answer checkbox - disabled until quiz starts');
    retryWrongCheckbox.protect().setDescription('Retry Wrong Questions checkbox - disabled until quiz starts');
    // Clear checkbox values when disabling
    rightCheckbox.setValue(false);
    wrongCheckbox.setValue(false);
    showAnswerCheckbox.setValue(false);
    retryWrongCheckbox.setValue(false);
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

// === Helper function to get wrong questions list ===
function getWrongQuestions(sheet) {
  const wrongQuestionsCell = sheet.getRange('E1');
  const wrongQuestionsStr = wrongQuestionsCell.getValue();
  if (!wrongQuestionsStr || wrongQuestionsStr === '') {
    return [];
  }
  try {
    return JSON.parse(wrongQuestionsStr);
  } catch (e) {
    return [];
  }
}

// === Helper function to set wrong questions list ===
function setWrongQuestions(sheet, wrongQuestions) {
  const wrongQuestionsCell = sheet.getRange('E1');
  wrongQuestionsCell.setValue(JSON.stringify(wrongQuestions));
}

// === Helper function to reset wrong questions list ===
function resetWrongQuestions(sheet) {
  setWrongQuestions(sheet, []);
}

// === Helper function to add a question to wrong questions list ===
function addWrongQuestion(sheet, questionText) {
  const wrongQuestions = getWrongQuestions(sheet);
  if (!wrongQuestions.includes(questionText)) {
    wrongQuestions.push(questionText);
    setWrongQuestions(sheet, wrongQuestions);
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
    sheet.getRange('A4').setValue('');
    // Clear question
    sheet.getRange('A8').setValue('');     // Clear answer
    sheet.getRange('B2').setValue(false);
    // Uncheck Start Quiz checkbox
    resetQuestionCounter(sheet);           // Reset question counter
    resetUsedQuestions(sheet);
    // Reset used questions list
    resetWrongQuestions(sheet);            // Reset wrong questions list
    resetAnswerCounts(sheet);
    // Reset answer counts
    toggleRightWrongCheckboxes(sheet, false); // Disable Right/Wrong/Show Answer/Retry Wrong Questions checkboxes
    return;
  }

  // === If Start Quiz checkbox in B2 is checked ===
  if (sheet.getName() === 'quiz' && cell === 'B2') {
    const isChecked = range.getValue();
    const category = sheet.getRange('A1').getValue();
    const questionCell = sheet.getRange('A4');
    const answerCell = sheet.getRange('A8');
    const retryWrongChecked = sheet.getRange('B3').getValue();
    if (isChecked && category) {
      const datastore = e.source.getSheetByName('datastore');
      const data = datastore.getDataRange().getValues();
      const filtered = data.filter((row, index) => index !== 0 && row[1] === category);
      if (filtered.length > 0) {
        let questionsToUse = [];
        if (retryWrongChecked) {
          // Get wrong questions for retry mode
          const wrongQuestions = getWrongQuestions(sheet);
          questionsToUse = filtered.filter(row => wrongQuestions.includes(row[2]));
          
          if (questionsToUse.length === 0) {
            questionCell.setValue("No wrong questions available to retry for this category.");
            resetQuestionCounter(sheet);
            toggleRightWrongCheckboxes(sheet, false);
            return;
          }
          
          // In retry mode, don't reset wrong questions - we want to keep the list
        } else {
          // Normal mode - reset wrong questions list when starting a new quiz
          resetWrongQuestions(sheet);
          questionsToUse = filtered;
        }
        
        // Reset used questions and answer counts when starting a quiz
        resetUsedQuestions(sheet);
        resetAnswerCounts(sheet);
        
        const usedQuestions = getUsedQuestions(sheet);
        const availableQuestions = questionsToUse.filter(row => !usedQuestions.includes(row[2]));
        if (availableQuestions.length > 0) {
          const random = availableQuestions[Math.floor(Math.random() * availableQuestions.length)];
          questionCell.setValue(random[2]); // Question in column C
          addUsedQuestion(sheet, random[2]);
          // Track this question as used
          setQuestionCounter(sheet, 1);
          // Start with question 1
          // Enable Right/Wrong/Show Answer checkboxes when quiz starts
          toggleRightWrongCheckboxes(sheet, true);
          // Check if Show Answer checkbox is already checked and display answer if so
          const showAnswerChecked = sheet.getRange('B7').getValue();
          if (showAnswerChecked) {
            toggleAnswer(sheet, e.source, true);
          }
        } else {
          if (retryWrongChecked) {
            questionCell.setValue("No more wrong questions available to retry for this category.");
          } else {
            questionCell.setValue("No more questions available for this category.");
          }
          resetQuestionCounter(sheet);
          resetUsedQuestions(sheet);
          toggleRightWrongCheckboxes(sheet, false);
        }
      } else {
        questionCell.setValue("No questions available for this category.");
        resetQuestionCounter(sheet);
        resetUsedQuestions(sheet);
        toggleRightWrongCheckboxes(sheet, false);
      }
    } else {
      questionCell.setValue('');
      answerCell.setValue('');
      // Clear answer when quiz stops
      resetQuestionCounter(sheet);
      resetUsedQuestions(sheet);
      resetAnswerCounts(sheet);
      // Reset answer counts when quiz stops
      // Disable Right/Wrong/Show Answer/Retry Wrong Questions checkboxes when quiz stops
      toggleRightWrongCheckboxes(sheet, false);
    }
  }

  // === If Retry Wrong Questions checkbox in B3 is checked/unchecked ===
  if (sheet.getName() === 'quiz' && cell === 'B3') {
    const isChecked = range.getValue();
    const startQuizChecked = sheet.getRange('B2').getValue();
    
    // If Start Quiz is running, stop it and clear the question
    if (startQuizChecked) {
      sheet.getRange('B2').setValue(false);
      // Uncheck Start Quiz
      sheet.getRange('A4').setValue(''); // Clear question
      sheet.getRange('A8').setValue('');
      // Clear answer
      resetQuestionCounter(sheet);
      resetUsedQuestions(sheet);
      resetAnswerCounts(sheet);
      toggleRightWrongCheckboxes(sheet, false);
    }
    // Allow the checkbox to stay checked/unchecked as per user's selection
  }

  // === If Right checkbox in B5 is checked ===
  if (sheet.getName() === 'quiz' && cell === 'B5') {
    const isChecked = range.getValue();
    const startQuizChecked = sheet.getRange('B2').getValue();
    
    // Only proceed if Start Quiz is checked
    if (isChecked && startQuizChecked) {
      const currentQuestion = sheet.getRange('A4').getValue();
      // --- MODIFICATION START ---
      // When a question is answered correctly, decrement its count in the tracker.
      if (currentQuestion && currentQuestion !== '') {
        decrementWrongAnswersTracker(currentQuestion);
      }
      // --- MODIFICATION END ---
      
      // If in retry mode, remove current question from wrong questions list
      const retryWrongChecked = sheet.getRange('B3').getValue();
      if (retryWrongChecked) {
        if (currentQuestion && currentQuestion !== '') {
          removeFromWrongQuestions(sheet, currentQuestion);
        }
      }
      
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
      // Add current question to wrong questions list for session
      const currentQuestion = sheet.getRange('A4').getValue();
      if (currentQuestion && currentQuestion !== '') {
        addWrongQuestion(sheet, currentQuestion);
        // --- MODIFICATION START ---
        // Track the wrong answer in the WrongAnswersTracker sheet
        const category = sheet.getRange('A1').getValue();
        updateWrongAnswersTracker(currentQuestion, category);
        // --- MODIFICATION END ---
      }
      
      // Increment wrong answers count for session score
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

// === Helper function to remove question from wrong questions list ===
function removeFromWrongQuestions(sheet, questionText) {
  const wrongQuestions = getWrongQuestions(sheet);
  const updatedWrongQuestions = wrongQuestions.filter(q => q !== questionText);
  setWrongQuestions(sheet, updatedWrongQuestions);
}

// === Helper function to show next random question ===
function showNextQuestion(sheet, spreadsheet) {
  const category = sheet.getRange('A1').getValue();
  const questionCell = sheet.getRange('A4');
  const answerCell = sheet.getRange('A8');
  const showAnswerCheckbox = sheet.getRange('B7');
  const retryWrongChecked = sheet.getRange('B3').getValue();
  const currentCount = getQuestionCounter(sheet);
  if (category) {
    const datastore = spreadsheet.getSheetByName('datastore');
    const data = datastore.getDataRange().getValues();
    const filtered = data.filter((row, index) => index !== 0 && row[1] === category);
    if (filtered.length > 0) {
      let questionsToUse = [];
      if (retryWrongChecked) {
        // Get wrong questions for retry mode
        const wrongQuestions = getWrongQuestions(sheet);
        questionsToUse = filtered.filter(row => wrongQuestions.includes(row[2]));
        
        // In retry mode, check if all wrong questions have been completed
        if (questionsToUse.length === 0) {
          questionCell.setValue("Quiz Complete! All wrong questions have been answered correctly.");
          answerCell.setValue(''); // Clear answer when quiz completes
          // Disable checkboxes and stop quiz
          toggleRightWrongCheckboxes(sheet, false);
          sheet.getRange('B2').setValue(false); // Uncheck Start Quiz
          resetQuestionCounter(sheet);
          resetUsedQuestions(sheet);
          return;
        }
      } else {
        // Normal mode - check if we've reached 5 questions
        if (currentCount >= 5) { 
          questionCell.setValue("Quiz Complete! You have answered 5 questions.");
          answerCell.setValue(''); // Clear answer when quiz completes
          // Disable checkboxes and stop quiz
          toggleRightWrongCheckboxes(sheet, false);
          sheet.getRange('B2').setValue(false); // Uncheck Start Quiz
          resetQuestionCounter(sheet);
          resetUsedQuestions(sheet);
          // Reset used questions when quiz completes
          return;
        }
        questionsToUse = filtered;
      }
      
      const usedQuestions = getUsedQuestions(sheet);
      const availableQuestions = questionsToUse.filter(row => !usedQuestions.includes(row[2]));
      
      if (availableQuestions.length > 0) {
        const random = availableQuestions[Math.floor(Math.random() * availableQuestions.length)];
        questionCell.setValue(random[2]); // Question in column C
        addUsedQuestion(sheet, random[2]);
        // Track this question as used
        
        // Increment question counter
        setQuestionCounter(sheet, currentCount + 1);
        // Clear both Right/Wrong checkboxes for the next question
        sheet.getRange('B5').setValue(false);
        sheet.getRange('B6').setValue(false);
        // Handle Show Answer checkbox - if checked, show new answer;
        // if unchecked, clear answer
        const showAnswerChecked = showAnswerCheckbox.getValue();
        if (showAnswerChecked) {
          toggleAnswer(sheet, spreadsheet, true);
        } else {
          answerCell.setValue('');
        }
      } else {
        // No more unique questions available, end quiz early
        if (retryWrongChecked) {
          questionCell.setValue("Quiz Complete! All wrong questions have been answered correctly.");
        } else {
          questionCell.setValue("Quiz Complete! No more unique questions available for this category.");
        }
        answerCell.setValue(''); // Clear answer when quiz completes
        // Disable checkboxes and stop quiz
        toggleRightWrongCheckboxes(sheet, false);
        sheet.getRange('B2').setValue(false); // Uncheck Start Quiz
        resetQuestionCounter(sheet);
        resetUsedQuestions(sheet);
      }
    } else {
      questionCell.setValue("No questions available for this category.");
      answerCell.setValue('');
      // Clear answer
      resetQuestionCounter(sheet);
      resetUsedQuestions(sheet);
    }
  }
}