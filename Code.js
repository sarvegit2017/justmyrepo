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

  // === 4. Initially disable Right and Wrong checkboxes ===
  quizSheet.getRange('B5').protect().setDescription('Right checkbox - disabled until quiz starts');
  quizSheet.getRange('B6').protect().setDescription('Wrong checkbox - disabled until quiz starts');

  // === 5. Add onEdit trigger (only if not already set) ===
  const triggers = ScriptApp.getProjectTriggers();
  const hasOnEdit = triggers.some(trigger => trigger.getHandlerFunction() === 'handleCheckboxEdit');

  if (!hasOnEdit) {
    ScriptApp.newTrigger('handleCheckboxEdit')
      .forSpreadsheet(ss)
      .onEdit()
      .create();
  }
}

// === Helper function to enable/disable Right and Wrong checkboxes ===
function toggleRightWrongCheckboxes(sheet, enable) {
  const rightCheckbox = sheet.getRange('B5');
  const wrongCheckbox = sheet.getRange('B6');
  
  if (enable) {
    // Remove protection to enable checkboxes
    const protections = sheet.getProtections(SpreadsheetApp.ProtectionType.RANGE);
    protections.forEach(protection => {
      const range = protection.getRange();
      if (range.getA1Notation() === 'B5' || range.getA1Notation() === 'B6') {
        protection.remove();
      }
    });
  } else {
    // Add protection to disable checkboxes
    rightCheckbox.protect().setDescription('Right checkbox - disabled until quiz starts');
    wrongCheckbox.protect().setDescription('Wrong checkbox - disabled until quiz starts');
    // Clear checkbox values when disabling
    rightCheckbox.setValue(false);
    wrongCheckbox.setValue(false);
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

// === Respond to checkbox and category edits ===
function handleCheckboxEdit(e) {
  if (!e) return;

  const sheet = e.source.getSheetByName('quiz');
  const range = e.range;
  const cell = range.getA1Notation();

  // === If category changed in A1 ===
  if (sheet.getName() === 'quiz' && cell === 'A1') {
    sheet.getRange('A4').setValue('');     // Clear question
    sheet.getRange('B2').setValue(false);  // Uncheck Start Quiz checkbox
    resetQuestionCounter(sheet);           // Reset question counter
    toggleRightWrongCheckboxes(sheet, false); // Disable Right/Wrong checkboxes
    return;
  }

  // === If Start Quiz checkbox in B2 is checked ===
  if (sheet.getName() === 'quiz' && cell === 'B2') {
    const isChecked = range.getValue();
    const category = sheet.getRange('A1').getValue();
    const questionCell = sheet.getRange('A4');

    if (isChecked && category) {
      const datastore = e.source.getSheetByName('datastore');
      const data = datastore.getDataRange().getValues();
      const filtered = data.filter((row, index) => index !== 0 && row[1] === category);

      if (filtered.length > 0) {
        const random = filtered[Math.floor(Math.random() * filtered.length)];
        questionCell.setValue(random[2]); // Question in column C
        setQuestionCounter(sheet, 1); // Start with question 1
        // Enable Right/Wrong checkboxes when quiz starts
        toggleRightWrongCheckboxes(sheet, true);
      } else {
        questionCell.setValue("No questions available for this category.");
        resetQuestionCounter(sheet);
        // Disable checkboxes if no questions available
        toggleRightWrongCheckboxes(sheet, false);
      }
    } else {
      questionCell.setValue('');
      resetQuestionCounter(sheet);
      // Disable Right/Wrong checkboxes when quiz stops
      toggleRightWrongCheckboxes(sheet, false);
    }
  }

  // === If Right checkbox in B5 is checked ===
  if (sheet.getName() === 'quiz' && cell === 'B5') {
    const isChecked = range.getValue();
    const startQuizChecked = sheet.getRange('B2').getValue();
    
    // Only proceed if Start Quiz is checked
    if (isChecked && startQuizChecked) {
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
      // Uncheck Right checkbox
      sheet.getRange('B5').setValue(false);
      // Show next question
      showNextQuestion(sheet, e.source);
    } else if (isChecked && !startQuizChecked) {
      // If Start Quiz is not checked, uncheck this checkbox
      sheet.getRange('B6').setValue(false);
    }
  }
}

// === Helper function to show next random question ===
function showNextQuestion(sheet, spreadsheet) {
  const category = sheet.getRange('A1').getValue();
  const questionCell = sheet.getRange('A4');
  const currentCount = getQuestionCounter(sheet);

  // Check if we've reached 5 questions
  if (currentCount >= 5) {
    questionCell.setValue("Quiz Complete! You have answered 5 questions.");
    // Disable checkboxes and stop quiz
    toggleRightWrongCheckboxes(sheet, false);
    sheet.getRange('B2').setValue(false); // Uncheck Start Quiz
    resetQuestionCounter(sheet);
    return;
  }

  if (category) {
    const datastore = spreadsheet.getSheetByName('datastore');
    const data = datastore.getDataRange().getValues();
    const filtered = data.filter((row, index) => index !== 0 && row[1] === category);

    if (filtered.length > 0) {
      const random = filtered[Math.floor(Math.random() * filtered.length)];
      questionCell.setValue(random[2]); // Question in column C
      
      // Increment question counter
      setQuestionCounter(sheet, currentCount + 1);
      
      // Clear both Right/Wrong checkboxes for the next question
      sheet.getRange('B5').setValue(false);
      sheet.getRange('B6').setValue(false);
    } else {
      questionCell.setValue("No questions available for this category.");
      resetQuestionCounter(sheet);
    }
  }
}