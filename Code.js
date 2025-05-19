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

  // === 3. Add onEdit trigger (only if not already set) ===
  const triggers = ScriptApp.getProjectTriggers();
  const hasOnEdit = triggers.some(trigger => trigger.getHandlerFunction() === 'handleCheckboxEdit');

  if (!hasOnEdit) {
    ScriptApp.newTrigger('handleCheckboxEdit')
      .forSpreadsheet(ss)
      .onEdit()
      .create();
  }
}

// === 4. Respond to checkbox and category edits ===
function handleCheckboxEdit(e) {
  if (!e) return;

  const sheet = e.source.getSheetByName('quiz');
  const range = e.range;
  const cell = range.getA1Notation();

  // === If category changed in A1 ===
  if (sheet.getName() === 'quiz' && cell === 'A1') {
    sheet.getRange('A4').setValue('');      // Clear question
    sheet.getRange('B2').setValue(false);   // Uncheck checkbox
    return;
  }

  // === If checkbox in B2 is checked ===
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
      } else {
        questionCell.setValue("No questions available for this category.");
      }
    } else {
      questionCell.setValue('');
    }
  }
}
