function showRandomQuestion() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var quizSheet = ss.getSheetByName('Quiz');
  var dataSheet = ss.getSheetByName('datastore');

  var selectedCategory = quizSheet.getRange('C2').getValue();
  if (!selectedCategory) {
    SpreadsheetApp.getUi().alert('Please select a category first.');
    return;
  }

  var data = dataSheet.getDataRange().getValues();
  var qaPairs = [];

  for (var i = 1; i < data.length; i++) {
    if (data[i][1] === selectedCategory) {
      qaPairs.push({ question: data[i][2], answer: data[i][3] });
    }
  }

  if (qaPairs.length === 0) {
    quizSheet.getRange('C4').setValue('No questions found in this category.');
    quizSheet.getRange('C10').setValue('');
    return;
  }

  var random = qaPairs[Math.floor(Math.random() * qaPairs.length)];
  quizSheet.getRange('C4').setValue(random.question);
  quizSheet.getRange('Z1').setValue(random.answer);
  quizSheet.getRange('C10').setValue('');

  var countCell = quizSheet.getRange('F1');
  var count = countCell.getValue();
  countCell.setValue((!isNaN(count) ? count : 0) + 1);

  quizSheet.getRange('C6').setValue(false);
  quizSheet.getRange('C8').setValue(false);
  quizSheet.getRange('C14').setValue(false);
  quizSheet.getRange('C16').setValue(false);
}

function showAnswer() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var quizSheet = ss.getSheetByName('Quiz');
  var answer = quizSheet.getRange('Z1').getValue();

  quizSheet.getRange('C10').setValue(answer || 'No answer available.');
  quizSheet.getRange('C8').setValue(false);
}

function updateCountAndAskNext(type) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var quizSheet = ss.getSheetByName('Quiz');
  var cell = type === 'right' ? 'C13' : 'C15';
  var count = quizSheet.getRange(cell).getValue();
  quizSheet.getRange(cell).setValue((!isNaN(count) ? count : 0) + 1);

  quizSheet.getRange('C14').setValue(false);
  quizSheet.getRange('C16').setValue(false);

  showRandomQuestion();
}
