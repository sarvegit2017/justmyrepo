function onEdit(e) {
  var ss = e.source;
  var sheet = ss.getSheetByName('Quiz');
  var range = e.range;
  if (!sheet || sheet.getName() !== range.getSheet().getName()) return;

  var cell = range.getA1Notation();

  // Handle Start Quiz checkbox
  if (cell === 'C3' && range.getValue() === true) {
    startQuiz();
  }

  // Handle Show Answer checkbox
  if (cell === 'C11' && range.getValue() === true) {
    showAnswer();
  }

  // Handle Next Question checkbox
  if (cell === 'C16' && range.getValue() === true) {
    showRandomQuestion();
  }

  // Handle Right checkbox
  if (cell === 'C19' && range.getValue() === true) {
    updateCountAndAskNext('right');
  }

  // Handle Wrong checkbox
  if (cell === 'C21' && range.getValue() === true) {
    updateCountAndAskNext('wrong');
  }
}