function onEdit(e) {
  var ss = e.source;
  var sheet = ss.getSheetByName('Quiz');
  var range = e.range;
  if (!sheet || sheet.getName() !== range.getSheet().getName()) return;

  var cell = range.getA1Notation();

  if (cell === 'C6' && range.getValue() === true) {
    showRandomQuestion();
  }

  if (cell === 'C8' && range.getValue() === true) {
    showAnswer();
  }

  if (cell === 'C14' && range.getValue() === true) {
    updateCountAndAskNext('right');
  }

  if (cell === 'C16' && range.getValue() === true) {
    updateCountAndAskNext('wrong');
  }
}
