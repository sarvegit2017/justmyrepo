function startQuiz() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var quizSheet = ss.getSheetByName('Quiz');
  
  var selectedCategory = quizSheet.getRange('C2').getValue();
  if (!selectedCategory) {
    SpreadsheetApp.getUi().alert('Please select a category first.');
    quizSheet.getRange('C3').setValue(false);
    return;
  }
  
  // Reset counters
  quizSheet.getRange('F3').setValue(0);  // Questions Asked
  quizSheet.getRange('F4').setValue('0/5'); // Progress
  quizSheet.getRange('F5').setValue('0%'); // Score
  quizSheet.getRange('C20').setValue(0); // Right Count
  quizSheet.getRange('C22').setValue(0); // Wrong Count
  
  // Show the first question
  showRandomQuestion();
  
  // Uncheck the Start Quiz checkbox
  quizSheet.getRange('C3').setValue(false);
}

function showRandomQuestion() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var quizSheet = ss.getSheetByName('Quiz');
  var dataSheet = ss.getSheetByName('datastore');

  // Check if we've already completed 5 questions
  var countCell = quizSheet.getRange('F3');  // Questions Asked cell
  var count = countCell.getValue();
  
  if (count >= 5) {
    quizSheet.getRange('F3').setValue("Quiz complete");
    quizSheet.getRange('C6').setValue("Quiz complete! You answered 5 questions.");
    
    // Calculate and show final score
    var rightCount = quizSheet.getRange('C20').getValue();
    var percentage = Math.round((rightCount / 5) * 100);
    quizSheet.getRange('F5').setValue(percentage + '%');
    quizSheet.getRange('C12').setValue('Your final score: ' + percentage + '% (' + rightCount + ' out of 5 correct)');
    return;
  }

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
    quizSheet.getRange('C6').setValue('No questions found in this category.');
    quizSheet.getRange('C12').setValue('');
    return;
  }

  var random = qaPairs[Math.floor(Math.random() * qaPairs.length)];
  quizSheet.getRange('C6').setValue(random.question);
  quizSheet.getRange('Z1').setValue(random.answer);
  quizSheet.getRange('C12').setValue('');

  // Update the question counter and progress
  var newCount = (!isNaN(count) ? count : 0) + 1;
  countCell.setValue(newCount);
  quizSheet.getRange('F4').setValue(newCount + '/5'); // Update progress

  // Reset checkboxes
  quizSheet.getRange('C11').setValue(false); // Show Answer
  quizSheet.getRange('C16').setValue(false); // Next Question
  quizSheet.getRange('C19').setValue(false); // Right
  quizSheet.getRange('C21').setValue(false); // Wrong
}

function showAnswer() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var quizSheet = ss.getSheetByName('Quiz');
  var answer = quizSheet.getRange('Z1').getValue();

  quizSheet.getRange('C12').setValue(answer || 'No answer available.');
  quizSheet.getRange('C11').setValue(false);
}

function updateCountAndAskNext(type) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var quizSheet = ss.getSheetByName('Quiz');
  
  // Update the appropriate counter based on right/wrong
  var cell = type === 'right' ? 'C20' : 'C22';
  var count = quizSheet.getRange(cell).getValue();
  quizSheet.getRange(cell).setValue((!isNaN(count) ? count : 0) + 1);

  // Update current score percentage
  var rightCount = quizSheet.getRange('C20').getValue();
  var questionsAsked = quizSheet.getRange('F3').getValue();
  var percentage = Math.round((rightCount / questionsAsked) * 100);
  quizSheet.getRange('F5').setValue(percentage + '%');

  // Reset the checkboxes
  quizSheet.getRange('C19').setValue(false); // Right checkbox
  quizSheet.getRange('C21').setValue(false); // Wrong checkbox

  // Check if we've completed 5 questions
  if (questionsAsked >= 5) {
    // Don't ask next question, just show completion message
    quizSheet.getRange('F3').setValue("Quiz complete");
    quizSheet.getRange('C6').setValue("Quiz complete! You answered 5 questions.");
    
    // Calculate and show final score
    percentage = Math.round((rightCount / 5) * 100);
    quizSheet.getRange('F5').setValue(percentage + '%');
    quizSheet.getRange('C12').setValue("Your final score: " + percentage + "% (" + rightCount + " out of 5 correct)");
    return;
  }

  showRandomQuestion();
}