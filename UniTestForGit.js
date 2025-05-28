// let __wrongAnswersTrackerBackup = null;

// function backupWrongAnswersTracker() {
//   const ss = SpreadsheetApp.getActiveSpreadsheet();
//   const trackerSheet = ss.getSheetByName('WrongAnswersTracker');
//   if (trackerSheet) {
//     __wrongAnswersTrackerBackup = trackerSheet.getDataRange().getValues();
//   } else {
//     __wrongAnswersTrackerBackup = null; // No tracker sheet to backup
//   }
// }

// function restoreWrongAnswersTracker() {
//   const ss = SpreadsheetApp.getActiveSpreadsheet();
//   let trackerSheet = ss.getSheetByName('WrongAnswersTracker');

//   // Ensure the tracker sheet exists and has headers
//   if (!trackerSheet) {
//     // If the sheet didn't exist before backup, setupWrongAnswersTrackerSheet will create it.
//     // This function is assumed to be in Code.txt and handles header creation.
//     setupWrongAnswersTrackerSheet(); 
//     trackerSheet = ss.getSheetByName('WrongAnswersTracker'); // Get reference after creation
//   } else {
//     // Ensure headers are present if sheet existed but was empty
//     const headerRange = trackerSheet.getRange('A1:C1');
//     if (headerRange.isBlank()) {
//       headerRange.setValues([['Question', 'Category', 'Wrong Count']]).setFontWeight('bold');
//     }
//   }

//   // Clear existing content below the header
//   const lastRow = trackerSheet.getLastRow();
//   if (lastRow > 1) {
//     trackerSheet.getRange(2, 1, lastRow - 1, trackerSheet.getLastColumn()).clearContent();
//   }

//   if (__wrongAnswersTrackerBackup !== null && __wrongAnswersTrackerBackup.length > 1) {
//     // Write the backed-up data, skipping the header row from the backup
//     const numCols = __wrongAnswersTrackerBackup[0].length;
//     trackerSheet.getRange(2, 1, __wrongAnswersTrackerBackup.length - 1, numCols)
//                 .setValues(__wrongAnswersTrackerBackup.slice(1));
//   } else if (__wrongAnswersTrackerBackup === null) {
//     // If there was no backup (meaning the sheet didn't exist before backup), ensure it's empty except headers
//     // This case is already handled by the clear above and the initial setup.
//   }
// }


// function setupTestResultsSheet() {
//   const ss = SpreadsheetApp.getActiveSpreadsheet();
//   let testSheet = ss.getSheetByName('testResults');
//   if (!testSheet) {
//     testSheet = ss.insertSheet('testResults');
//     testSheet.getRange('A1:D1').setValues([['Unit Test Name', 'Test Name', 'Result', 'Details']]);
//     testSheet.getRange('A1:D1').setFontWeight('bold');
//     testSheet.setFrozenRows(1);
//     testSheet.autoResizeColumns(1, 4);
//     testSheet.setColumnWidth(4, 300);
//   }
//   return testSheet;
// }

// function clearTestResults() {
//   const ss = SpreadsheetApp.getActiveSpreadsheet();
//   let testSheet = ss.getSheetByName('testResults');
//   if (testSheet) {
//     const lastRow = testSheet.getLastRow();
//     if (lastRow > 1) {
//       testSheet.getRange(2, 1, lastRow - 1, 4).clearContent();
//       testSheet.getRange(2, 1, lastRow - 1, 4).clearFormat();
//     }
//   }
// }

// function recordTestResult(unitTestName, testName, passed, details) {
//   const testSheet = setupTestResultsSheet();
//   const lastRow = Math.max(1, testSheet.getLastRow());

//   testSheet.getRange(lastRow + 1, 1, 1, 4).setValues([
//     [
//       unitTestName,
//       testName,
//       passed ? 'PASSED' : 'FAILED',
//       details
//     ]
//   ]);
//   const resultCell = testSheet.getRange(lastRow + 1, 3);
//   resultCell.setBackground(passed ? '#b7e1cd' : '#f4c7c3');

//   testSheet.autoResizeColumns(1, 4);
// }

// function isRangeProtected(sheet, rangeA1) {
//   const protections = sheet.getProtections(SpreadsheetApp.ProtectionType.RANGE);
//   return protections.some(protection => {
//     const protectedRange = protection.getRange();
//     return protectedRange.getA1Notation() === rangeA1;
//   });
// }

// function createMockEditEvent(spreadsheet, range, newValue, oldValue) {
//   range.setValue(newValue);

//   return {
//     source: spreadsheet,
//     range: range,
//     value: newValue,
//     oldValue: oldValue
//   };
// }

// // Helper function to setup test data in datastore sheet
// function setupTestDatastore() {
//   const ss = SpreadsheetApp.getActiveSpreadsheet();
//   let datastoreSheet = ss.getSheetByName('datastore');

//   if (!datastoreSheet) {
//     datastoreSheet = ss.insertSheet('datastore');
//   }

// }

// // Helper function to reset quiz sheet to clean state
// function resetQuizSheet() {
//   const ss = SpreadsheetApp.getActiveSpreadsheet();
//   const quizSheet = ss.getSheetByName('quiz');

//   // Preserve wrong questions before clearing the sheet
//   const wrongQuestions = getWrongQuestions(quizSheet);
//   // Log wrong questions before reset
//   const wrongQuestionsBefore = getWrongQuestions(quizSheet);
//   console.log('Before Reset:', JSON.stringify(wrongQuestionsBefore));
//   // Clear all values and reset state
//   quizSheet.getRange('A1').setValue('Test Category');
//   quizSheet.getRange('B2').setValue(false); // Start Quiz
//   quizSheet.getRange('B3').setValue(false);
//   // Retry Wrong Questions
//   quizSheet.getRange('A4').setValue(''); // Question
//   quizSheet.getRange('B5').setValue(false); // Right
//   quizSheet.getRange('B6').setValue(false); // Wrong
//   quizSheet.getRange('B7').setValue(false);
//   // Show Answer
//   quizSheet.getRange('A8').setValue(''); // Answer
//   quizSheet.getRange('B9').setValue(0); // Right count
//   quizSheet.getRange('B10').setValue(0);
//   // Wrong count

//   // Reset counters and lists
//   resetQuestionCounter(quizSheet);
//   resetUsedQuestions(quizSheet);
//   // Restore wrong questions after reset
//   setWrongQuestions(quizSheet, wrongQuestions);

//   // Remove all protections
//   toggleRightWrongCheckboxes(quizSheet, false);
//   // Log wrong questions after reset
//   const wrongQuestionsAfter = getWrongQuestions(quizSheet);
//   console.log('After Reset:', JSON.stringify(wrongQuestionsAfter));
//   // Restore wrong questions after reset
//   setWrongQuestions(quizSheet, wrongQuestionsBefore);
// }





// /**
//  * Run all unit tests and record results in the test results sheet
//  */
// function runAllTestsGit() {
//   clearTestResults();
//   backupWrongAnswersTracker(); // Backup before running tests

//   testWrongQuestionsResetInNormalMode();
//   testUpdateWrongAnswersTracker_NewQuestion();
//   testUpdateWrongAnswersTracker_ExistingQuestion();
//   testHandleCheckboxEdit_WrongCheckboxTracking();
//   testHandleCheckboxEdit_RightCheckboxNoTracking();

//   restoreWrongAnswersTracker(); // Restore after all tests
// }

