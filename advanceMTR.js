// Purpose: to duplicate last month's tab in the MTR Workbook (Google Sheets) and prepare a new sheet for current month MTR reporting

// Future Development: None currently identified

// To note: 
// This script is developed as a Google Apps Script container script: i.e. a script that is bound to a specific file, such as a Google Sheets, Google Docs, or Google Forms file. This container script acts as the file's custom script, allowing users to extend the functionality of the file by adding custom functions, triggers, and menu items to enhance its behavior and automation.

// Instructions for Using this Script in your container file:
// 1. Open a new or existing Google Sheets file where you want to use the script.
// 2. Click on "Extensions" in the top menu and select "Apps Script" from the dropdown menu. This will open the Google Apps Script editor in a new tab.
// 3. Copy and paste the provided script into the Apps Script editor, replacing the existing code (if any).
// 4. Sheet Naming Conventions: The script assumes that the sheets in your Google Sheets file follow a specific naming convention to work correctly. The script will look for sheets with names that match this format to find the most recent sheet. The naming convention is as follows: 
  // Sheets representing past months should be named in the format "MonthName Year" (e.g., "January 2023").
// 5. In the function replaceCellText, replace cell 'c1' with desired cell in which to have the generated date entered
// 6. In the function clearUniqueWorkText, replace cell 'b19' with desired cell in which to have the context cleared
// Before running the script, make sure to save the script file in the Apps Script editor by clicking on the floppy disk icon or pressing Ctrl + S (Windows/Linux) or Cmd + S (Mac). Once saved, you can run the createNewTab function by clicking on the play button ▶️ or using the keyboard shortcut Ctrl + Enter (Windows/Linux) or Cmd + Enter (Mac). This function will create a new sheet with the name of the next month and based on the most recent sheet.

////////////////////////////////////////////////

// helper function to duplicate last month's sheet
function duplicateOldSheet(ss) {
  var sheets = ss.getSheets();
  var mostRecentSheet;
  var mostRecentDate = new Date(1970, 0, 1); // Set an initial date far in the past

  for (var i = 0; i < sheets.length; i++) {
    var sheetName = sheets[i].getName();
    var dateParts = sheetName.split(" ");
    if (dateParts.length === 2) {
      var sheetDate = new Date(dateParts[1], getMonthIndex(dateParts[0]), 1);
      if (sheetDate > mostRecentDate) {
        mostRecentSheet = sheets[i];
        mostRecentDate = sheetDate;
      }
    }
  }

  var newSheet = mostRecentSheet.copyTo(ss);
  var nextMonth = updateTabName(mostRecentDate.getMonth());
  newSheet.setName(nextMonth);
  return newSheet;
}

//helper function to move the newSheet to the front of the workbook
function moveNewSheet(ss, newSheet) {
  ss.setActiveSheet(newSheet);
  ss.moveActiveSheet(1);
}

function getMonthIndex(monthName) {
  var months = [
    "January", "February", "March", "April", "May", "June",
    "July", "August", "September", "October", "November", "December"
  ];
  return months.indexOf(monthName);
}

// helper function to find the Current month
function findCurrentMonth() {
  var today = new Date();
  var currentMonth = today.getMonth();
  return currentMonth;
}

// helper function to update the tabName for the newSheet with next month in the format MonthName YYYY
function updateTabName(currentMonth) {
  var months = [
    "January", "February", "March", "April", "May", "June",
    "July", "August", "September", "October", "November", "December"
  ];
  var nextMonth = (currentMonth + 1) % 12;
  var year = new Date().getFullYear();
  var tabName = months[nextMonth] + " " + year;
  return tabName;
}

// helper function to find cell c1 and replace with the last date of nextMonth in the format MM/DD/YY
function replaceCellText(newSheet, nextMonth) {
  var lastDay = new Date(new Date(nextMonth).getFullYear(), new Date(nextMonth).getMonth() + 1, 0);
  newSheet.getRange("C1").setValue(Utilities.formatDate(lastDay, "GMT", "MM/dd/YYYY")); // <-- Replace cell c1 with desired date cell
}

// helper function to find the merged cell beginning in B19 and clear the contents. 
function clearUniqueWorkText(newSheet) {
  var range = newSheet.getRange("B19");
  var mergedRange = range.getMergedRanges();
  if (mergedRange.length > 0) {
    mergedRange[0].clearContent();
  }
}

// Primary function to create the new tab
function createNewTab() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var currentMonth = findCurrentMonth();
  var newSheet = duplicateOldSheet(ss);
  var nextMonth = updateTabName(currentMonth);
  moveNewSheet(ss, newSheet);
  replaceCellText(newSheet, nextMonth);
  clearUniqueWorkText(newSheet);
}