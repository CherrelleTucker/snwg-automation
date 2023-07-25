// Purpose: to duplicate last month's tab and prepare a new sheet for current month MTR reporting

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
    newSheet.getRange("C1").setValue(Utilities.formatDate(lastDay, "GMT", "MM/dd/YYYY"));
  }
  
  // helper function to find the merged cell beginning in B19 and clear the contents. 
  function clearUniqueWorkText(newSheet) {
    var range = newSheet.getRange("B19");
    var mergedRange = range.getMergedRanges();
    if (mergedRange.length > 0) {
      mergedRange[0].clearContent();
    }
  }
  