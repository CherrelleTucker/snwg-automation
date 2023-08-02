// Purpose: Library script to calculate the current Program Increment and return result in the {{Adjusted PI}} placeholder text in active document

// Future Development: none currently identified

// To note: 
// This script is developed as a Google Apps Script library script: i.e. a script that is not bound to a specific file, such as a Google Sheets, Google Docs, or Google Forms file. This standalone script project contains reusable code and functions and can be shared and included in multiple other scripts, allowing developers to easily reuse code across different projects.

// How to use a library script: 
// In the script editor, click on "File" > "Project properties."
// In the "Script properties" tab, you will find the "Script ID." Copy this ID.
// To use the library in another script project, open the script editor of that project.
// Click on "Resources" > "Libraries."
// In the "Add a library" section, paste the Script ID and click "Add."
// Choose the version of the library you want to use (usually, you'll want to use the latest version).
// Set the identifier, which is the name you will use to reference the library functions in your main script project.
// Click "Save."
// After adding the library, you can use its functions in your main script project by calling them with the specified identifier. This way, you can take advantage of the shared code and easily maintain and update common functionalities across multiple projects.

// How to use this script: 
// Ensure placeholder text {{Adjusted PI}} is in the target file (the file you would like the information populated in).
// Verify that all Adjustment Weeks have been accounted for in Adjustment Weeks section of the library script
// Results will populate the placeholder text in the target file in FY.Q.S "Week" W format.

///////////////////////////////////

// Adjustment Weeks to account for holidays, vacations, etc. Add the first Sunday of the Adjustment Week at the bottom of the list in 'YYYY-MM-DD' format. 
var ADJUSTMENT_WEEKS = []; // Add adjustment week dates here.
ADJUSTMENT_WEEKS.push(new Date('2023-07-02')); // Innovation Week #2
ADJUSTMENT_WEEKS.push(new Date('2023-07-16')); // IGARSS 2023
// Add dates as 'yyyy-mm-dd'

///////////////////////////////////////

// Global variables
var BASE_DATE = new Date('2023-04-16');  // The date for '23.3.1 Week 1', the beginning sprint of record for the purposes of this script. 
var BASE_FY = 23;
var BASE_PI = 3;
var BASE_SPRINT = 1;
var BASE_WEEK = 1;

// Helper function to convert a date to the number of milliseconds since Jan 1, 1970
function toMillis(date) {
  return Date.UTC(date.getFullYear(), date.getMonth(), date.getDate());
}

// Helper function to calculate Adjustment Weeks
function adjustmentWeeks(date) {
  var adjustmentWeeks = 0;
  var currentTime = toMillis(date);
  for(var i = 0; i < ADJUSTMENT_WEEKS.length; i++) {
    if(toMillis(ADJUSTMENT_WEEKS[i]) <= currentTime) {
      adjustmentWeeks++;
    }
  }
  return adjustmentWeeks;
}

// Helper function to calculate the current fiscal year
function getFiscalYear(date) {
  var year = date.getFullYear();
  var month = date.getMonth();
  
  // The fiscal year starts in October (month index 9)
  return month >= 9 ? year + 1 : year;
}

// Helper function to get current PI (Q) (1-4)
function getQuarter(date) {
  // In JavaScript, months are 0-indexed, so October is 9 and September is 8
  var month = date.getMonth();
  
  // The first quarter of the fiscal year starts in October
  if (month >= 9) {
    return ((month - 9) / 3 | 0) + 1;
  } else {
    return ((month + 3) / 3 | 0) + 1;
  }
}

// Helper function to get current Sprint (1-6)
function getSprint(date) {
  var fiscalYearStart = new Date(getFiscalYear(date), 9, 1);
  var diffWeeks = Math.ceil(((date - fiscalYearStart + 1) / (24 * 60 * 60 * 1000)) / 7);
  
  // Subtract adjustment weeks
  diffWeeks -= adjustmentWeeks(date);
  
  // Subtract the base week of the quarter
  var baseWeek = (getQuarter(date) - 1) * 13;
  diffWeeks -= baseWeek;
  
  // Calculate the sprint based on the week of the fiscal year, after accounting for adjustment weeks
  var sprint = Math.ceil(diffWeeks / 2);
  
  return sprint;
}

// Helper function to get current week (1-2)
function getWeek(date) {
  return Math.ceil(((date - new Date(getFiscalYear(date), 9, 1) + 1) / (24 * 60 * 60 * 1000)) / 7) % 2 === 0 ? 2 : 1;
}

// Helper function to get week of the year
function getWeekOfYear(date) {
  var start = new Date(date.getFullYear(), 0, 1);
  var diff = date - start + (start.getTimezoneOffset() - date.getTimezoneOffset()) * 60 * 1000;
  var oneDay = 1000 * 60 * 60 * 24;
  var day = Math.floor(diff / oneDay);
  return Math.ceil(day / 7);
}

// Helper function to rename FY.Q.6.1 to "Innovation Week"
function renameInnovation(pi) {
  return pi.endsWith(".6.1") ? "Innovation Week" : pi;
}

// Helper function to rename FY.Q.6.2 to "Next PI Planning"
function renameNextPiPlanning(pi) {
  return pi.endsWith(".6.2") ? "Next PI Planning" : pi;
}

// Secondary function to replace placeholder text {{Adjusted PI}} with the calculated Adjusted PI
function populatePlaceholder() {
  var date = new Date();
  var adjustedPI = getCurrentPI(date);
  var body = DocumentApp.getActiveDocument().getBody();
  var text = body.editAsText();
  text.replaceText("{{Adjusted PI}}", adjustedPI);
}

// Primary function to calculate the Adjusted PI
function getCurrentPI(date) {
  var daysPassed = (toMillis(date) - toMillis(BASE_DATE)) / (24 * 60 * 60 * 1000) - adjustmentWeeks(date) * 7;
  var weeksPassed = Math.floor(daysPassed / 7);
  
  // calculate the total two week periods passed since the BASE_DATE
  var totalTwoWeekPeriods = Math.floor(weeksPassed / 2);
  
  var week = weeksPassed % 2 + 1;
  var sprint = totalTwoWeekPeriods % 6 + 1;
  var pi = Math.floor(totalTwoWeekPeriods / 6) % 4 + BASE_PI;
  var fy = BASE_FY + Math.floor(totalTwoWeekPeriods / (6 * 4));
  var piStr = "FY" + fy + "." + pi + "." + sprint + " Week " + week;  // Format result as FY.Q.S Week W
  piStr = renameInnovation(piStr);
  piStr = renameNextPiPlanning(piStr);
  
  return piStr;
}

// Call function to test with specific date
function testGetCurrentPI() {
  var currentDate = new Date('2023-08-21');
  var currentPI = getCurrentPI(currentDate);
  Logger.log(currentPI);
}

