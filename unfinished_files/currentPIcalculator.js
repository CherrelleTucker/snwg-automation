// Purpose: Library script to calculate the current Program Increment and return result in the {{Current PI}} placeholder text in active document
// In Development: If in Sprint 6.1, called "Innovation Week", return "Innovation Week", instead of numbers. If in PI Planning week, return "next PI Planning Week."
// Future Development: there is one week of Planning Vacation from 2023-07-14 to 2023-07-21, that breaks the pattern. PI23.4 starts on 2023-07-24.


function getCurrentPI() {
  // Get the current date
  var currentDate = new Date();

  // Determine the current fiscal year
  var fiscalYear = currentDate.getFullYear();
  if (currentDate.getMonth() >= 9) {
    fiscalYear++; // Fiscal year starts from October
  }
  fiscalYear = fiscalYear.toString().substr(-2); // Extract the last two digits

  // Determine the current quarter
  var fiscalMonth = (currentDate.getMonth() - 9 + 12) % 12 + 1; // Get the month number in fiscal year (where Oct is the 1st month)
  var currentQuarter = Math.ceil(fiscalMonth / 3);

  // Determine the current sprint
  var sprintDuration = 2; // Duration of each sprint in weeks
  var sprintsBeforePIPlanning = 5;
  var sprintsInPI = 6; // Includes the 5 sprints and the Innovation week and next PI Planning week
  var weeksInPI = sprintsInPI * sprintDuration;
  var weeksSincePIPlanning = weeksInPI + sprintsBeforePIPlanning * sprintDuration;
  var weeksSinceLastPIPlanning = weeksSincePIPlanning % weeksInPI;
  var currentSprint = Math.floor(weeksSinceLastPIPlanning / sprintDuration) + 1;

  // Determine the current week of the sprint
  var currentWeek = "Week " + (weeksSinceLastPIPlanning % sprintDuration + 1); // Add "Week " before the week number

  // Check if it's the Innovation week or next PI Planning week
  if (currentSprint === 6) {
    currentSprint = "Innovation, " + currentWeek;
    currentWeek = ""; // Clear the week number as it's included in the sprint description
  } else if (weeksSinceLastPIPlanning >= weeksInPI) {
    currentSprint = "Next PI Planning";
    currentWeek = "";
  }

  // Format the PI number
  var currentPI = fiscalYear + '.' + currentQuarter + '.' + currentSprint + (currentWeek !== "" ? '.' + currentWeek : '');

  return currentPI;
}

// primary function to populate the result of the PI calculation
function populateCurrentPI() {
  // Get the current PI
  var currentPI = getCurrentPI();

  // Get the active document
  var doc = DocumentApp.getActiveDocument();
  if (doc == null) {
    Logger.log("No active document found.");
    return;
  }

  // Find and replace the placeholder text
  var placeholderText = '{{Current PI}}'; //Escaping special characters for regex
  var body = doc.getBody();
  body.replaceText(placeholderText, currentPI);
}

// Call the function to replace placeholders with the current PI
populateCurrentPI();
