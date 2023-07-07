// Purpose: Calculate the current PI. 
// // Project Increments (PI) are 12 weeks long.
// Projects Sprints within PIs are 2 weeks long.
// After Sprint 5, there are 2 weeks called "Innovation and Planning".
// This is followed by one week of "PI Planning".
// The next PI begins the Monday of the following week. 
// The fiscal year (FY) is from October 1 to September 30. 
// There are 4 PI Quarters in a Fiscal year, begining October of the preceding calendar year. Example: Q1FY23 begins October 1, 2022. 
// PI 23.3.1 begins on Monday, April 17 2023
// Done: Sprint display is in the format (fiscal)YY.Q.D.W
// In Development: If in Sprint 6, called "Innovation and Planning", return "Innovation and Planning", instead of numbers. If in PI Planning week, return "PI Planning."
// Future Development: there is one week of Planning Vacation from 2023-07-14 to 2023-07-21, that breaks the pattern. It restarts on 2023-07-24.


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
  var sprintsInPI = 6; // Includes the 5 sprints and the Innovation and Planning sprint
  var weeksInPI = sprintsInPI * sprintDuration;
  var weeksSincePIPlanning = weeksInPI + sprintsBeforePIPlanning * sprintDuration;
  var weeksSinceLastPIPlanning = weeksSincePIPlanning % weeksInPI;
  var currentSprint = Math.floor(weeksSinceLastPIPlanning / sprintDuration) + 1;

  // Determine the current week of the sprint
  var currentWeek = "Week " + (weeksSinceLastPIPlanning % sprintDuration + 1); // Add "Week " before the week number

  // Check if it's the Innovation and Planning sprint or PI Planning week
  if (currentSprint === 6) {
    currentSprint = "Innovation and Planning, " + currentWeek;
    currentWeek = ""; // Clear the week number as it's included in the sprint description
  } else if (weeksSinceLastPIPlanning >= weeksInPI) {
    currentSprint = "PI Planning";
    currentWeek = "";
  }

  // Format the PI number
  var currentPI = fiscalYear + '.' + currentQuarter + '.' + currentSprint + (currentWeek !== "" ? '.' + currentWeek : '');

  return currentPI;
}

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
