/*
Script Name: createPiCalendar

Description:
This Google Apps Script is designed to automatically populate a Google Calendar with 
project management events such as sprints, sprint reviews, an innovation week, and project 
increment (PI) planning activities. It also changes the color of these events for easy 
identification and categorization. All inputs for Fiscal Year, Project Increment, and 
start date are predefined within the script.

Prerequisites:
- Familiarity with Google Apps Script.
- Access to a Google Calendar and its ID.
- Basic understanding of JavaScript.

Setup:
1. Set the `calendarId` variable to the ID of your target Google Calendar.
   Example: var calendarId = "your_calendar_id_here";
2. Predefine the Fiscal Year, Project Increment, and start date within the script.
   Example: 
   var startingFY = 24; // Fiscal Year
   var startingPI = 2; // Project Increment
   var startDate = new Date("2024-04-13"); // Start date

Execution:
- Run the `populateCalendarEvents` function from the script editor to execute the script.

Script Functions:
- Creates events for sprints, sprint reviews, innovation week, and PI planning in the specified Google Calendar.
- Assigns different colors to each type of event based on a predefined color map.

Outputs:
- Events will be created and visualized in the specified Google Calendar.

Post-Execution:
- Verify the created events in the Google Calendar.
- Check the script's execution logs for any errors or important messages.

Troubleshooting:
- Ensure the correct Google Calendar ID is provided.
- Review execution logs for any errors if events do not appear as expected.
- Confirm that the script has necessary permissions to modify the Google Calendar.
- Uncomment clearEvents.gs when testing to clear all events prior to the generation of new events. 

Note: This script is intended for project management purposes and should be customized according to specific project timelines and requirements.
 */

// "c_8798ebb71e4f29ffc300845dabe847152b8c92e2afd062e0e31242d7fce780cd@group.calendar.google.com"; Test calendar ID

// Global variables
var calendarId = "c_e6e532cefc5ddfdd7f3c715e7a07326607cd240d951991f6a4e3b87653e67ef3@group.calendar.google.com"; // IMPACT PI Calendar ID
var startingFY; // Fiscal Year
var startingPI; // Project Increment
var startDate; // Start Date for PI

// Main function to populate calendar events either based on user input or predefined values.
function populateCalendarEvents() {
  // Option for user input - uncomment the following line to enable
  // getUserInputValues();

  // Or use predefined values
  startingFY = 24; // Fiscal Year
  startingPI = 2; // Project Increment (base four)
  startDate = new Date("2024-04-13"); // PI start date; use the date of the Saturday before the first Monday. 

  // Populating the events
  populateSprintEvents(startDate, startingFY, startingPI);
  populateInnovationWeek(startDate);
  populateNextPIPlanningWeek(startDate, startingFY, startingPI);
  changeEventColors();
}

// Function to get user input for Fiscal Year and Project Increment
function getUserInputValues() {
  var ui = SpreadsheetApp.getUi();
  var response = ui.prompt('Please enter the new PI number in format FY.PI (e.g., 24.2)');
  var input = response.getResponseText();
  var splitInput = input.split(".");
  startingFY = parseInt(splitInput[0]);
  startingPI = parseInt(splitInput[1]);
  startDate = new Date(); // Alternatively, prompt for a specific start date
}

// Prompt for future dashboard development: Function to prompt the user to enter the PI number in the format YY.Q and parse the input. 
/*function getUserInputValues() {
  var ui = SpreadsheetApp.getUi();
  var response = ui.prompt('Please enter the new PI number in format YY.Q');
  var input = response.getResponseText();
  var splitInput = input.split(".");
  startingFY = parseInt(splitInput[0]);
  startingPI = parseInt(splitInput[1]);
  startDate = new Date(); // Set to current date or prompt user for a specific start date
}*/

//Function to populate a series of sprint events and sprint review events in the calendar./
function populateSprintEvents(piStartDate, fiscalYear, projectIncrement) {
  for (var sprint = 1; sprint <= 5; sprint++) {
    var sprintStartDate = new Date(piStartDate.getTime());
    sprintStartDate.setDate(sprintStartDate.getDate() + (sprint - 1) * 14); // Calculate start date for each sprint

    // Adjust to start on a Saturday
    if (sprintStartDate.getDay() != 6) {
      sprintStartDate.setDate(sprintStartDate.getDate() + (6 - sprintStartDate.getDay()));
    }

    var sprintEndDate = new Date(sprintStartDate.getTime());
    sprintEndDate.setDate(sprintStartDate.getDate() + 13); // 2 weeks later

    var eventName = "PI " + fiscalYear + "." + projectIncrement + " Sprint " + sprint;
    createEvent(eventName, sprintStartDate, sprintEndDate);

    // Sprint Review on the last Friday of each sprint
    var sprintReviewStart = new Date(sprintEndDate.getTime());
    sprintReviewStart.setDate(sprintReviewStart.getDate()); // Move to Friday
    sprintReviewStart.setHours(10, 0, 0, 0); // Set time to 10:00 AM

    var sprintReviewEnd = new Date(sprintReviewStart.getTime());
    sprintReviewEnd.setHours(12, 0, 0, 0); // Set time to 12:00 PM

    createEvent("Sprint Review - " + eventName, sprintReviewStart, sprintReviewEnd);
  }
}

//function to create a Google Calendar event with the specified name, start and end dates.
function createEvent(eventName, startDate, endDate) {
  var calendar = CalendarApp.getCalendarById(calendarId);
  // Ensure endDate is after startDate
  if (endDate > startDate) {
    calendar.createEvent(eventName, startDate, endDate);
  } else {
    Logger.log("Error creating event: " + eventName + ". End date is before start date.");
  }
}

// Function to populate the event for Innovation Week. Innovation Week starts on the Saturday following the last sprint and lasts for 7 days.
function populateInnovationWeek(piStartDate) {
  var innovationWeekStartDate = new Date(piStartDate);
  
  // Set the start date for Innovation Week: Saturday following the last sprint
  innovationWeekStartDate.setDate(innovationWeekStartDate.getDate() + 10 * 7); // After Sprint 5, Week 2
  if (innovationWeekStartDate.getDay() !== 6) {
    // Adjust to the next Saturday
    innovationWeekStartDate.setDate(innovationWeekStartDate.getDate() + (6 - innovationWeekStartDate.getDay()));
  }

  var innovationWeekEndDate = new Date(innovationWeekStartDate);
  innovationWeekEndDate.setDate(innovationWeekStartDate.getDate() + 6); // 7 days later, ensuring a 7-day week

  createEvent("Innovation Week", innovationWeekStartDate, innovationWeekEndDate);
}

// Function to calculate the start and end dates and create events for Next PI Planning Week.
function populateNextPIPlanningWeek(piStartDate, fiscalYear, projectIncrement) {
  var nextPI = (projectIncrement % 4) + 1; // Incrementing PI within base 4
  var nextFY = fiscalYear + (nextPI === 1 ? 1 : 0); // Incrementing FY if PI resets

  var planningWeekStartDate = new Date(piStartDate);
  planningWeekStartDate.setDate(planningWeekStartDate.getDate() + 10 * 7); // 10 weeks after the PI start date

  // Adjust to start on the Saturday following the end of the last sprint
  planningWeekStartDate.setDate(planningWeekStartDate.getDate() + 7); // Move to the day after Innovation Week ends
  if (planningWeekStartDate.getDay() != 6) {
    planningWeekStartDate.setDate(planningWeekStartDate.getDate() + (6 - planningWeekStartDate.getDay()));
  }

  var planningWeekEndDate = new Date(planningWeekStartDate);
  planningWeekEndDate.setDate(planningWeekEndDate.getDate() + 6); // 1 week later

  createEvent("Next PI Planning Week", planningWeekStartDate, planningWeekEndDate);
  createPIPlanningEvents(planningWeekStartDate, nextFY, nextPI);
}

// Function to create the three PI planning week events
function createPIPlanningEvents(planningWeekStartDate, nextFY, nextPI) {
  // PI Planning - Welcome on Tuesday
  var welcomeDate = new Date(planningWeekStartDate);
  welcomeDate.setDate(welcomeDate.getDate() + 3); // Move to Tuesday
  var welcomeStartTime = new Date(welcomeDate.getFullYear(), welcomeDate.getMonth(), welcomeDate.getDate(), 10, 0); // 10 AM CT
  var welcomeEndTime = new Date(welcomeDate.getFullYear(), welcomeDate.getMonth(), welcomeDate.getDate(), 11, 0); // 11 AM CT
  createEvent("PI " + nextFY + "." + nextPI + " Planning - Welcome", welcomeStartTime, welcomeEndTime);

  // PI Planning - Management Review on Thursday
  var managementReviewDate = new Date(planningWeekStartDate);
  managementReviewDate.setDate(managementReviewDate.getDate() + 5); // Move to Thursday
  var managementReviewStartTime = new Date(managementReviewDate.getFullYear(), managementReviewDate.getMonth(), managementReviewDate.getDate(), 9, 30); // 9:30 AM CT
  var managementReviewEndTime = new Date(managementReviewDate.getFullYear(), managementReviewDate.getMonth(), managementReviewDate.getDate(), 12, 0); // 12 PM CT
  createEvent("PI " + nextFY + "." + nextPI + " Planning - Management Review", managementReviewStartTime, managementReviewEndTime);

  // PI Planning - Final Presentation on Friday
  var finalPresentationDate = new Date(planningWeekStartDate);
  finalPresentationDate.setDate(finalPresentationDate.getDate() + 6); // Move to Friday
  var finalPresentationStartTime = new Date(finalPresentationDate.getFullYear(), finalPresentationDate.getMonth(), finalPresentationDate.getDate(), 10, 0); // 10 AM CT
  var finalPresentationEndTime = new Date(finalPresentationDate.getFullYear(), finalPresentationDate.getMonth(), finalPresentationDate.getDate(), 11, 0); // 11 AM CT
  createEvent("PI " + nextFY + "." + nextPI + " Planning - Final Presentation", finalPresentationStartTime, finalPresentationEndTime);
}

// Function to create a Google Calendar event
function createEvent(eventName, startDate, endDate) {
  var calendar = CalendarApp.getCalendarById(calendarId);
  calendar.createEvent(eventName, startDate, endDate);
}

/////////////Recolor Events////////////////

// Helper function to fetch events within a specified date range
function getEventsWithinDateRange(startDate, endDate) {
  var calendar = CalendarApp.getCalendarById(calendarId);
  return calendar.getEvents(startDate, endDate);
}

// Helper function to define the color map for event titles
function getColorMap() {
  return {
    'Sprint 1': '10', // Basil color
    'Sprint 2': '2', // Sage color
    'Sprint 3': '5', // Banana color
    'Sprint 4': '6', // Tangerine color
    'Sprint 5': '11', // Tomato color
    'Innovation Week': '9', // Grape color
    'Next PI Planning': '3', // Blueberry color
    'IMPACT PI Planning Welcome': '10', // Basil color
    'IMPACT PI Planning Management Review': '10', // Basil color
    'IMPACT PI Planning Final Presentation': '10', // Basil color
  };
}

// function to change event colors based on their titles
function changeEventColors() {
  var today = new Date();
  var oneYearAgo = new Date(today.getTime() - 700 * 24 * 60 * 60 * 1000); // 365 days ago 
  var oneYearFuture = new Date(today.getTime() + 700 * 24 * 60 * 60 * 1000); // 365 days in the future
  var events = getEventsWithinDateRange(oneYearAgo, oneYearFuture);

  var colorMap = getColorMap();

  for (var i = 0; i < events.length; i++) {
    var event = events[i];
    var title = event.getTitle();
    var currentColor = event.getColor();

    for (var keyword in colorMap) {
      if (title.includes(keyword) && currentColor != colorMap[keyword]) {
        event.setColor(colorMap[keyword]);
        Logger.log('Changed color of event: ' + title);
      }
    }
  }
}
