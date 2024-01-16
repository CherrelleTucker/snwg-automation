// function: to poulate PI increments and PI events for the 23.4, 24.1, and 24.2 project increments in a public Google Calendar. On the Monday of a PI Planning Week, the next available PI (quarter) generates and populates. 

// feature suggestion: slackbot for posting links

// access target calendar; Test calendar ID: c_8798ebb71e4f29ffc300845dabe847152b8c92e2afd062e0e31242d7fce780cd@group.calendar.google.com

// create events for each sprint (Sprint Reviews) and Next PI Planning week (Welcome, Management Review, and Final Presentation)

///////////////////////////////////////

// Global variables
// var calendarId = "c_8798ebb71e4f29ffc300845dabe847152b8c92e2afd062e0e31242d7fce780cd@group.calendar.google.com"; // Test calendar ID; update to preferred calendar ID before deployment

// immediately invoked function expression - adjustedPIcalculator to calculate current PI  ****DO NOT TOUCH****
var adjustedPIcalculator = (function(){
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

  // Helper function to calculate the Adjusted PI
  function getPI(inputDate) {
    if (!(inputDate instanceof Date)) {
      // Attempt to parse the inputDate as a string and convert it to a Date object
      inputDate = new Date(inputDate);

      // Check if the parsed inputDate is valid
      if (isNaN(inputDate.getTime())) {
        throw new Error("Invalid date provided to getPI function.");
      }
    }

    var daysPassed = (toMillis(inputDate) - toMillis(BASE_DATE)) / (24 * 60 * 60 * 1000) - adjustmentWeeks(inputDate) * 7;
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

  // Primary function to replace placeholder text {{Adjusted PI}} with the calculated Adjusted PI
  function replacePlaceholderWithPI(document, adjustedPI) {
    var body = document.getBody();
    var text = body.editAsText();
    text.replaceText("{{Adjusted PI}}", adjustedPI);
  }

  // Make the getPI function accessible outside the IIFE
  return {
    getPI: getPI,
    replacePlaceholderWithPI: replacePlaceholderWithPI
  };
})();

// Helper function to create Sprint Review events
function createSprintReviewEvents(piStartDate, calendarId, piLabel) {
  var sprintDuration = 14; // 14 days
  var numSprints = 5;
  var reviewHoursStart = 10; // 10am
  var reviewHoursEnd = 12; // 12pm

  for (var i = 1; i <= numSprints; i++) {
    var sprintStartDate = new Date(piStartDate.getTime() + (i - 1) * sprintDuration * 24 * 60 * 60 * 1000);
    var sprintEndDate = new Date(sprintStartDate.getTime() + (sprintDuration - 1) * 24 * 60 * 60 * 1000);

    // Adjusting the endDate to fall on a Friday
    var reviewDate = new Date(sprintEndDate);
    reviewDate.setDate(sprintEndDate.getDate() - (sprintEndDate.getDay() + 2) % 7);

    // Create start and end date objects for the Sprint Review event
    var reviewStartDateTime = new Date(reviewDate.getFullYear(), reviewDate.getMonth(), reviewDate.getDate(), reviewHoursStart);
    var reviewEndDateTime = new Date(reviewDate.getFullYear(), reviewDate.getMonth(), reviewDate.getDate(), reviewHoursEnd);

    var sprintLabel = "PI " + piLabel + " Sprint " + i; // Updated event label

    // Update the event title to include only "PI FY.Q Sprint S"
    createEvent(sprintLabel, reviewStartDateTime, reviewEndDateTime, calendarId);
  }
}

// Helper function to create PI Planning events
function createPIPlanningEvents(piStartDate, calendarId, piLabel) {
  var flexWeekDuration = 7; // 7 days

  for (var i = 1; i <= 2; i++) { // Assuming 2 PI Planning weeks per PI
    var piPlanningStartDate = new Date(piStartDate.getTime() + (i * 6 + 5) * 14 * 24 * 60 * 60 * 1000);
    var piPlanningEndDate = new Date(piPlanningStartDate.getTime() + (flexWeekDuration - 1) * 24 * 60 * 60 * 1000);

    var piPlanningLabel = "PI " + piLabel + " PI Planning Week " + i;

    // Event details
    var eventDetails = [
      { title: "IMPACT PI Planning Welcome", dayOfWeek: 2, startTime: 10, endTime: 11.5 },
      { title: "IMPACT PI Planning Management Review", dayOfWeek: 4, startTime: 9.5, endTime: 12 },
      { title: "IMPACT PI Planning Final Presentation", dayOfWeek: 5, startTime: 10, endTime: 11.5 }
    ];

    // Create the events for the PI Planning week
    eventDetails.forEach(function (eventDetail) {
      var eventDate = new Date(piPlanningStartDate.getFullYear(), piPlanningStartDate.getMonth(), piPlanningStartDate.getDate() + eventDetail.dayOfWeek);
      var eventStartDateTime = new Date(eventDate.getFullYear(), eventDate.getMonth(), eventDate.getDate(), eventDetail.startTime);
      var eventEndDateTime = new Date(eventDate.getFullYear(), eventDate.getMonth(), eventDate.getDate(), eventDetail.endTime);

      var eventTitle = "PI " + piLabel + " " + eventDetail.title;

      createEvent(eventTitle, eventStartDateTime, eventEndDateTime, calendarId);
    });

    // Create the PI Planning week itself
    createEvent(piPlanningLabel, piPlanningStartDate, piPlanningEndDate, calendarId);
  }
}

// Helper function to create an event in Google Calendar
function createEvent(eventTitle, eventStartDateTime, eventEndDateTime, calendarId) {
  // Create an object for the new event
  var event = {
    'summary': eventTitle,
    'start': {
      'dateTime': eventStartDateTime.toISOString(),
      'timeZone': 'America/New_York'  // Adjust timezone accordingly
    },
    'end': {
      'dateTime': eventEndDateTime.toISOString(),
      'timeZone': 'America/New_York'  // Adjust timezone accordingly
    }
  };

  // Add the event to the calendar
  var calendar = CalendarApp.getCalendarById(calendarId);
  var createdEvent = calendar.createEvent(eventTitle, eventStartDateTime, eventEndDateTime, {
    'description': eventTitle,
    'location': 'Location', // Set your desired location
  });
  Logger.log('Created event with ID: ' + createdEvent.getId());
}

// Primary function to create events for each PI increment
function populatePIEvents() {
  var calendarId = "c_8798ebb71e4f29ffc300845dabe847152b8c92e2afd062e0e31242d7fce780cd@group.calendar.google.com"; // Test calendar ID; update to preferred calendar ID before deployment

  // Get the current Adjusted PI using the adjustedPIcalculator library
  var currentPI = adjustedPIcalculator.getPI(new Date());

  // Calculate the next PI
  var nextAdjustedPI = getNextAdjustedPI(currentPI);

  // Get the start date for the next PI
  var nextStartDate = getStartDateForAdjustedPI(nextAdjustedPI);

  // Create sprints, flex weeks, Sprint Review events, and PI Planning events for the next PI
  createSprints(nextStartDate, calendarId, nextAdjustedPI);
  createFlexWeeks(nextStartDate, calendarId, nextAdjustedPI);
  createSprintReviewEvents(nextStartDate, calendarId, nextAdjustedPI);
  createPIPlanningEvents(nextStartDate, calendarId, nextAdjustedPI);
}