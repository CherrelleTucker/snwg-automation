// function: to poulate PI increments and PI events for the 23.4, 24.1, and 24.2 project increments in a public Google Calendar. On the Monday of a PI Planning Week, the next available PI (quarter) generates and populates. 

// feature suggestion: slackbot for posting links; convert and delete all Q references to PI.

// calculate PI event information given the following:
  // 23.4 PI start: 2023-07-24
  // FY = Fiscal year; Oct 1 - Sept 30; 
  // PI = project increment = 12 weeks long, also 1 Quarter (Q)
  // SP = Sprint = two weeks long; There are 6 sprints per PI = 5 regular 2-week sprints + "Innovation Week" as sprint 6 week 1 (Flex week title 1) + "Next PI Planning Week" as sprint 6 week 2 (flex week title 2). These 2 weeks are refered to as "Flex Weeks"
// sprintNumber = FY.PI.S = FY.Q.S
// piNumber = FY.PI = FY.Q
// Sprint 1-5 PI label format: " "PI" FY.Q. "Sprint" S"
// Flex week label format: " "PI" FY.Q. "<flex week title 1 or 2>"
// access target calendar; Test calendar ID: c_8798ebb71e4f29ffc300845dabe847152b8c92e2afd062e0e31242d7fce780cd@group.calendar.google.com
// create events for each sprint (Sprint Reviews) and Next PI Planning week (Welcome, Management Review, and Final Presentation)
// 
// once event is created, comanion script changeColor.gs runs
// there occasionally adjustment weeks to account for office holidays 

//primary function to create events denoting 2 week long sprint weeks, one weeek long Innovation, and one week long Next PI planning
function createSprintsAndFlexWeeks() {
  var calendarId = "c_8798ebb71e4f29ffc300845dabe847152b8c92e2afd062e0e31242d7fce780cd@group.calendar.google.com"; // update to preferred calendar ID

  // Array of PIs with their start dates and adjustment weeks
  var pis = [
    { label: "23.4", startDate: '2023-07-24', adjustmentWeeks: 0 }, // PI 23.4
    { label: "24.1", startDate: '2023-10-23' }, // PI 24.1
    { label: "24.2", startDate: '2024-01-22' } // PI 24.2
  ];

  // Calculate the next two PI numbers
  var lastPi = pis[pis.length - 1];
  var lastStartDate = new Date(lastPi.startDate);
  for (var i = 0; i < 2; i++) {
    var nextStartDate = new Date(lastStartDate.getFullYear(), lastStartDate.getMonth() + 3, lastStartDate.getDate());
    var nextPiLabel = getNextPINumber(lastPi.label); // Get the label for the next PI
    pis.push({ label: nextPiLabel, startDate: nextStartDate.toISOString(), adjustmentWeeks: 0 }); // No assumption of adjustment weeks
    lastStartDate = nextStartDate;
  }

  for (var i = 0; i < pis.length; i++) {
    var pi = pis[i];
    var piStartDate = new Date(pi.startDate);

    // Create sprints, flex weeks, Sprint Review events, and PI Planning events
    createSprints(piStartDate, calendarId, pi.label);
    createFlexWeeks(piStartDate, calendarId, pi.label);
    createSprintReviewEvents(piStartDate, calendarId, pi.label);
    createPIPlanningEvents(piStartDate, calendarId, pi.label);

    // Calculate the next PI's start date by adding 3 months to the current PI's start date
    // and adjusting for the adjustment weeks
    var nextStartDate = new Date(piStartDate.getFullYear(), piStartDate.getMonth() + 3, piStartDate.getDate() + pi.adjustmentWeeks * 7);
  }
}

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

    // Get formatted Sprint label
    var sprintLabel = "Sprint " + i;

    // Create event title
    var eventTitle = "Sprint Review - PI " + piLabel + " " + sprintLabel;

    createEvent(eventTitle, reviewStartDateTime, reviewEndDateTime, calendarId);
  }
}

// Helper function to format the sprint label
function getFormattedSprintLabel(piLabel, sprintNumber) {
  var fiscalYear = getFiscalYear();
  var fiscalQuarter = getFiscalQuarter();
  return "PI " + piLabel + " Sprint " + sprintNumber;
}

// helper function to create PI Planning events
function createPIPlanningEvents(piStartDate, calendarId, piLabel) {
  var flexWeekDuration = 7; // 7 days

  // Calculate the start date of the PI Planning week
  var flexWeekStartDate2 = new Date(piStartDate.getTime() + 5 * 14 * 24 * 60 * 60 * 1000);

  // Event details
  var eventDetails = [
    { title: "IMPACT PI Planning Welcome", dayOfWeek: 2, startTime: 10, endTime: 11.5 },
    { title: "IMPACT PI Planning Management Review", dayOfWeek: 4, startTime: 9.5, endTime: 12 },
    { title: "IMPACT PI Planning Final Presentation", dayOfWeek: 5, startTime: 10, endTime: 11.5 }
  ];

  // Create the events
  eventDetails.forEach(function(eventDetail) {
    var eventDate = new Date(flexWeekStartDate2.getFullYear(), flexWeekStartDate2.getMonth(), flexWeekStartDate2.getDate() + eventDetail.dayOfWeek);
    var eventStartDateTime = new Date(eventDate.getFullYear(), eventDate.getMonth(), eventDate.getDate(), eventDetail.startTime);
    var eventEndDateTime = new Date(eventDate.getFullYear(), eventDate.getMonth(), eventDate.getDate(), eventDetail.endTime);

    var eventTitle = "PI " + piLabel + " " + eventDetail.title;

    createEvent(eventTitle, eventStartDateTime, eventEndDateTime, calendarId);
  });
}
  
// Helper function to calculate the next PI number
function getNextPINumber(currentPiLabel) {
  var currentQuarter = parseInt(currentPiLabel.split('.')[1]);
  var currentYear = parseInt(currentPiLabel.split('.')[0]);

  var nextQuarter = currentQuarter % 4 + 1;
  var nextYear = (currentQuarter === 4) ? currentYear + 1 : currentYear;

  return nextYear + "." + nextQuarter;
}

// helper function to title events
function getFormattedSprintLabel(piLabel, sprintNumber) {
  var fiscalYear = getFiscalYear(piLabel);
  var quarter = getFiscalQuarter(piLabel);
  return "PI " + fiscalYear + "." + quarter + " Sprint " + sprintNumber;
}
  
// helper function to create sprints
function createSprints(piStartDate, calendarId, piLabel) {
  var sprintDuration = 14; // 14 days
  var numSprints = 5;

  for (var i = 1; i <= numSprints; i++) {
    var sprintStartDate = new Date(piStartDate.getTime() + (i - 1) * sprintDuration * 24 * 60 * 60 * 1000);
    var sprintEndDate = new Date(sprintStartDate.getTime() + (sprintDuration - 1) * 24 * 60 * 60 * 1000);
    var sprintLabel = "PI " + piLabel + " Sprint " + i;

    createEvent(sprintLabel, sprintStartDate, sprintEndDate, calendarId);
  }
}

// helper function to create flex weeks 
function createFlexWeeks(piStartDate, calendarId, piLabel) {
  var flexWeekDuration = 7; // 7 days
  var flexWeekTitle1 = "Innovation Week";
  var flexWeekTitle2 = "Next PI Planning";

  var flexWeekStartDate1 = new Date(piStartDate.getTime() + 5 * 14 * 24 * 60 * 60 * 1000);
  var flexWeekEndDate1 = new Date(flexWeekStartDate1.getTime() + (flexWeekDuration - 1) * 24 * 60 * 60 * 1000);
  createEvent(flexWeekTitle1, flexWeekStartDate1, flexWeekEndDate1, calendarId);

  var flexWeekStartDate2 = new Date(flexWeekEndDate1.getTime() + 24 * 60 * 60 * 1000);
  var flexWeekEndDate2 = new Date(flexWeekStartDate2.getTime() + (flexWeekDuration - 1) * 24 * 60 * 60 * 1000);
  createEvent(flexWeekTitle2, flexWeekStartDate2, flexWeekEndDate2, calendarId);
}
  
// helper function to create events
function createEvent(eventTitle, startDate, endDate, calendarId) {
  var calendar = CalendarApp.getCalendarById(calendarId);

  // Check if the event already exists
  var existingEvents = calendar.getEvents(startDate, endDate);
  for (var i = 0; i < existingEvents.length; i++) {
    var event = existingEvents[i];
    if (event.getTitle() === eventTitle && event.getStartTime().getTime() === startDate.getTime()) {
      // Event already exists, so we skip creating this event
      Logger.log("Event already exists: " + eventTitle);
      return;
    }
  }

  // If we've gotten here, the event doesn't exist and we can create it
  var event = calendar.createEvent(eventTitle, startDate, endDate);
  Logger.log("Created event: " + eventTitle);
}
  
// helper function to calculate fiscal quarter
function getFiscalQuarter(date) {
  var currentDate = new Date(date);
  var currentMonth = currentDate.getMonth();

  // Fiscal year starts in October. Adjust the month accordingly to get the correct quarter.
  var fiscalMonth = (currentMonth + 9) % 12;
  return Math.floor(fiscalMonth / 3) + 1;
}

// helper function to calculate fiscal year
function getFiscalYear(date) {
  var currentDate = new Date(date);
  var currentYear = currentDate.getFullYear();
  var currentMonth = currentDate.getMonth();

  // Fiscal year starts in October. If the current month is October or later, 
  // the fiscal year is the next calendar year. 
  if (currentMonth >= 9) {
    return String(currentYear + 1).slice(-2);
  } else {
    return String(currentYear).slice(-2);
  }
}

// Helper function to add adjustment weeks
function addAdjustmentWeeks(adjustmentStartDate, numWeeks, calendarId) {
  // Parse the adjustment start date
  var adjustmentStartDateTime = new Date(adjustmentStartDate);
  
  // Calculate the adjustment end date
  var adjustmentEndDateTime = new Date(adjustmentStartDateTime.getTime() + numWeeks * 7 * 24 * 60 * 60 * 1000);
  
  // Create an "Adjustment Week" event for the adjustment period
  createEvent("Adjustment Week", adjustmentStartDateTime, adjustmentEndDateTime, calendarId);
  
  // Get all events after the adjustment start date
  var calendar = CalendarApp.getCalendarById(calendarId);
  var futureEvents = calendar.getEvents(adjustmentStartDateTime, new Date('2099-12-31'));
  
  // Iterate over all future events, moving them by the number of adjustment weeks
  futureEvents.forEach(function(event) {
    var eventStartTime = event.getStartTime();
    var eventEndTime = event.getEndTime();
    
    // Calculate the new start and end times for the event
    var newEventStartTime = new Date(eventStartTime.getTime() + numWeeks * 7 * 24 * 60 * 60 * 1000);
    var newEventEndTime = new Date(eventEndTime.getTime() + numWeeks * 7 * 24 * 60 * 60 * 1000);
    
    // Update the event with the new start and end times
    event.setTime(newEventStartTime, newEventEndTime);
  });
}
  
// Helper function to format the sprint label
function getFormattedSprintLabel(piLabel, sprintNumber) {
  var fiscalYear = getFiscalYear();
  var fiscalQuarter = getFiscalQuarter();
  return "PI " + piLabel + " " + fiscalYear + "." + fiscalQuarter + " Sprint " + sprintNumber;
}

// Call the main function to create sprints and flex weeks
createSprintsAndFlexWeeks();
