// Purpose: This script is designed to automate the creation of routine presentation and Jamboard files required for IMPACT PI Planning Week with a trigger to execute one week prior to the next PI Welcome event on the IMPACT PI Calendar. 

// How to use:
//  Manual execution: change the "Select funtion to run" dropdown to "callLibrariesInOrder". Select Run.

///////////////////////////////////////////////////

// source calendar ID
var sourceCalendarId = 'xxxxxxxxxxxxxxxxxxxx@group.calendar.google.com';


// helper function to call the Welcome Library that creates the PI Welcome slide presentation
function callWelcomeLibrary() {
  welcomePIplanningIMPACT.createNewPresentation(); 
}

// helper function to call the Management Review Library that creates the PI Management Review slide presentation
function callMgtReviewLibrary() {
  managementReviewPiPlanningImpact.createNewPresentation(); 
}

// helper function to call the Final Presentation Library that creates the PI Final Presentation slide presentation, the PI kudos Jamboard, and the Start/Stop/Continue Jamboard
function callFinalPresentationLibrary() {
  finalPresentationPiPlanningImpact.createNewPresentation();
}

// helper function that calls the script to attach to attach the newly created files to their calendar events. 
function callAttachEventsLibrary() {
  attachFilesToEventsIMPACT.updateCalendarEvents();
}

// Primary function to call all libraries in order. 
function callLibrariesInOrder() {
  callWelcomeLibrary();
  callMgtReviewLibrary();
  callFinalPresentationLibrary();
  callAttachEventsLibrary();
}

// Trigger function to execute one week before the next PI Welcome event on the specific calendar
function executeOneWeekBeforeWelcomeEvent() {
  var sourceCalendar = CalendarApp.getCalendarById(sourceCalendarId);
  var today = new Date();
  var oneWeekFromToday = new Date(today.getTime() + 7 * 24 * 60 * 60 * 1000);
  var events = sourceCalendar.getEvents(today, oneWeekFromToday);

  // Find the next PI Welcome event
  var nextPIWelcomeEvent = null;
  for (var i = 0; i < events.length; i++) {
    var event = events[i];
    if (event.getTitle().toLowerCase().indexOf('pi welcome') !== -1) {
      nextPIWelcomeEvent = event;
      break;
    }
  }

  if (nextPIWelcomeEvent) {
    // Calculate the date one week before the PI Welcome event
    var oneWeekBeforePIWelcome = new Date(nextPIWelcomeEvent.getStartTime().getTime() - 7 * 24 * 60 * 60 * 1000);

    // Set up the time-driven trigger to call the libraries one week before the PI Welcome event
    ScriptApp.newTrigger('callLibrariesInOrder')
      .timeBased()
      .at(oneWeekBeforePIWelcome)
      .create();
  } else {
    Logger.log("No upcoming PI Welcome event found within the next week.");
  }
}
