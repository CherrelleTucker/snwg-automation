// Purpose: 
// This script is a Google Apps Script designed to work with the IMPACT conference room Google Calendar. Its purpose is to recolor events on a specific calendar based on their locations. It does so by assigning different colors to events associated with specific room identifiers (such as 'mccarthy', '1063', 'teams', etc.) specified in the colorMap object. Meeting rooms that are double booked are flagged with the color Graphite.

// Future development: 
// none currently identified

// To note:
// This script is developed as a Google Apps Script standalone script. It is developed to operate independently and does not require any external application or service to function. It is a self-contained piece of code with a time-based daily trigger.

// To use: 
// Copy the entire script provided below.
// Open the Google Apps Script editor by going to "Tools" > "Script editor" in the Google Sheets document that corresponds to the Google Calendar you want to work with.
// In the Apps Script editor, paste the copied script into the script editor window.
// Save the script by clicking "File" > "Save" or using the keyboard shortcut "Ctrl + S" (or "Cmd + S" on Mac).
// Replace the placeholder value for CALENDAR_ID with the ID of your desired Google Calendar. You can find the Calendar ID by opening your Google Calendar, clicking on the three vertical dots next to the calendar name, selecting "Settings and sharing," and copying the "Calendar ID" under the "Integrate calendar" section. Make sure to keep the single quotes around the Calendar ID.
// Review the colorMap object and adjust the room identifiers and associated colors as needed. Each room identifier should have its corresponding color value specified. Room identifiers are case-insensitive.
// After making any changes, save the script again.
// Test the script by running it once. To do this, click on the function dropdown menu in the Apps Script editor (usually located at the top), and select "ColorEvents." Then, click the play button (▶️) to run the function. This will execute the script and recolor the events based on the specified room identifiers.
// Check your Google Calendar to see if the events have been recolored as intended. If everything looks good, you can proceed to set up a daily trigger to run the script automatically.
  // Set Up Daily Trigger:
    // To ensure the script runs automatically on a daily basis, follow these steps:
    // In the Apps Script editor, click on the "Triggers" icon (clock-shaped) located on the left-hand side.
    // Click on the "+ Add Trigger" button.
    // In the "Choose which function to run" dropdown, select "ColorEvents."
    // In the "Select event source" dropdown, choose "Time-driven."
    // In the "Select type of time based trigger" dropdown, choose "Day timer."
    // Choose the time you want the script to run daily (e.g., 2:00 AM).
    // Click "Save" to create the daily trigger.
 
///////////////////////////////////////////////////////////////////////

// Test calendar ID: c_8798ebb71e4f29ffc300845dabe847152b8c92e2afd062e0e31242d7fce780cd@group.calendar.google.com 
// IMPACT Conference room calendar ID: mn9msmqj2nqobs0md4gmgfnabk@group.calendar.google.com
// Global constant for Calendar ID
var CALENDAR_ID = 'mn9msmqj2nqobs0md4gmgfnabk@group.calendar.google.com'; //<---replace with desired calendar id

// Helper function to assign Google colors to room identifiers
function assignColorToLocation(locationWords) {
  var colorMap = {
    '3': ["mccarthy", "1063"], // Grape
    '4': ["teams"], // Flamingo
    '5': ["1062", "lovelace"], // Banana
    '6': ["hamilton", "1063a"], // Tangerine
    '7': ["3098", "3rd", "CR3098", "CR3098:"], // Peacock
    '9': ["3084", "CR3084"], // Blueberry
    '10': ["1030", "turing"], // Basil
    '11': ["1034"], // Tomato
  };

  for (var color in colorMap) {
    var phrases = colorMap[color];
    for (var k = 0; k < phrases.length; k++) {
      var phrase = phrases[k];
      if (locationWords.includes(phrase)) {
        return color;
      }
    }
  }
  return ''; // Return empty string if no color matches
}

// Helper function to check if an event is a virtual event based on the absence of room identifiers
function isVirtualEvent(event, colorMap) {
  var locationWords = event.getLocation().toLowerCase().trim().split(/\s+/);
  for (var color in colorMap) {
    var phrases = colorMap[color];
    for (var k = 0; k < phrases.length; k++) {
      var phrase = phrases[k];
      if (locationWords.includes(phrase)) {
        return false; // Event has a room identifier, not a virtual event
      }
    }
  }
  return true; // Event does not have any room identifiers, it is a virtual event
}

// Helper function to check for overlaps within the same colored events, skipping virtual events
function checkOverlapsForSameColorEvents(sameColorEvents, colorMap) {
  // Sort events based on their start time for efficient overlap checks
  sameColorEvents.sort(function(a, b) {
    return a.getStartTime() - b.getStartTime();
  });

  for (var j = 0; j < sameColorEvents.length - 1; j++) {
    var currentEvent = sameColorEvents[j];
    var nextEvent = sameColorEvents[j + 1];

    if (!isVirtualEvent(currentEvent, colorMap) && !isVirtualEvent(nextEvent, colorMap) && currentEvent.getEndTime() > nextEvent.getStartTime()) {
      currentEvent.setColor('8');
      nextEvent.setColor('8');

    // Log the overlapping events
    Logger.log("Overlap detected between the events: '" + currentEvent.getTitle() + "' and '" + nextEvent.getTitle() + "'. Both set to color '8'");

    }
  }
}

// Primary function to recolor events
function ColorEvents() {
  var today = new Date();
  var nextweek = new Date();
  nextweek.setDate(nextweek.getDate() + 7); 

  var calendar = CalendarApp.getCalendarById(CALENDAR_ID);  
  var events = calendar.getEvents(today, nextweek);

  // Define the colorMap here
  var colorMap = {
    '3': ["mccarthy", "1063"], // Grape
    '4': ["teams"], // Flamingo
    '5': ["1062", "lovelace"], // Banana
    '6': ["hamilton", "1063a"], // Tangerine
    '7': ["3098", "3rd", "CR3098", "CR3098:"], // Peacock
    '9': ["3084", "CR3084"], // Blueberry
    '10': ["1030", "turing"], // Basil
    '11': ["1034"], // Tomato
  };

  var locationEventsMap = {};
  var eventsWithoutLocation = [];

for (var j = 0; j < events.length; j++) {
    var event = events[j];
    var location = event.getLocation().toLowerCase().replace('#', ''); // Remove the '#' character
    var locationWords = location.trim().split(/\s+/);
    var color = assignColorToLocation(locationWords);

    if (color !== '') {
      event.setColor(color);

      // Log the color assignment
      Logger.log("Event: '" + event.getTitle() + "' set to color '" + color + "'");

      if (!locationEventsMap[color]) {
        locationEventsMap[color] = [];
      }
      locationEventsMap[color].push(event);
    } else {
      eventsWithoutLocation.push(event);
    }
  }

  // Set color to Flamingo for events without a location after checking for overlaps
  for (var j = 0; j < eventsWithoutLocation.length; j++) {
    var event = eventsWithoutLocation[j];
    event.setColor('4');
    
    // Log the color assignment for events without location
    Logger.log("Event without location: '" + event.getTitle() + "' set to color '4' (Flamingo)");
  }
  
  // Check overlaps within the same colored events
  for (var color in locationEventsMap) {
    var sameColorEvents = locationEventsMap[color];
    checkOverlapsForSameColorEvents(sameColorEvents, colorMap); // Pass colorMap here
  }
}
