// Purpose: automatically search the week ahead in the IMPACT conference room calendar to change calendar events to match their meeting room colors. Flag meetings that are occuring in a space that has been double booked with dark color.

//Primary function containing all other functions
function ColorEvents() {
  // look ahead 7 days
  var today = new Date();
  var nextweek = new Date();
  nextweek.setDate(nextweek.getDate() + 7); 

  //Assign Google colors to room identifiers
  var colorMap = {
    '3': ["mccarthy", "1063"], // Grape
    '4': ["teams"], // Flamingo
    '5': ["1062", "lovelace"], // Banana
    '6': ["hamilton", "1063a"], // Tangerine
    '7': ["3098", "3rd","CR3098","CR3098:"], // Peacock
    '9': ["3084","CR3084"], // Blueberry
    '10': ["1030", "turing"], // Basil
    '11': ["1034"], // Tomato
  };

  // Get calendar with events to change
  var calendar = CalendarApp.getCalendarById('mn9msmqj2nqobs0md4gmgfnabk@group.calendar.google.com');  //<---replace with desired calendar id
  var events = calendar.getEvents(today, nextweek);

  var locationEventsMap = {};

  for (var j = 0; j < events.length; j++) {
    var e = events[j];
    var locationWords = e.getLocation().toLowerCase().trim().split(/\s+/);

    // change events to their assigned color
    var eventColor = '';
    for (var color in colorMap) {
      var phrases = colorMap[color];
      for (var k = 0; k < phrases.length; k++) {
        var phrase = phrases[k];
        if (locationWords.includes(phrase)) {
          e.setColor(color);
          eventColor = color;
          break;
        }
      }
      if (eventColor) break;
    }

    if (!locationEventsMap[eventColor]) {
      locationEventsMap[eventColor] = [];
    }
    locationEventsMap[eventColor].push(e);
  }

  // Set color to Flamingo for events without a location after checking for overlaps
  for (var j = 0; j < events.length; j++) {
    var e = events[j];
    var location = e.getLocation().toLowerCase().trim();
    if (location == "") {
      e.setColor('4');
    }
  }
  
  // Check overlaps within the same colored events, skipping virtual events
  for (var color in locationEventsMap) {
    var sameColorEvents = locationEventsMap[color];
    if (color === '4') continue; // Skip overlap check for Flamingo events

    // Sort events based on their start time for efficient overlap checks
    sameColorEvents.sort(function(a, b) {
      return a.getStartTime() - b.getStartTime();
    });

    for (var j = 0; j < sameColorEvents.length - 1; j++) {
      var currentEvent = sameColorEvents[j];
      var nextEvent = sameColorEvents[j + 1];

      if (currentEvent.getEndTime() > nextEvent.getStartTime()) {
        currentEvent.setColor('8');
        nextEvent.setColor('8');
 
      }
    }
  }
}
