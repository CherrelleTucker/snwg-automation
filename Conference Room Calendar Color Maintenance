// Purpose: automatically searches week/month ahead to change group calendar events to match their meeting room colors. Flag meetings that are occuring in a space that has been double booked.
// Done: account for variations of spelling and location entry "key phrase"
// Done: Meeting colors change to Graphite when locations are double booked. 
// Done: Events with no location are assigned color Flamingo.
// Color key: format = Google color assignment/ google color name  = "Key phrases in room name" 
  // 1/Lavender = color is unassigned
  // 2/Sage = color is unassigned
  // 3/Grape =  "McCarthy" "1063"
  // 4/Flamingo = no location listed in event or Microsoft Teams
  // 5/Banana =  "1062" "Lovelace"; 
  // 6/Tangerine =  "Hamilton" "1063A"; 
  // 7/Peacock = "3098" "3rd"  
  // 8/Graphite  = double booked
  // 9/Blueberry = "3084" "CR3084"
  // 10/Basil =  "1030" "Turing"; 
  // 11/Tomato = "1034"

//Final calendar ID: mn9msmqj2nqobs0md4gmgfnabk@group.calendar.google.com

function ColorEvents() {
  var today = new Date();
  var nextweek = new Date();
  nextweek.setDate(nextweek.getDate() + 7);

  var colorMap = {
    '3': ["mccarthy", "1063"], // Grape
    '5': ["1062", "lovelace"], // Banana
    '6': ["hamilton", "1063a"], // Tangerine
    '7': ["3098", "3rd","CR3098","CR3098:"], // Peacock
    '9': ["3084","CR3084"], // Blueberry
    '10': ["1030", "turing"], // Basil
    '11': ["1034"], // Tomato
  };

  var calendar = CalendarApp.getCalendarById('mn9msmqj2nqobs0md4gmgfnabk@group.calendar.google.com');
  var events = calendar.getEvents(today, nextweek);

  var locationEventsMap = {};

  for (var j = 0; j < events.length; j++) {
    var e = events[j];
    var locationWords = e.getLocation().toLowerCase().trim().split(/\s+/);

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
  
  // Check overlaps within the same colored events
  for (var color in locationEventsMap) {
    var sameColorEvents = locationEventsMap[color];

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
