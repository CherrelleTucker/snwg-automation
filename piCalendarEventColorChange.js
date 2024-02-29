// Purpose:  to change the color of events in a Google Calendar based on their titles. It fetches events within a specified date range, applies a color map to match event titles with corresponding colors, and updates the events' colors accordingly. The script allows for easy categorization and visualization of events with specific keywords in the event title, helping users quickly identify and differentiate events based on their significance or type. To run after populatatePIEvents.gs is fully executed.

///////////////////////////////////////////////

// Global variable
var calendarId = 'xxxxxxxxxxxxxxxxxxxxxx@group.calendar.google.com';  // IMPACT PI calendar

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

// Primary function to change event colors based on their titles
function changeEventColors() {
  var today = new Date();
  var oneYearAgo = new Date(today.getTime() - 365 * 24 * 60 * 60 * 1000); // 365 days ago 
  var oneYearFuture = new Date(today.getTime() + 365 * 24 * 60 * 60 * 1000); // 365 days in the future
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
