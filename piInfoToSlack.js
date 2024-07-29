const SLACK_WEBHOOK_URL = "SLACK_WEBHOOK"; // Replace with your Slack webhook URL
const CALENDAR_ID = "CALENDAR_ID";

/**
 * Sends time-specific events for the current day to Slack.
 * Filters out all-day events and formats the remaining events with their descriptions and locations.
 */
function sendDailyTimeSpecificEvents() {
  const today = new Date();
  today.setHours(0, 0, 0, 0); // Start of the day
  const endOfDay = new Date(today);
  endOfDay.setHours(23, 59, 59, 999); // End of the day

  const events = CalendarApp.getCalendarById(CALENDAR_ID).getEvents(today, endOfDay);
  const timeSpecificEvents = events.filter(event => 
    !event.isAllDayEvent() &&
    (event.getStartTime().getTime() >= today.getTime() && event.getEndTime().getTime() <= endOfDay.getTime())
  );

  const message = buildEventMessage(timeSpecificEvents, "Today's PI events:", formatTimeSpecificEvent);

  if (message) {
    sendToSlack(message);
  }
}

/**
 * Sends all-day and multi-day events for the current week to Slack.
 * Identifies all-day and multi-day events, formats them, and includes their descriptions and locations.
 */
function sendWeeklyAllDayEvents() {
  const today = new Date();
  const monday = new Date(today);
  monday.setDate(today.getDate() - today.getDay() + 1); // Get the Monday of the current week
  monday.setHours(0, 0, 0, 0);
  const sunday = new Date(monday);
  sunday.setDate(monday.getDate() + 6); // Get the Sunday of the current week
  sunday.setHours(23, 59, 59, 999);

  const events = CalendarApp.getCalendarById(CALENDAR_ID).getEvents(monday, sunday);
  const allDayEvents = events.filter(event => (event.isAllDayEvent() || isMultiDayEvent(event, monday, sunday)) && event.getStartTime().getDay() !== 6);

  const currentWeekEvents = CalendarApp.getCalendarById(CALENDAR_ID).getEventsForDay(today);
  const currentWeekEvent = currentWeekEvents.find(event => event.getStartTime() < today && event.getEndTime() > today);

  let allEvents = allDayEvents;
  if (currentWeekEvent && !allDayEvents.includes(currentWeekEvent)) {
    allEvents.push(currentWeekEvent);
  }

  allEvents.reverse(); // Reverse the order of events

  const message = buildEventMessage(allEvents, `This week's PI events:`, formatWeeklyEvent);

  if (message) {
    sendToSlack(message);
  }
}

/**
 * Checks if an event spans multiple days.
 * 
 * @param {CalendarEvent} event - The event to check.
 * @param {Date} weekStart - The start date of the week.
 * @param {Date} weekEnd - The end date of the week.
 * @returns {boolean} True if the event is multi-day, otherwise false.
 */
function isMultiDayEvent(event, weekStart, weekEnd) {
  const startTime = event.getStartTime();
  const endTime = event.getEndTime();
  return (startTime < weekEnd && endTime > weekStart);
}

/**
 * Builds a formatted message string for a list of events.
 * 
 * @param {CalendarEvent[]} events - The list of events to format.
 * @param {string} header - The header to include in the message.
 * @param {function} formatEvent - The function to format individual events.
 * @returns {string} The formatted message string.
 */
function buildEventMessage(events, header, formatEvent) {
  if (events.length === 0) {
    return null;
  }

  const message = events.map(formatEvent).join("\n\n");

  return `${header}\n${message}`;
}

/**
 * Formats a time-specific event for inclusion in the message.
 * 
 * @param {CalendarEvent} event - The event to format.
 * @returns {string} The formatted event string.
 */
function formatTimeSpecificEvent(event) {
  const start = event.getStartTime();
  const end = event.getEndTime();
  let formattedEvent = `  • ${event.getTitle()} (${start.toTimeString().split(' ')[0]} - ${end.toTimeString().split(' ')[0]})`;

  return formattedEvent;
}

/**
 * Formats an all-day or multi-day event for inclusion in the message.
 * 
 * @param {CalendarEvent} event - The event to format.
 * @returns {string} The formatted event string.
 */
function formatWeeklyEvent(event) {
  const start = event.getStartTime();
  const end = event.getEndTime();
  const description = cleanHtml(event.getDescription());

  let formattedEvent;

  if (description) {
    formattedEvent = `• ${description}`;
  } else if (event.isAllDayEvent() || start.toDateString() !== end.toDateString()) {
    formattedEvent = `• ${event.getTitle()} (${start.toLocaleDateString('en-US', { month: 'short', day: 'numeric' })} - ${end.toLocaleDateString('en-US', { month: 'short', day: 'numeric' })})`;
  } else {
    formattedEvent = `• ${event.getTitle()} (${start.toLocaleDateString('en-US', { month: 'short', day: 'numeric' })} ${start.toLocaleTimeString('en-US', { hour: '2-digit', minute: '2-digit' })})`;
  }

  return formattedEvent;
}

/**
 * Cleans HTML content by removing tags and extra whitespace.
 * 
 * @param {string} html - The HTML content to clean.
 * @returns {string} The cleaned text content.
 */
function cleanHtml(html) {
  if (!html) return '';

  // Remove HTML tags and extra whitespace
  const plainText = html.replace(/<[^>]*>/g, '').replace(/\s+/g, ' ').trim();
  return plainText;
}

/**
 * Sends a message to Slack using the specified webhook URL.
 * 
 * @param {string} message - The message to send.
 */
function sendToSlack(message) {
  const payload = {
    text: message
  };

  const options = {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify(payload)
  };

  UrlFetchApp.fetch(SLACK_WEBHOOK_URL, options);
}

/**
 * Webhook handler to respond to Slack commands.
 * Parses the command text and searches for events with titles matching the query.
 * 
 * @param {object} e - The event parameter containing the request data.
 * @returns {object} The response to send back to Slack.
 */
function doPost(e) {
  try {
    const slackData = parseFormData(e.postData.contents);
    Logger.log(`slackData: ${JSON.stringify(slackData)}`); // Log the slackData for debugging
    const searchTerm = slackData.text.trim();
    Logger.log(`Command text: ${searchTerm}`); // Log the command text for debugging

    if (searchTerm) {
      const events = searchEventsByTitle(searchTerm);
      Logger.log(`Filtered events: ${JSON.stringify(events.map(event => event.getTitle()))}`); // Log the event titles for debugging
      const message = buildEventMessage(events, `Events matching "${searchTerm}":`, formatWeeklyEvent) || `No event found for ${searchTerm}`;

      return ContentService.createTextOutput(JSON.stringify({ text: message }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    return ContentService.createTextOutput(JSON.stringify({ text: "Invalid command. Please use the format: /picalendar [search term]" }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (error) {
    // Log the error for debugging
    Logger.log(error.toString());
    return ContentService.createTextOutput(JSON.stringify({ text: "An error occurred. Please try again." }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * Test function to mimic the doPost function and output logs to the Execution log.
 */
function testDoPost() {
  const testPayload = {
    token: "testToken",
    team_id: "T0001",
    team_domain: "example",
    channel_id: "C2147483705",
    channel_name: "test",
    user_id: "U2147483697",
    user_name: "Steve",
    command: "/picalendar",
    text: "24.1",
    response_url: "https://hooks.slack.com/commands/1234/5678",
    trigger_id: "13345224609.738474920.8088930838d88f008e0"
  };
  
  const e = {
    postData: {
      contents: Object.entries(testPayload).map(([key, value]) => `${encodeURIComponent(key)}=${encodeURIComponent(value)}`).join('&')
    }
  };

  try {
    const slackData = parseFormData(e.postData.contents);
    Logger.log(`slackData: ${JSON.stringify(slackData)}`); // Log the slackData for debugging
    const searchTerm = slackData.text.trim();
    Logger.log(`Command text: ${searchTerm}`); // Log the command text for debugging

    if (searchTerm) {
      const events = searchEventsByTitle(searchTerm);
      Logger.log(`Filtered events: ${JSON.stringify(events.map(event => event.getTitle()))}`); // Log the event titles for debugging
      const message = buildEventMessage(events, `Events matching "${searchTerm}":`, formatWeeklyEvent) || `No event found for ${searchTerm}`;
      Logger.log(`Message: ${message}`); // Log the message for debugging
    } else {
      Logger.log("Invalid command. Please use the format: /picalendar [search term]");
    }
  } catch (error) {
    // Log the error for debugging
    Logger.log(error.toString());
  }
}

/**
 * Parses URL-encoded form data into an object.
 * 
 * @param {string} data - The URL-encoded form data.
 * @returns {object} The parsed data.
 */
function parseFormData(data) {
  const result = {};
  const pairs = data.split('&');
  for (let i = 0; i < pairs.length; i++) {
    const pair = pairs[i].split('=');
    result[decodeURIComponent(pair[0])] = decodeURIComponent(pair[1] || '');
  }
  return result;
}

/**
 * Searches for events with titles matching the given search term.
 * 
 * @param {string} searchTerm - The term to search for in event titles.
 * @returns {CalendarEvent[]} The list of matching events.
 */
function searchEventsByTitle(searchTerm) {
  const startDate = new Date();
  startDate.setFullYear(startDate.getFullYear() - 1); // Search for events from one year ago
  const endDate = new Date(startDate);
  endDate.setFullYear(startDate.getFullYear() + 2); // Search for events up to one year in the future

  const events = CalendarApp.getCalendarById(CALENDAR_ID).getEvents(startDate, endDate);
  Logger.log(`Total events: ${events.length}`); // Log the total number of events for debugging
  Logger.log(`Event titles: ${events.map(event => event.getTitle()).join(', ')}`); // Log all event titles for debugging
  const regex = new RegExp(searchTerm, 'i'); // Create a case-insensitive regex for the search term
  return events.filter(event => regex.test(event.getTitle()));
}

// Set up triggers in the Apps Script UI:
// 1. Weekly trigger for sendWeeklyAllDayEvents() on Monday
// 2. Daily trigger for sendDailyTimeSpecificEvents() every day
