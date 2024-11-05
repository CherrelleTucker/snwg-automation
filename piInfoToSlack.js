/**
 * SlackCalendarBot Class to interact with Google Calendar and send event details to Slack.
 * 
 * This class fetches events from a specified calendar and posts relevant information 
 * to a Slack channel using a webhook URL.
 */
const SCRIPT_PROPERTIES = PropertiesService.getScriptProperties();
const SLACK_WEBHOOK_URL = SCRIPT_PROPERTIES.getProperty('SLACK_WEBHOOK_URL');
const CALENDAR_ID = "c_e6e532cefc5ddfdd7f3c715e7a07326607cd240d951991f6a4e3b87653e67ef3@group.calendar.google.com";

/**
 * SlackCalendarBot class handles the Google Calendar events and Slack notifications.
 */
class SlackCalendarBot {
  /**
   * Constructor to initialize SlackCalendarBot instance.
   * @param {string} calendarId - The ID of the Google Calendar.
   * @param {string} slackWebhookUrl - The Slack Webhook URL to send messages.
   */
  constructor(calendarId, slackWebhookUrl) {
    this.calendar = CalendarApp.getCalendarById(calendarId);
    this.slackWebhookUrl = slackWebhookUrl;
  }

  /**
   * Send events happening today that have specific start and end times.
   */
  sendDailyTimeSpecificEvents() {
    const today = new Date();
    today.setHours(0, 0, 0, 0); // Start of the day
    const endOfDay = new Date(today);
    endOfDay.setHours(23, 59, 59, 999); // End of the day

    const events = this.calendar.getEvents(today, endOfDay);
    // Filter events to exclude all-day events and ignored events
    const timeSpecificEvents = events.filter(event => 
      !event.isAllDayEvent() &&
      !this.isIgnoredEvent(event) &&
      (event.getStartTime().getTime() >= today.getTime() && event.getEndTime().getTime() <= endOfDay.getTime())
    );

    const message = this.buildEventMessage(timeSpecificEvents, "Today's PI events:", this.formatTimeSpecificEvent);

    if (message) {
      this.sendToSlack(message);
    }
  }

  /**
   * Send all-day or multi-day events happening this week.
   */
  sendWeeklyAllDayEvents() {
    const today = new Date();
    const monday = this.getMondayOfCurrentWeek(today);
    const sunday = this.getSundayOfCurrentWeek(monday);

    const events = this.calendar.getEvents(monday, sunday);
    // Filter events to include only all-day events or multi-day events, excluding ignored events
    const allDayEvents = events.filter(event => (event.isAllDayEvent() || this.isMultiDayEvent(event, monday, sunday)) && !this.isIgnoredEvent(event));

    const message = this.buildEventMessage(allDayEvents, `This week's PI events:`, this.formatWeeklyEvent);

    if (message) {
      this.sendToSlack(message);
    }
  }

  /**
   * Get the document descriptions for the most recent past event and the next upcoming event.
   * @returns {string} - A message containing the past and upcoming event descriptions.
   */
  getDocsForRecentAndUpcomingEvents() {
    const now = new Date();
    const events = this.calendar.getEvents(new Date(now.getFullYear() - 1, now.getMonth(), now.getDate()), new Date(now.getFullYear() + 1, now.getMonth(), now.getDate()));
    
    // Filter events that are not all-day and have a duration of 4 hours or less, excluding ignored events
    const timeSpecificEvents = events.filter(event => {
      const startTime = event.getStartTime();
      const endTime = event.getEndTime();
      const duration = (endTime - startTime) / (1000 * 60 * 60); // Duration in hours
      return !event.isAllDayEvent() && duration <= 4 && !this.isIgnoredEvent(event);
    });
    
    // Get the most recent past event and the next future event
    const pastEvents = timeSpecificEvents.filter(event => event.getEndTime() < now).sort((a, b) => b.getEndTime() - a.getEndTime());
    const futureEvents = timeSpecificEvents.filter(event => event.getStartTime() > now).sort((a, b) => a.getStartTime() - b.getStartTime());
    
    const pastEvent = pastEvents.length > 0 ? pastEvents[0] : null;
    const futureEvent = futureEvents.length > 0 ? futureEvents[0] : null;

    const pastEventDescription = pastEvent ? this.cleanHtml(pastEvent.getDescription()) || "This file has not yet been created" : "This file has not yet been created";
    const futureEventDescription = futureEvent ? this.cleanHtml(futureEvent.getDescription()) || "This file has not yet been created" : "This file has not yet been created";
    
    let message = "Most Recent Past Event:\n";
    message += pastEvent ? `• ${pastEventDescription}` : "No past event found.";
    
    message += "\n\nNext Event:\n";
    message += futureEvent ? `• ${futureEventDescription}` : "No future event found.";
    
    return message;
  }

  /**
   * Get multi-day events happening today.
   * @returns {string} - A message containing current multi-day events.
   */
  getCurrentMultiDayEvents() {
    const today = new Date();
    today.setHours(0, 0, 0, 0); // Start of the day
    const endOfDay = new Date(today);
    endOfDay.setHours(23, 59, 59, 999); // End of the day

    const events = this.calendar.getEvents(today, endOfDay);
    // Filter events to include only multi-day events, excluding ignored events
    const multiDayEvents = events.filter(event => this.isMultiDayEvent(event, today, endOfDay) && !this.isIgnoredEvent(event));

    let message = "Current PI & sprint:\n";
    if (multiDayEvents.length > 0) {
      message += multiDayEvents.map(event => `• ${event.getTitle()}`).join("\n");
    } else {
      message += "No multi-day events found for today.";
    }

    return message;
  }

  /**
   * Determine if an event is a multi-day event.
   * @param {Event} event - The event to check.
   * @param {Date} start - The start date of the time period.
   * @param {Date} end - The end date of the time period.
   * @returns {boolean} - True if the event is a multi-day event, false otherwise.
   */
  isMultiDayEvent(event, start, end) {
    const startTime = event.getStartTime();
    const endTime = event.getEndTime();
    return (startTime < end && endTime > start) || (endTime.getDate() !== startTime.getDate());
  }

  /**
   * Determine if an event should be ignored based on its title.
   * @param {Event} event - The event to check.
   * @returns {boolean} - True if the event should be ignored, false otherwise.
   */
  isIgnoredEvent(event) {
    const title = event.getTitle().toLowerCase();
    return title.includes("resource risk") || title.includes("po sync");
  }

  /**
   * Build a message from a list of events.
   * @param {Array} events - The list of events to include in the message.
   * @param {string} header - The header text for the message.
   * @param {Function} formatEvent - A function to format each event.
   * @returns {string|null} - The formatted message or null if there are no events.
   */
  buildEventMessage(events, header, formatEvent) {
    if (events.length === 0) {
      return null;
    }

    const message = events.map(formatEvent).join("\n\n");

    return `${header}\n${message}`;
  }

  /**
   * Format a time-specific event for the message.
   * @param {Event} event - The event to format.
   * @returns {string} - The formatted event string.
   */
  formatTimeSpecificEvent(event) {
    const start = event.getStartTime();
    const end = event.getEndTime();
    return `  • ${event.getTitle()} (${start.toTimeString().split(' ')[0]} - ${end.toTimeString().split(' ')[0]})`;
  }

  /**
   * Format a weekly event for the message.
   * @param {Event} event - The event to format.
   * @returns {string} - The formatted event string.
   */
  formatWeeklyEvent(event) {
    const start = event.getStartTime();
    const end = event.getEndTime();
    let formattedEvent;

    if (event.isAllDayEvent() || start.toDateString() !== end.toDateString()) {
      formattedEvent = `• ${event.getTitle()} (${start.toLocaleDateString('en-US', { month: 'short', day: 'numeric' })} - ${end.toLocaleDateString('en-US', { month: 'short', day: 'numeric' })})`;
    } else {
      formattedEvent = event.getTitle().startsWith("PI") 
        ? `• ${event.getTitle()} (${start.toLocaleDateString('en-US', { month: 'short', day: 'numeric' })} ${start.toLocaleTimeString('en-US', { hour: '2-digit', minute: '2-digit' })})`
        : `   ${event.getTitle()} (${start.toLocaleDateString('en-US', { month: 'short', day: 'numeric' })} ${start.toLocaleTimeString('en-US', { hour: '2-digit', minute: '2-digit' })})`;
    }

    return formattedEvent;
  }

  /**
   * Remove HTML tags and extra whitespace from a string.
   * @param {string} html - The HTML string to clean.
   * @returns {string} - The cleaned plain text.
   */
  cleanHtml(html) {
    if (!html) return '';

    // Remove HTML tags and extra whitespace
    const plainText = html.replace(/<[^>]*>/g, '').replace(/\s+/g, ' ').trim();
    return plainText;
  }

  /**
   * Send a message to Slack using the webhook URL.
   * @param {string} message - The message to send.
   */
  sendToSlack(message) {
    const payload = {
      text: message
    };

    const options = {
      method: "post",
      contentType: "application/json",
      payload: JSON.stringify(payload)
    };

    UrlFetchApp.fetch(this.slackWebhookUrl, options);
  }

  /**
   * Get the Monday of the current week.
   * @param {Date} today - The current date.
   * @returns {Date} - The Monday of the current week.
   */
  getMondayOfCurrentWeek(today) {
    const monday = new Date(today);
    monday.setDate(today.getDate() - today.getDay() + 1); // Get the Monday of the current week
    monday.setHours(0, 0, 0, 0);
    return monday;
  }

  /**
   * Get the Sunday of the current week.
   * @param {Date} monday - The Monday of the current week.
   * @returns {Date} - The Sunday of the current week.
   */
  getSundayOfCurrentWeek(monday) {
    const sunday = new Date(monday);
    sunday.setDate(monday.getDate() + 6); // Get the Sunday of the current week
    sunday.setHours(23, 59, 59, 999);
    return sunday;
  }

  /**
   * Get the current PI (Program Increment) number from multi-day events.
   * @returns {string|null} - The current PI number or null if not found.
   */
  getCurrentPI() {
    const today = new Date();
    today.setHours(0, 0, 0, 0); // Start of the day
    const endOfDay = new Date(today);
    endOfDay.setHours(23, 59, 59, 999); // End of the day

    const events = this.calendar.getEvents(today, endOfDay);
    const currentPIEvent = events.find(event => {
      const title = event.getTitle();
      return this.isMultiDayEvent(event, today, endOfDay) && /PI \d+\.\d+ Sprint \d+/.test(title) && !this.isIgnoredEvent(event);
    });

    if (currentPIEvent) {
      const match = currentPIEvent.getTitle().match(/PI (\d+\.\d+)/);
      return match ? match[1] : null;
    }

    return null;
  }

  /**
   * Handle incoming POST requests from Slack.
   * @param {object} e - The event object containing POST data.
   * @returns {ContentService.TextOutput} - The response to send back to Slack.
   */
  doPost(e) {
    try {
      const slackData = this.parseFormData(e.postData.contents);
      Logger.log(`slackData: ${JSON.stringify(slackData)}`); // Log the slackData for debugging
      const commandText = slackData.command.trim();
      const searchTerm = slackData.text.trim();
      Logger.log(`Command text: ${commandText}`); // Log the command text for debugging
      Logger.log(`Search term: ${searchTerm}`); // Log the search term for debugging

      if (commandText === '/picalendar') {
        let termToSearch = searchTerm;
        if (!termToSearch) {
          termToSearch = this.getCurrentPI();
          if (!termToSearch) {
            return ContentService.createTextOutput(JSON.stringify({ text: "No current PI found." }))
              .setMimeType(ContentService.MimeType.JSON);
          }
        }

        const events = this.searchEventsByTitle(termToSearch);
        Logger.log(`Filtered events: ${JSON.stringify(events.map(event => event.getTitle()))}`); // Log the event titles for debugging
        const message = this.buildEventMessage(events, `Events matching "${termToSearch}":`, this.formatWeeklyEvent) || `No event found for ${termToSearch}`;

        return ContentService.createTextOutput(JSON.stringify({ text: message }))
          .setMimeType(ContentService.MimeType.JSON);
      } else if (commandText === '/pidocs') {
        const message = this.getDocsForRecentAndUpcomingEvents();
        return ContentService.createTextOutput(JSON.stringify({ text: message }))
          .setMimeType(ContentService.MimeType.JSON);
      } else if (commandText === '/picurrent') {
        const message = this.getCurrentMultiDayEvents();
        return ContentService.createTextOutput(JSON.stringify({ text: message }))
          .setMimeType(ContentService.MimeType.JSON);
      } else {
        return ContentService.createTextOutput(JSON.stringify({ text: "Invalid command. Please use the format: /picalendar [search term], /pidocs, or /picurrent" }))
          .setMimeType(ContentService.MimeType.JSON);
      }
    } catch (error) {
      // Log the error for debugging
      Logger.log(error.toString());
      return ContentService.createTextOutput(JSON.stringify({ text: "An error occurred. Please try again." }))
        .setMimeType(ContentService.MimeType.JSON);
    }
  }

  /**
   * Parse URL-encoded form data from Slack POST request.
   * @param {string} data - The URL-encoded form data.
   * @returns {object} - An object containing parsed key-value pairs.
   */
  parseFormData(data) {
    const result = {};
    const pairs = data.split('&');
    for (let i = 0; i < pairs.length; i++) {
      const pair = pairs[i].split('=');
      result[decodeURIComponent(pair[0])] = decodeURIComponent(pair[1] || '');
    }
    return result;
  }

  /**
   * Search events by title in the Google Calendar.
   * @param {string} searchTerm - The term to search for in event titles.
   * @returns {Array} - A list of events matching the search term.
   */
  searchEventsByTitle(searchTerm) {
    const startDate = new Date();
    startDate.setFullYear(startDate.getFullYear() - 1); // Search for events from one year ago
    const endDate = new Date(startDate);
    endDate.setFullYear(startDate.getFullYear() + 2); // Search for events up to one year in the future

    const events = this.calendar.getEvents(startDate, endDate);
    Logger.log(`Total events: ${events.length}`); // Log the total number of events for debugging
    Logger.log(`Event titles: ${events.map(event => event.getTitle()).join(', ')}`); // Log all event titles for debugging
    const regex = new RegExp(searchTerm, 'i'); // Create a case-insensitive regex for the search term
    return events.filter(event => regex.test(event.getTitle()) && !this.isIgnoredEvent(event));
  }
}

// Example usage:
// Set up triggers in the Apps Script UI:
// 1. Weekly trigger for bot.sendWeeklyAllDayEvents() on Monday
// 2. Daily trigger for bot.sendDailyTimeSpecificEvents() every day

const bot = new SlackCalendarBot(CALENDAR_ID, SLACK_WEBHOOK_URL);

/**
 * Handle POST requests from Slack.
 * @param {object} e - The event object containing POST data.
 * @returns {ContentService.TextOutput} - The response to send back to Slack.
 */
function doPost(e) {
  return bot.doPost(e);
}

/**
 * Test the doPost function using a mock payload.
 */
function testDoPost() {
  bot.testDoPost();
}
