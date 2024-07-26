const SLACK_WEBHOOK_URL = "SLACK_WEBHOOK"; // Replace with your Slack webhook URL
const CALENDAR_ID = "CALENDAR URL";

class SlackCalendarBot {
  constructor(calendarId, slackWebhookUrl) {
    this.calendar = CalendarApp.getCalendarById(calendarId);
    this.slackWebhookUrl = slackWebhookUrl;
  }

  sendDailyTimeSpecificEvents() {
    const today = new Date();
    today.setHours(0, 0, 0, 0); // Start of the day
    const endOfDay = new Date(today);
    endOfDay.setHours(23, 59, 59, 999); // End of the day

    const events = this.calendar.getEvents(today, endOfDay);
    const timeSpecificEvents = events.filter(event => 
      !event.isAllDayEvent() &&
      (event.getStartTime().getTime() >= today.getTime() && event.getEndTime().getTime() <= endOfDay.getTime())
    );

    const message = this.buildEventMessage(timeSpecificEvents, "Today's PI events:", this.formatTimeSpecificEvent);

    if (message) {
      this.sendToSlack(message);
    }
  }

  sendWeeklyAllDayEvents() {
    const today = new Date();
    const monday = this.getMondayOfCurrentWeek(today);
    const sunday = this.getSundayOfCurrentWeek(monday);

    const events = this.calendar.getEvents(monday, sunday);
    const allDayEvents = events.filter(event => event.isAllDayEvent() || this.isMultiDayEvent(event, monday, sunday));

    const message = this.buildEventMessage(allDayEvents, `This week's PI events:`, this.formatWeeklyEvent);

    if (message) {
      this.sendToSlack(message);
    }
  }

  getDocsForRecentAndUpcomingEvents() {
    const now = new Date();
    const events = this.calendar.getEvents(new Date(now.getFullYear() - 1, now.getMonth(), now.getDate()), new Date(now.getFullYear() + 1, now.getMonth(), now.getDate()));
    
    const timeSpecificEvents = events.filter(event => {
      const startTime = event.getStartTime();
      const endTime = event.getEndTime();
      const duration = (endTime - startTime) / (1000 * 60 * 60); // Duration in hours
      return !event.isAllDayEvent() && duration <= 4;
    });
    
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

  getCurrentMultiDayEvents() {
    const today = new Date();
    today.setHours(0, 0, 0, 0); // Start of the day
    const endOfDay = new Date(today);
    endOfDay.setHours(23, 59, 59, 999); // End of the day

    const events = this.calendar.getEvents(today, endOfDay);
    const multiDayEvents = events.filter(event => this.isMultiDayEvent(event, today, endOfDay));

    let message = "Current PI & sprint:\n";
    if (multiDayEvents.length > 0) {
      message += multiDayEvents.map(event => `• ${event.getTitle()}`).join("\n");
    } else {
      message += "No multi-day events found for today.";
    }

    return message;
  }

  isMultiDayEvent(event, start, end) {
    const startTime = event.getStartTime();
    const endTime = event.getEndTime();
    return (startTime < end && endTime > start) || (endTime.getDate() !== startTime.getDate());
  }

  buildEventMessage(events, header, formatEvent) {
    if (events.length === 0) {
      return null;
    }

    const message = events.map(formatEvent).join("\n\n");

    return `${header}\n${message}`;
  }

  formatTimeSpecificEvent(event) {
    const start = event.getStartTime();
    const end = event.getEndTime();
    return `  • ${event.getTitle()} (${start.toTimeString().split(' ')[0]} - ${end.toTimeString().split(' ')[0]})`;
  }

  formatWeeklyEvent(event) {
    const start = event.getStartTime();
    const end = event.getEndTime();
    let formattedEvent;

    if (event.isAllDayEvent() || start.toDateString() !== end.toDateString()) {
      formattedEvent = `• ${event.getTitle()} (${start.toLocaleDateString('en-US', { month: 'short', day: 'numeric' })} - ${end.toLocaleDateString('en-US', { month: 'short', day: 'numeric' })})`;
    } else {
      formattedEvent = `• ${event.getTitle()} (${start.toLocaleDateString('en-US', { month: 'short', day: 'numeric' })} ${start.toLocaleTimeString('en-US', { hour: '2-digit', minute: '2-digit' })})`;
    }

    return formattedEvent;
  }

  cleanHtml(html) {
    if (!html) return '';

    // Remove HTML tags and extra whitespace
    const plainText = html.replace(/<[^>]*>/g, '').replace(/\s+/g, ' ').trim();
    return plainText;
  }

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

  getMondayOfCurrentWeek(today) {
    const monday = new Date(today);
    monday.setDate(today.getDate() - today.getDay() + 1); // Get the Monday of the current week
    monday.setHours(0, 0, 0, 0);
    return monday;
  }

  getSundayOfCurrentWeek(monday) {
    const sunday = new Date(monday);
    sunday.setDate(monday.getDate() + 6); // Get the Sunday of the current week
    sunday.setHours(23, 59, 59, 999);
    return sunday;
  }

  doPost(e) {
    try {
      const slackData = this.parseFormData(e.postData.contents);
      Logger.log(`slackData: ${JSON.stringify(slackData)}`); // Log the slackData for debugging
      const commandText = slackData.command.trim();
      const searchTerm = slackData.text.trim();
      Logger.log(`Command text: ${commandText}`); // Log the command text for debugging
      Logger.log(`Search term: ${searchTerm}`); // Log the search term for debugging

      if (commandText === '/picalendar') {
        if (searchTerm) {
          const events = this.searchEventsByTitle(searchTerm);
          Logger.log(`Filtered events: ${JSON.stringify(events.map(event => event.getTitle()))}`); // Log the event titles for debugging
          const message = this.buildEventMessage(events, `Events matching "${searchTerm}":`, this.formatWeeklyEvent) || `No event found for ${searchTerm}`;

          return ContentService.createTextOutput(JSON.stringify({ text: message }))
            .setMimeType(ContentService.MimeType.JSON);
        }
        return ContentService.createTextOutput(JSON.stringify({ text: "Invalid command. Please use the format: /picalendar [search term]" }))
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

  parseFormData(data) {
    const result = {};
    const pairs = data.split('&');
    for (let i = 0; i < pairs.length; i++) {
      const pair = pairs[i].split('=');
      result[decodeURIComponent(pair[0])] = decodeURIComponent(pair[1] || '');
    }
    return result;
  }

  searchEventsByTitle(searchTerm) {
    const startDate = new Date();
    startDate.setFullYear(startDate.getFullYear() - 1); // Search for events from one year ago
    const endDate = new Date(startDate);
    endDate.setFullYear(startDate.getFullYear() + 2); // Search for events up to one year in the future

    const events = this.calendar.getEvents(startDate, endDate);
    Logger.log(`Total events: ${events.length}`); // Log the total number of events for debugging
    Logger.log(`Event titles: ${events.map(event => event.getTitle()).join(', ')}`); // Log all event titles for debugging
    const regex = new RegExp(searchTerm, 'i'); // Create a case-insensitive regex for the search term
    return events.filter(event => regex.test(event.getTitle()));
  }

  testDoPost() {
    const testPayload = {
      token: "testToken",
      team_id: "T0001",
      team_domain: "example",
      channel_id: "C2147483705",
      channel_name: "test",
      user_id: "U2147483697",
      user_name: "Steve",
      command: "/picurrent",
      text: "",
      response_url: "https://hooks.slack.com/commands/1234/5678",
      trigger_id: "13345224609.738474920.8088930838d88f008e0"
    };
    
    const e = {
      postData: {
        contents: Object.entries(testPayload).map(([key, value]) => `${encodeURIComponent(key)}=${encodeURIComponent(value)}`).join('&')
      }
    };

    try {
      const slackData = this.parseFormData(e.postData.contents);
      Logger.log(`slackData: ${JSON.stringify(slackData)}`); // Log the slackData for debugging
      const commandText = slackData.command.trim();
      const searchTerm = slackData.text.trim();
      Logger.log(`Command text: ${commandText}`); // Log the command text for debugging
      Logger.log(`Search term: ${searchTerm}`); // Log the search term for debugging

      if (commandText === '/picalendar') {
        if (searchTerm) {
          const events = this.searchEventsByTitle(searchTerm);
          Logger.log(`Filtered events: ${JSON.stringify(events.map(event => event.getTitle()))}`); // Log the event titles for debugging
          const message = this.buildEventMessage(events, `Events matching "${searchTerm}":`, this.formatWeeklyEvent) || `No event found for ${searchTerm}`;
          Logger.log(`Message: ${message}`); // Log the message for debugging
        } else {
          Logger.log("Invalid command. Please use the format: /picalendar [search term]");
        }
      } else if (commandText === '/pidocs') {
        const message = this.getDocsForRecentAndUpcomingEvents();
        Logger.log(`Message: ${message}`); // Log the message for debugging
      } else if (commandText === '/picurrent') {
        const message = this.getCurrentMultiDayEvents();
        Logger.log(`Message: ${message}`); // Log the message for debugging
      } else {
        Logger.log("Invalid command. Please use the format: /picalendar [search term], /pidocs, or /picurrent");
      }
    } catch (error) {
      // Log the error for debugging
      Logger.log(error.toString());
    }
  }
}

// Example usage:
// Set up triggers in the Apps Script UI:
// 1. Weekly trigger for bot.sendWeeklyAllDayEvents() on Monday
// 2. Daily trigger for bot.sendDailyTimeSpecificEvents() every day

const bot = new SlackCalendarBot(CALENDAR_ID, SLACK_WEBHOOK_URL);

function doPost(e) {
  return bot.doPost(e);
}

function testDoPost() {
  bot.testDoPost();
}
