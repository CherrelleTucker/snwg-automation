/**
 * Conference Events Slack Bot - Complete Implementation
 * 
 * A Google Apps Script application that integrates with Slack to provide conference event information.
 * The bot responds to slash commands to display event information from a Google Calendar.
 * 
 * @author Your Name
 * @version 1.0
 * @lastModified 2024-03-19
 * 
 * Slash Commands:
 * - /conftoday: Shows all events scheduled for today
 * - /confnow: Shows currently ongoing events
 * - /confnext: Shows events starting within the next hour
 * - /conftomorrow: Shows all events scheduled for tomorrow
 * 
 * Required Properties:
 * - SLACK_BOT_TOKEN: Slack bot user OAuth token (set in Script Properties)
 */

/** Command handler functions - Must be defined before doPost */

/**
 * Handles /conftoday command to fetch today's events
 * @param {Object} requestData - Slack request data containing response_url
 */
function handleconfToday(requestData) {
    try {
        Logger.log('Handling /conftoday command');
        const events = fetchEventsFromCalendar();
        const todayEvents = filterEventsForToday(events);
        const formattedMessage = formatEventsMessage(todayEvents);
        sendSlackResponse(requestData.response_url, formattedMessage);
    } catch (error) {
        Logger.log('Error in handleconfToday: ' + error.message);
        sendSlackResponse(requestData.response_url, 'Error processing today\'s events: ' + error.message);
    }
}

/**
 * Handles /confnow command to fetch current events
 * @param {Object} requestData - Slack request data containing response_url
 */
function handleconfNow(requestData) {
    try {
        Logger.log('Handling /confnow command');
        const events = fetchEventsFromCalendar();
        const nowEvents = filterEventsForNow(events);
        const formattedMessage = formatEventsMessage(nowEvents);
        sendSlackResponse(requestData.response_url, formattedMessage);
    } catch (error) {
        Logger.log('Error in handleconfNow: ' + error.message);
        sendSlackResponse(requestData.response_url, 'Error processing current events: ' + error.message);
    }
}

/**
 * Handles /confnext command to fetch upcoming events
 * @param {Object} requestData - Slack request data containing response_url
 */
function handleconfNext(requestData) {
    try {
        Logger.log('Handling /confnext command');
        const events = fetchEventsFromCalendar();
        const nextEvents = filterEventsForNextHour(events);
        const formattedMessage = formatEventsMessage(nextEvents);
        sendSlackResponse(requestData.response_url, formattedMessage);
    } catch (error) {
        Logger.log('Error in handleconfNext: ' + error.message);
        sendSlackResponse(requestData.response_url, 'Error processing next events: ' + error.message);
    }
}

/**
 * Handles /conftomorrow command to fetch tomorrow's events
 * @param {Object} requestData - Slack request data containing response_url
 */
function handleconftomorrow(requestData) {
    try {
        Logger.log('Handling /conftomorrow command');
        const events = fetchEventsFromCalendar();
        const tomorrowEvents = filterEventsForTomorrow(events);
        const formattedMessage = formatEventsMessage(tomorrowEvents);
        sendSlackResponse(requestData.response_url, formattedMessage);
    } catch (error) {
        Logger.log('Error in handleconftomorrow: ' + error.message);
        sendSlackResponse(requestData.response_url, 'Error processing tomorrow\'s events: ' + error.message);
    }
}

/**
 * Main entry point for handling Slack slash commands
 * @param {Object} e - Event object containing POST data from Slack
 * @return {TextOutput} Empty response or error message in JSON format
 */
function doPost(e) {
    // Declare requestData in outer scope for error handling
    let requestData;
    try {
        Logger.log('doPost triggered');

        // Validate Slack bot token
        const SLACK_BOT_TOKEN = PropertiesService.getScriptProperties().getProperty('SLACK_BOT_TOKEN');
        if (!SLACK_BOT_TOKEN) {
            throw new Error('Slack bot token not found in script properties');
        }

        // Validate incoming request
        if (!e.postData || !e.postData.contents) {
            throw new Error('No post data received');
        }

        // Parse and validate request data
        requestData = parseUrlEncoded(e.postData.contents);
        Logger.log('Request data parsed successfully: ' + JSON.stringify(requestData));

        if (!requestData.command || !requestData.response_url) {
            throw new Error('Invalid request data structure');
        }

        // Process command
        const command = requestData.command.trim().toLowerCase();
        Logger.log('Command received: ' + command);

        // Send immediate acknowledgment
        sendSlackResponse(requestData.response_url, 'Processing your request...');

        // Route command to appropriate handler
        switch (command) {
            case '/conftoday':
                handleconfToday(requestData);
                break;
            case '/confnow':
                handleconfNow(requestData);
                break;
            case '/confnext':
                handleconfNext(requestData);
                break;
            case '/conftomorrow':
                handleconftomorrow(requestData);
                break;
            default:
                sendSlackResponse(requestData.response_url, 
                    'Unknown command. Please use /conftoday, /confnow, /confnext, or conftomorrow.');
        }

        return ContentService.createTextOutput('');

    } catch (error) {
        Logger.log('Error in doPost: ' + error.message);

        // Log request data if available
        if (e && e.postData && e.postData.contents) {
            Logger.log('Request Data: ' + e.postData.contents);
        }

        // Send error message back to Slack if possible
        if (requestData && requestData.response_url) {
            sendSlackResponse(requestData.response_url, 'Error processing request: ' + error.message);
        }

        return ContentService.createTextOutput(JSON.stringify({
            error: error.message
        })).setMimeType(ContentService.MimeType.JSON);
    }
}

/** Helper Functions */

/**
 * Parses URL-encoded form data into a key-value object
 * @param {string} data - URL-encoded form data string
 * @return {Object} Parsed key-value pairs
 */
function parseUrlEncoded(data) {
    const params = data.split('&');
    const result = {};

    params.forEach(param => {
        const [key, value] = param.split('=');
        // Replace '+' with space and decode URI components
        result[decodeURIComponent(key)] = decodeURIComponent(value.replace(/\+/g, ' '));
    });

    return result;
}

/**
 * Fetches events from Google Calendar within a 48-hour window
 * @return {Array<Object>} Array of event objects with normalized properties
 */
function fetchEventsFromCalendar() {
    try {
        Logger.log('Fetching events from Google Calendar');
        
        // Calendar ID for conference events
        const calendarId = 'c_389708f8a51569fb24e7df2705bbd14898ae728db0cccc7c288aec23f60f73be@group.calendar.google.com';
        
        // Set time window for events (now to 48 hours ahead)
        const startTime = new Date();
        const endTime = new Date();
        endTime.setDate(endTime.getDate() + 2);

        // Fetch events from calendar
        const events = CalendarApp.getCalendarById(calendarId).getEvents(startTime, endTime);

        // Normalize event data for consistent handling
        return events.map(event => ({
            title: event.getTitle(),
            startTime: event.getStartTime().toISOString(),
            endTime: event.getEndTime().toISOString(),
            location: event.getLocation(),
            description: event.getDescription() || 'No description available'
        }));

    } catch (error) {
        Logger.log('Error in fetchEventsFromCalendar: ' + error.message);
        return [];
    }
}

/**
 * Filters events occurring on the current day
 * @param {Array<Object>} events - Array of event objects
 * @return {Array<Object>} Filtered events for today
 */
function filterEventsForToday(events) {
    try {
        // Set today's date to midnight for date-only comparison
        const today = new Date();
        today.setHours(0, 0, 0, 0);

        return events.filter(event => {
            const eventDate = new Date(event.startTime);
            eventDate.setHours(0, 0, 0, 0);
            return eventDate.getTime() === today.getTime();
        });

    } catch (error) {
        Logger.log('Error in filterEventsForToday: ' + error.message);
        return [];
    }
}

/**
 * Filters currently ongoing events
 * @param {Array<Object>} events - Array of event objects
 * @return {Array<Object>} Currently ongoing events
 */
function filterEventsForNow(events) {
    try {
        const currentTime = new Date();

        return events.filter(event => {
            const eventStartTime = new Date(event.startTime);
            const eventEndTime = new Date(event.endTime);
            return eventStartTime <= currentTime && eventEndTime >= currentTime;
        });

    } catch (error) {
        Logger.log('Error in filterEventsForNow: ' + error.message);
        return [];
    }
}

/**
 * Filters events starting within the next hour
 * @param {Array<Object>} events - Array of event objects
 * @return {Array<Object>} Events starting in the next hour
 */
function filterEventsForNextHour(events) {
    try {
        const currentTime = new Date();
        const nextHourTime = new Date(currentTime.getTime() + 3600000); // Add 1 hour in milliseconds

        return events.filter(event => {
            const eventStartTime = new Date(event.startTime);
            return eventStartTime > currentTime && eventStartTime <= nextHourTime;
        });

    } catch (error) {
        Logger.log('Error in filterEventsForNextHour: ' + error.message);
        return [];
    }
}

/**
 * Filters events occurring tomorrow
 * @param {Array<Object>} events - Array of event objects
 * @return {Array<Object>} Events scheduled for tomorrow
 */
function filterEventsForTomorrow(events) {
    try {
        // Set tomorrow's date range
        const tomorrow = new Date();
        tomorrow.setDate(tomorrow.getDate() + 1);
        tomorrow.setHours(0, 0, 0, 0);

        const endOfTomorrow = new Date(tomorrow);
        endOfTomorrow.setHours(23, 59, 59, 999);

        return events.filter(event => {
            const eventStartTime = new Date(event.startTime);
            return eventStartTime >= tomorrow && eventStartTime <= endOfTomorrow;
        });

    } catch (error) {
        Logger.log('Error in filterEventsForTomorrow: ' + error.message);
        return [];
    }
}

/**
 * Extracts presenter name from the clean portion of description
 * @param {string} description - Event description text
 * @return {string} Cleaned presenter name
 */
function extractHostType(description) {
    try {
        // Look for "Presenter: " followed by text up to the next section
        const presenterMatch = description.match(/Presenter:\s*([^Location\n]+)/i);
        if (presenterMatch && presenterMatch[1]) {
            // Clean up the presenter name
            return presenterMatch[1]
                .replace(/Session Link.*$/, '') // Remove session link and everything after
                .replace(/<[^>]+>/g, '') // Remove any HTML tags
                .trim();
        }
        return 'Unknown Presenter';
    } catch (error) {
        Logger.log('Error in extractHostType: ' + error.message);
        return 'Unknown Presenter';
    }
}


/**
 * Extracts presentation type from event description
 * @param {string} description - Event description text
 * @return {string} Extracted presentation type
 */
function extractPresentationType(description) {
    try {
        // First try the explicit Presentation Type field
        let typeMatch = description.match(/Presentation Type:\s*(.*?)(?=\n|$)/i);
        
        // If not found, try to find it in parentheses after the title
        if (!typeMatch || !typeMatch[1]) {
            typeMatch = description.match(/\((\w+)\).*?(?=;|$)/);
        }
        
        if (typeMatch && typeMatch[1]) {
            return typeMatch[1].trim();
        }
        
        return 'Unknown Type';
    } catch (error) {
        Logger.log('Error in extractPresentationType: ' + error.message);
        return 'Unknown Type';
    }
}

/**
 * Extracts presenter name from event description
 * @param {Object} event - Calendar event object
 * @return {string} Presenter name
 */
function extractPresenter(event) {
  try {
    // Check for explicit Presenter field first
    const presenterMatch = event.description.match(/Presenter:\s*([^\n]+)/);
    if (presenterMatch && presenterMatch[1].trim() !== 'Unknown Presenter') {
      return presenterMatch[1].trim();
    }
    
    // Fall back to Convener/CoConvener if no presenter
    const convenerMatch = event.description.match(/Convener\/CoConvener:\s*([^\n]+)/);
    if (convenerMatch) {
      return convenerMatch[1].trim();
    }

    // If neither is found, check for standalone Presenter line
    const standalonePresentMatch = event.description.match(/^([^:\n]+)$/m);
    if (standalonePresentMatch) {
      return standalonePresentMatch[1].trim();
    }

    return 'Unknown Presenter';
  } catch (error) {
    Logger.log('Error in extractPresenter: ' + error.message);
    return 'Unknown Presenter';
  }
}

/**
 * Extracts session link from event description
 * @param {Object} event - Calendar event object
 * @return {string|null} Session link URL
 */
function extractSessionLink(event) {
  try {
    // Try multiple session link formats
    const patterns = [
      /Session Link:[\s\n]*(https?:\/\/[^\s\n]+)/i,
      /Session Link[\s\n]+(https?:\/\/[^\s\n]+)/i,
      /(https?:\/\/agu\.confex\.com\/agu\/agu24\/meetingapp\.cgi\/[^\s\n]+)/i
    ];

    for (const pattern of patterns) {
      const match = event.description.match(pattern);
      if (match && match[1]) {
        return match[1].trim();
      }
    }
    return null;
  } catch (error) {
    Logger.log('Error in extractSessionLink: ' + error.message);
    return null;
  }
}

/**
 * Formats events into a Slack message with preserved hyperlinks and footer links
 * @param {Array<Object>} events - Array of event objects to format
 * @return {Object} Formatted message object for Slack using Block Kit
 */
function formatEventsMessage(events) {
  try {
    const blocks = [];
    
    if (events.length === 0) {
      blocks.push({
        type: "section",
        text: {
          type: "mrkdwn",
          text: "No events found for the specified time period."
        }
      });
    } else {
      events.forEach(event => {
        const eventStartTime = formatTime(event.startTime);
        const eventEndTime = formatTime(event.endTime);
        const presentationType = extractPresentationType(event.description);
        const location = event.location.split('(')[0].trim();
        const presenter = extractPresenter(event);
        const sessionLink = extractSessionLink(event);

        let messageText = [
          `• ${eventStartTime} - ${eventEndTime}`,
          `*${event.title}* (${presentationType})`,
          `Presenter: ${presenter}`,
          `Location: ${location}`
        ];

        // Add session link if available
        if (sessionLink) {
          messageText.push(`<${sessionLink}|Session Link>`);
        }

        blocks.push({
          type: "section",
          text: {
            type: "mrkdwn",
            text: messageText.join('\n')
          }
        });
      });
    }

    // Add divider before footer
    blocks.push({
      type: "divider"
    });

    // Add footer links
    blocks.push({
      type: "section",
      text: {
        type: "mrkdwn",
        text: "• <https://docs.google.com/spreadsheets/d/1wVwjI213qPl3XpJCApzNyzGpRIvP93JTrP7XITM1xo4/edit?gid=952720194#gid=952720194|AGU Fall 2024 STI Sheet>\n• <https://calendar.google.com/calendar/u/0?cid=Y18zODk3MDhmOGE1MTU2OWZiMjRlN2RmMjcwNWJiZDE0ODk4YWU3MjhkYjBjY2NjN2MyODhhZWMyM2Y2MGY3M2JlQGdyb3VwLmNhbGVuZGFyLmdvb2dsZS5jb20|IMPACT Conference Attendance Calendar>"
      }
    });

    return { blocks };
  } catch (error) {
    Logger.log('Error in formatEventsMessage: ' + error.message);
    return {
      blocks: [{
        type: "section",
        text: {
          type: "mrkdwn",
          text: "Error formatting events message"
        }
      }]
    };
  }
}

/**
 * Formats ISO time string into readable format for Slack
 * @param {string} isoString - ISO format time string
 * @return {string} Formatted time string
 */
function formatTime(isoString) {
    try {
        const date = new Date(isoString);
        const options = { hour: '2-digit', minute: '2-digit', timeZoneName: 'short' };
        return date.toLocaleTimeString('en-US', options);
    } catch (error) {
        Logger.log('Error in formatTime: ' + error.message);
        return 'Unknown Time';
    }
}

/**
 * Sends formatted response back to Slack using Block Kit
 * @param {string} responseUrl - Slack response URL
 * @param {string|Object} message - Message to send
 */
function sendSlackResponse(responseUrl, message) {
  try {
    let payload;
    if (typeof message === 'string') {
      payload = {
        blocks: [{
          type: "section",
          text: {
            type: "mrkdwn",
            text: message
          }
        }]
      };
    } else {
      payload = message;
    }

    const options = {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    };

    UrlFetchApp.fetch(responseUrl, options);
    Logger.log('Response sent to Slack successfully');
  } catch (error) {
    Logger.log('Error in sendSlackResponse: ' + error.message);
  }
}


/**
 * Test function for debugging /conftoday command
 * Simulates command execution and logs results
 */
function testconftoday() {
    try {
        Logger.log('Testing /conftoday command');
        const events = fetchEventsFromCalendar();
        const todayEvents = filterEventsForToday(events);
        const formattedMessage = formatEventsMessage(todayEvents);
        Logger.log('Formatted message for /conftoday: ' + formattedMessage);
    } catch (error) {
        Logger.log('Error in testconftoday: ' + error.message);
    }
}
