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
 * Filters events by acronym prefix
 * @param {Array<Object>} events - Array of event objects
 * @param {string} acronym - Acronym to filter by
 * @return {Array<Object>} Filtered events starting with the acronym
 */
function filterEventsByAcronym(events, acronym) {
  if (!acronym) return events;
  
  return events.filter(event => {
    const title = event.title || '';
    // Match acronym followed by colon at start of title
    return title.toUpperCase().startsWith(acronym.toUpperCase() + ':');
  });
}

/**
 * Modified command handlers to support acronym filtering
 */
function handleconfToday(requestData) {
  try {
    Logger.log('Handling /conftoday command');
    const acronym = requestData.text ? requestData.text.trim() : '';
    
    const events = fetchEventsFromCalendar();
    let todayEvents = filterEventsForToday(events);
    todayEvents = filterEventsByAcronym(todayEvents, acronym);
    
    if (todayEvents.length === 0 && acronym) {
      sendSlackResponse(requestData.response_url, 
        `No events found for ${acronym} for today`);
      return;
    }
    
    const formattedMessage = formatEventsMessage(todayEvents);
    sendSlackResponse(requestData.response_url, formattedMessage);
  } catch (error) {
    Logger.log('Error in handleconfToday: ' + error.message);
    sendSlackResponse(requestData.response_url, 
      'Error processing today\'s events: ' + error.message);
  }
}

function handleconfNow(requestData) {
  try {
    Logger.log('Handling /confnow command');
    const acronym = requestData.text ? requestData.text.trim() : '';
    
    const events = fetchEventsFromCalendar();
    let nowEvents = filterEventsForNow(events);
    nowEvents = filterEventsByAcronym(nowEvents, acronym);
    
    if (nowEvents.length === 0 && acronym) {
      sendSlackResponse(requestData.response_url, 
        `No events found for ${acronym} for this hour`);
      return;
    }
    
    const formattedMessage = formatEventsMessage(nowEvents);
    sendSlackResponse(requestData.response_url, formattedMessage);
  } catch (error) {
    Logger.log('Error in handleconfNow: ' + error.message);
    sendSlackResponse(requestData.response_url, 
      'Error processing current events: ' + error.message);
  }
}

function handleconfNext(requestData) {
  try {
    Logger.log('Handling /confnext command');
    const acronym = requestData.text ? requestData.text.trim() : '';
    
    const events = fetchEventsFromCalendar();
    let nextEvents = filterEventsForNextHour(events);
    nextEvents = filterEventsByAcronym(nextEvents, acronym);
    
    if (nextEvents.length === 0 && acronym) {
      sendSlackResponse(requestData.response_url, 
        `No events found for ${acronym} for the next hour`);
      return;
    }
    
    const formattedMessage = formatEventsMessage(nextEvents);
    sendSlackResponse(requestData.response_url, formattedMessage);
  } catch (error) {
    Logger.log('Error in handleconfNext: ' + error.message);
    sendSlackResponse(requestData.response_url, 
      'Error processing next events: ' + error.message);
  }
}

function handleconftomorrow(requestData) {
  try {
    Logger.log('Handling /conftomorrow command');
    const acronym = requestData.text ? requestData.text.trim() : '';
    
    const events = fetchEventsFromCalendar();
    let tomorrowEvents = filterEventsForTomorrow(events);
    tomorrowEvents = filterEventsByAcronym(tomorrowEvents, acronym);
    
    if (tomorrowEvents.length === 0 && acronym) {
      sendSlackResponse(requestData.response_url, 
        `No events found for ${acronym} for tomorrow`);
      return;
    }
    
    const formattedMessage = formatEventsMessage(tomorrowEvents);
    sendSlackResponse(requestData.response_url, formattedMessage);
  } catch (error) {
    Logger.log('Error in handleconftomorrow: ' + error.message);
    sendSlackResponse(requestData.response_url, 
      'Error processing tomorrow\'s events: ' + error.message);
  }
}

/**
 * Filters events for the current week (Sunday-Saturday)
 * @param {Array<Object>} events - Array of event objects
 * @return {Array<Object>} Events for current week
 */
function filterEventsForThisWeek(events) {
  try {
    const now = new Date();
    const currentDay = now.getDay();
    const sunday = new Date(now);
    sunday.setDate(now.getDate() - currentDay);
    sunday.setHours(0, 0, 0, 0);
    
    const saturday = new Date(sunday);
    saturday.setDate(sunday.getDate() + 6);
    saturday.setHours(23, 59, 59, 999);
    
    return events.filter(event => {
      const eventDate = new Date(event.startTime);
      return eventDate >= sunday && eventDate <= saturday;
    });
  } catch (error) {
    Logger.log('Error in filterEventsForThisWeek: ' + error.message);
    return [];
  }
}

/**
 * Filters events for next week (Sunday-Saturday)
 * @param {Array<Object>} events - Array of event objects
 * @return {Array<Object>} Events for next week
 */
function filterEventsForNextWeek(events) {
  try {
    const now = new Date();
    const currentDay = now.getDay();
    const nextSunday = new Date(now);
    nextSunday.setDate(now.getDate() + (7 - currentDay));
    nextSunday.setHours(0, 0, 0, 0);
    
    const nextSaturday = new Date(nextSunday);
    nextSaturday.setDate(nextSunday.getDate() + 6);
    nextSaturday.setHours(23, 59, 59, 999);
    
    return events.filter(event => {
      const eventDate = new Date(event.startTime);
      return eventDate >= nextSunday && eventDate <= nextSaturday;
    });
  } catch (error) {
    Logger.log('Error in filterEventsForNextWeek: ' + error.message);
    return [];
  }
}

/**
 * Handles /confthisweek command
 * @param {Object} requestData - Slack request data
 */
function handleconfThisWeek(requestData) {
  try {
    Logger.log('Handling /confthisweek command');
    const acronym = requestData.text ? requestData.text.trim() : '';
    
    const events = fetchEventsFromCalendar();
    let thisWeekEvents = filterEventsForThisWeek(events);
    thisWeekEvents = filterEventsByAcronym(thisWeekEvents, acronym);
    
    if (thisWeekEvents.length === 0 && acronym) {
      sendSlackResponse(requestData.response_url, 
        `No events found for ${acronym} this week`);
      return;
    }
    
    const formattedMessage = formatEventsMessage(thisWeekEvents);
    sendSlackResponse(requestData.response_url, formattedMessage);
  } catch (error) {
    Logger.log('Error in handleconfThisWeek: ' + error.message);
    sendSlackResponse(requestData.response_url, 
      'Error processing this week\'s events: ' + error.message);
  }
}

/**
 * Handles /confnextweek command
 * @param {Object} requestData - Slack request data
 */
function handleconfNextWeek(requestData) {
  try {
    Logger.log('Handling /confnextweek command');
    const acronym = requestData.text ? requestData.text.trim() : '';
    
    const events = fetchEventsFromCalendar();
    let nextWeekEvents = filterEventsForNextWeek(events);
    nextWeekEvents = filterEventsByAcronym(nextWeekEvents, acronym);
    
    if (nextWeekEvents.length === 0 && acronym) {
      sendSlackResponse(requestData.response_url, 
        `No events found for ${acronym} next week`);
      return;
    }
    
    const formattedMessage = formatEventsMessage(nextWeekEvents);
    sendSlackResponse(requestData.response_url, formattedMessage);
  } catch (error) {
    Logger.log('Error in handleconfNextWeek: ' + error.message);
    sendSlackResponse(requestData.response_url, 
      'Error processing next week\'s events: ' + error.message);
  }
}

/**
 * Main entry point for handling Slack slash commands and events
 * @param {Object} e - Event object containing POST data from Slack
 * @return {TextOutput} Empty response or error message in JSON format
 */
function doPost(e) {
  let requestData;
  try {
    Logger.log('doPost triggered');
    
    const SLACK_BOT_TOKEN = PropertiesService.getScriptProperties().getProperty('SLACK_BOT_TOKEN');
    if (!SLACK_BOT_TOKEN) throw new Error('Slack bot token not found');
    
    if (!e.postData?.contents) throw new Error('No post data received');

    // Try parsing as JSON first
    try {
      const jsonData = JSON.parse(e.postData.contents);
      
      // Handle Events API
      if (jsonData.type === 'url_verification') {
        return ContentService.createTextOutput(jsonData.challenge);
      }
      
      if (jsonData.type === 'event_callback') {
        if (jsonData.event.type === 'app_home_opened') {
          handleAppHomeOpened(jsonData.event);
        }
        return ContentService.createTextOutput();
      }
    } catch {
      // If JSON parse fails, handle as URL-encoded
      requestData = parseUrlEncoded(e.postData.contents);
    }
    
    // Handle Slash Commands
    if (!requestData?.command || !requestData?.response_url) {
      throw new Error('Invalid request structure');
    }
    
    const command = requestData.command.trim().toLowerCase();
    Logger.log('Command received: ' + command);
    
    sendSlackResponse(requestData.response_url, 'Processing your request...');
    
    switch (command) {
      case '/confnow': handleconfNow(requestData); break;
      case '/confnext': handleconfNext(requestData); break;
      case '/conftoday': handleconfToday(requestData); break;
      case '/conftomorrow': handleconftomorrow(requestData); break;
      case '/confthisweek': handleconfThisWeek(requestData); break;
      case '/confnextweek': handleconfNextWeek(requestData); break;
      case '/confmaps': handleconfMaps(requestData); break;
      default:
        sendSlackResponse(requestData.response_url, 
          'Unknown command. Use /confnow, /confnext, /conftoday, /conftomorrow, /confthisweek, /confnextweek, or /confmaps.');
    }
    
    return ContentService.createTextOutput('');
    
  } catch (error) {
    Logger.log('Error in doPost: ' + error.message);
    Logger.log('Request Data: ' + e.postData?.contents);
    
    if (requestData?.response_url) {
      sendSlackResponse(requestData.response_url, 'Error processing request: ' + error.message);
    }
    
    return ContentService.createTextOutput(JSON.stringify({error: error.message}))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// Add this function to your script
function doGet(e) {
  return ContentService.createTextOutput('Slack App is running');
}
/** Helper Functions */

/**
 * Handles /confmaps command
 * @param {Object} requestData - Slack request data
 */
function handleconfMaps(requestData) {
  try {
    Logger.log('Handling /confmaps command');
    const blocks = [
      {
        type: "section",
        text: {
          type: "mrkdwn",
          text: "*Conference Venue Maps & Resources*"
        }
      },
      {
        type: "section",
        text: {
          type: "mrkdwn",
          text: "• <https://raw.githubusercontent.com/CherrelleTucker/snwg-automation/c151e868eed07dfeafe66eac2bf9172ab779a9e1/conferenceBot/Walter%20E.%20Washington%20Convention%20Center%20Floorplan_Level2.png|Convention Center Level 2>\n" +
                "• <https://raw.githubusercontent.com/CherrelleTucker/snwg-automation/c151e868eed07dfeafe66eac2bf9172ab779a9e1/conferenceBot/Walter%20E.%20Washington%20Convention%20Center%20Floorplan_LevelStreet.png|Convention Center Street Level>\n" +
                "• <https://raw.githubusercontent.com/CherrelleTucker/snwg-automation/95f24e2e903c6be4a0612def66c40aa92803a76e/conferenceBot/Walter%20E.%20Washington%20Convention%20Center%20Floorplan_HallsABC.png|Convention Center Halls A-B-C>\n" +
                "• <https://raw.githubusercontent.com/CherrelleTucker/snwg-automation/d126cf7abf0aed484e16a5cd0cfa8313d6c8c127/conferenceBot/Walter%20E.%20Washington%20Convention%20Center%20Floorplan_Ballroom.png|Convention Center Ballroom>\n" +
                "• <https://raw.githubusercontent.com/CherrelleTucker/snwg-automation/c151e868eed07dfeafe66eac2bf9172ab779a9e1/conferenceBot/MarriottMarquis_Mint.png|Marriott Marquis Meeting Rooms>\n" +
                "• <https://raw.githubusercontent.com/CherrelleTucker/snwg-automation/c151e868eed07dfeafe66eac2bf9172ab779a9e1/conferenceBot/MarriottMarquis_Marquis3_4_12_13.png|Marriott Marquis Salons>\n" +
                "• <https://raw.githubusercontent.com/CherrelleTucker/snwg-automation/1c527b0276930f5d200305f1e75ff18406d33dc5/conferenceBot/DC_metro.png|DC Metro Map>"
        }
      },
      {
        type: "section",
        text: {
          type: "mrkdwn",
          text: "*Virtual Tours*\n• <https://truetour.app/properties/1443/experiences/50217?shareId=228365|Marriott Virtual Tour>\n• <https://eventsdc.com/venue/walter-e-washington-convention-center/virtual-tour|Convention Center Virtual Tour>"
        }
      }
    ];
    
    sendSlackResponse(requestData.response_url, { blocks });
  } catch (error) {
    Logger.log('Error in handleconfMaps: ' + error.message);
    sendSlackResponse(requestData.response_url, 'Error displaying maps: ' + error.message);
  }
}

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
 * Fetches events from Google Calendar with extended window
 * @return {Array<Object>} Array of event objects
 */
function fetchEventsFromCalendar() {
  try {
    Logger.log('Fetching events from Google Calendar');
    const calendarId = 'c_389708f8a51569fb24e7df2705bbd14898ae728db0cccc7c288aec23f60f73be@group.calendar.google.com';
    
    // Set time window for events (now to 14 days ahead to cover next week fully)
    const startTime = new Date();
    const endTime = new Date();
    endTime.setDate(endTime.getDate() + 14);
    
    const events = CalendarApp.getCalendarById(calendarId).getEvents(startTime, endTime);
    
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
function cleanUrl(url) {
  if (!url) return null;
  // Remove Google Calendar's tracking parameters and any other unwanted additions
  return url
    .split('&')[0]  // Remove all parameters after first &
    .split('%22')[0]  // Remove any encoded quotes
    .replace(/\?.*$/, '');  // Remove everything after ?
}

function extractSessionLink(event) {
  try {
    const patterns = [
      /Session Link:[\s\n]*(https?:\/\/[^\s\n]+)/i,
      /Session Link[\s\n]+(https?:\/\/[^\s\n]+)/i,
      /(https?:\/\/agu\.confex\.com\/agu\/agu24\/meetingapp\.cgi\/[^\s\n]+)/i
    ];

    for (const pattern of patterns) {
      const match = event.description.match(pattern);
      if (match && match[1]) {
        return cleanUrl(match[1].trim());
      }
    }
    return null;
  } catch (error) {
    Logger.log('Error in extractSessionLink: ' + error.message);
    return null;
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
 * Maps location names to their corresponding hyperlinks
 * @param {string} location - Raw location string from event
 * @return {string} Formatted location with appropriate hyperlink
 */
function formatLocation(location) {
  // Clean the location string
  const cleanLocation = location.split('(')[0].trim();
  
  // Location mapping object
  const locationMap = {
    // Main venues
    'Convention Center': 'https://www.google.com/maps?q=801+Allen+Y.+Lew+Place+NW,+Washington,+DC+20001',
    'Marriott Marquis': 'https://www.google.com/maps?q=901+Massachusetts+Ave+NW,+Washington,+DC+20001',
    
    // Level 2 floor plan
    'Hall D': 'https://github.com/CherrelleTucker/snwg-automation/blob/c151e868eed07dfeafe66eac2bf9172ab779a9e1/conferenceBot/Walter%20E.%20Washington%20Convention%20Center%20Floorplan_Level2.png',
    
    // Street level rooms
    '150 B': 'https://raw.githubusercontent.com/CherrelleTucker/snwg-automation/c151e868eed07dfeafe66eac2bf9172ab779a9e1/conferenceBot/Walter%20E.%20Washington%20Convention%20Center%20Floorplan_LevelStreet.png',
    '144 C': 'https://raw.githubusercontent.com/CherrelleTucker/snwg-automation/c151e868eed07dfeafe66eac2bf9172ab779a9e1/conferenceBot/Walter%20E.%20Washington%20Convention%20Center%20Floorplan_LevelStreet.png',
    '144 A': 'https://raw.githubusercontent.com/CherrelleTucker/snwg-automation/c151e868eed07dfeafe66eac2bf9172ab779a9e1/conferenceBot/Walter%20E.%20Washington%20Convention%20Center%20Floorplan_LevelStreet.png',
    '144 A-C': 'https://raw.githubusercontent.com/CherrelleTucker/snwg-automation/c151e868eed07dfeafe66eac2bf9172ab779a9e1/conferenceBot/Walter%20E.%20Washington%20Convention%20Center%20Floorplan_LevelStreet.png',
    'Salon H': 'https://raw.githubusercontent.com/CherrelleTucker/snwg-automation/c151e868eed07dfeafe66eac2bf9172ab779a9e1/conferenceBot/Walter%20E.%20Washington%20Convention%20Center%20Floorplan_LevelStreet.png',
    'Salon C': 'https://raw.githubusercontent.com/CherrelleTucker/snwg-automation/c151e868eed07dfeafe66eac2bf9172ab779a9e1/conferenceBot/Walter%20E.%20Washington%20Convention%20Center%20Floorplan_LevelStreet.png',
    
    
    // Convention center halls
    'Hall A': 'https://raw.githubusercontent.com/CherrelleTucker/snwg-automation/95f24e2e903c6be4a0612def66c40aa92803a76e/conferenceBot/Walter%20E.%20Washington%20Convention%20Center%20Floorplan_HallsABC.png',
    'Hall B': 'https://raw.githubusercontent.com/CherrelleTucker/snwg-automation/95f24e2e903c6be4a0612def66c40aa92803a76e/conferenceBot/Walter%20E.%20Washington%20Convention%20Center%20Floorplan_HallsABC.png',
    'Hall C': 'https://raw.githubusercontent.com/CherrelleTucker/snwg-automation/95f24e2e903c6be4a0612def66c40aa92803a76e/conferenceBot/Walter%20E.%20Washington%20Convention%20Center%20Floorplan_HallsABC.png',
    'Hall B-C': 'https://raw.githubusercontent.com/CherrelleTucker/snwg-automation/95f24e2e903c6be4a0612def66c40aa92803a76e/conferenceBot/Walter%20E.%20Washington%20Convention%20Center%20Floorplan_HallsABC.png',

    // eLightning
    'eLightning Theater 1': 'https://raw.githubusercontent.com/CherrelleTucker/snwg-automation/95f24e2e903c6be4a0612def66c40aa92803a76e/conferenceBot/Walter%20E.%20Washington%20Convention%20Center%20Floorplan_HallsABC.png',
    'eLightning Theater 2': 'https://raw.githubusercontent.com/CherrelleTucker/snwg-automation/95f24e2e903c6be4a0612def66c40aa92803a76e/conferenceBot/Walter%20E.%20Washington%20Convention%20Center%20Floorplan_HallsABC.png',
    'eLightning Theater 3': 'https://raw.githubusercontent.com/CherrelleTucker/snwg-automation/95f24e2e903c6be4a0612def66c40aa92803a76e/conferenceBot/Walter%20E.%20Washington%20Convention%20Center%20Floorplan_HallsABC.png',
    'eLightning Theater 4': 'https://raw.githubusercontent.com/CherrelleTucker/snwg-automation/95f24e2e903c6be4a0612def66c40aa92803a76e/conferenceBot/Walter%20E.%20Washington%20Convention%20Center%20Floorplan_HallsABC.png',
    'eLightning Theater 5': 'https://raw.githubusercontent.com/CherrelleTucker/snwg-automation/95f24e2e903c6be4a0612def66c40aa92803a76e/conferenceBot/Walter%20E.%20Washington%20Convention%20Center%20Floorplan_HallsABC.png',
    'eLightning': 'https://raw.githubusercontent.com/CherrelleTucker/snwg-automation/95f24e2e903c6be4a0612def66c40aa92803a76e/conferenceBot/Walter%20E.%20Washington%20Convention%20Center%20Floorplan_HallsABC.png',
    
    // Marriott rooms
    'Mint': 'https://raw.githubusercontent.com/CherrelleTucker/snwg-automation/c151e868eed07dfeafe66eac2bf9172ab779a9e1/conferenceBot/MarriottMarquis_Mint.png',
    'Congress': 'https://raw.githubusercontent.com/CherrelleTucker/snwg-automation/c151e868eed07dfeafe66eac2bf9172ab779a9e1/conferenceBot/MarriottMarquis_Mint.png',
    'Capitol': 'https://raw.githubusercontent.com/CherrelleTucker/snwg-automation/c151e868eed07dfeafe66eac2bf9172ab779a9e1/conferenceBot/MarriottMarquis_Mint.png',
    'Capital/Congress': 'https://raw.githubusercontent.com/CherrelleTucker/snwg-automation/c151e868eed07dfeafe66eac2bf9172ab779a9e1/conferenceBot/MarriottMarquis_Mint.png',
    'Capitol/Congress': 'https://raw.githubusercontent.com/CherrelleTucker/snwg-automation/c151e868eed07dfeafe66eac2bf9172ab779a9e1/conferenceBot/MarriottMarquis_Mint.png',
    'Liberty I-K': 'https://raw.githubusercontent.com/CherrelleTucker/snwg-automation/c151e868eed07dfeafe66eac2bf9172ab779a9e1/conferenceBot/MarriottMarquis_Mint.png',
    
    // Marriott salons
    'Marquis 3-4': 'https://raw.githubusercontent.com/CherrelleTucker/snwg-automation/c151e868eed07dfeafe66eac2bf9172ab779a9e1/conferenceBot/MarriottMarquis_Marquis3_4_12_13.png',
    'Marquis 12-13': 'https://raw.githubusercontent.com/CherrelleTucker/snwg-automation/c151e868eed07dfeafe66eac2bf9172ab779a9e1/conferenceBot/MarriottMarquis_Marquis3_4_12_13.png',
    'Salon 3': 'https://raw.githubusercontent.com/CherrelleTucker/snwg-automation/c151e868eed07dfeafe66eac2bf9172ab779a9e1/conferenceBot/MarriottMarquis_Marquis3_4_12_13.png',
    'Salon 4': 'https://raw.githubusercontent.com/CherrelleTucker/snwg-automation/c151e868eed07dfeafe66eac2bf9172ab779a9e1/conferenceBot/MarriottMarquis_Marquis3_4_12_13.png',
    'Salon 12': 'https://raw.githubusercontent.com/CherrelleTucker/snwg-automation/c151e868eed07dfeafe66eac2bf9172ab779a9e1/conferenceBot/MarriottMarquis_Marquis3_4_12_13.png',
    'Salon 13': 'https://raw.githubusercontent.com/CherrelleTucker/snwg-automation/c151e868eed07dfeafe66eac2bf9172ab779a9e1/conferenceBot/MarriottMarquis_Marquis3_4_12_13.png',

    
    // Convention center ballroom
    'Ballroom A': 'https://raw.githubusercontent.com/CherrelleTucker/snwg-automation/d126cf7abf0aed484e16a5cd0cfa8313d6c8c127/conferenceBot/Walter%20E.%20Washington%20Convention%20Center%20Floorplan_Ballroom.png',
    'Ballroom B': 'https://raw.githubusercontent.com/CherrelleTucker/snwg-automation/d126cf7abf0aed484e16a5cd0cfa8313d6c8c127/conferenceBot/Walter%20E.%20Washington%20Convention%20Center%20Floorplan_Ballroom.png',
    'Ballroom C': 'https://raw.githubusercontent.com/CherrelleTucker/snwg-automation/d126cf7abf0aed484e16a5cd0cfa8313d6c8c127/conferenceBot/Walter%20E.%20Washington%20Convention%20Center%20Floorplan_Ballroom.png',
    'Ballroom A-C': 'https://raw.githubusercontent.com/CherrelleTucker/snwg-automation/d126cf7abf0aed484e16a5cd0cfa8313d6c8c127/conferenceBot/Walter%20E.%20Washington%20Convention%20Center%20Floorplan_Ballroom.png',
  };

  // Check for exact match first
  if (locationMap[cleanLocation]) {
    return `<${locationMap[cleanLocation]}|${cleanLocation}>`;
  }

  // Check for location containing "Convention Center"
  if (cleanLocation.includes('Convention Center')) {
    return `<${locationMap['Convention Center']}|${cleanLocation}>`;
  }

  // Check for location containing "Marriott"
  if (cleanLocation.includes('Marriott')) {
    return `<${locationMap['Marriott Marquis']}|${cleanLocation}>`;
  }

  // Return unlinked location if no match found
  return cleanLocation;
}

/**
 * Update the formatEventsMessage function to use the new location formatting
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
        const location = formatLocation(event.location);
        const presenter = extractPresenter(event);
        const sessionLink = extractSessionLink(event);

        let messageText = [
          `• ${eventStartTime} - ${eventEndTime}`,
          `*${event.title}* (${presentationType})`,
          `Presenter: ${presenter}`,
          `Location: ${location}`
        ];

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

    // Add virtual tour links in footer
    blocks.push({
      type: "divider"
    });
    blocks.push({
      type: "section",
      text: {
        type: "mrkdwn",
        text: "• <https://truetour.app/properties/1443/experiences/50217?shareId=228365|Marriott Virtual Tour>\n• <https://eventsdc.com/venue/walter-e-washington-convention-center/virtual-tour|Convention Center Virtual Tour>"
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
 * Handles app_home_opened events
 * @param {Object} event - The event payload from Slack
 */
function handleAppHomeOpened(event) {
  if (event.tab === 'home') {
    publishHomeView(event.user);
  }
}

/**
 * Publishes the Home tab view for a user
 * @param {string} userId - The Slack user ID
 */
function publishHomeView(userId) {
  try {
    const view = {
    "type": "home",
    "blocks": [
      {
        "type": "header",
        "text": {
          "type": "plain_text",
          "text": "🎉 Welcome to the AGU24 Conference Bot!",
          "emoji": true
        }
      },
      {
        "type": "section",
        "text": {
          "type": "mrkdwn",
          "text": "Stay organized and navigate AGU Fall Meeting 2024 effortlessly with real-time event updates and venue information."
        }
      },
      {
        "type": "divider"
      },
      {
        "type": "header",
        "text": {
          "type": "plain_text",
          "text": "📅 Conference Navigation",
          "emoji": true
        }
      },
      {
        "type": "section",
        "text": {
          "type": "mrkdwn",
          "text": "*Available Commands:*\n• `/confnow` - See current events\n• `/confnext` - View next hour's events\n• `/conftoday` - Check today's schedule\n• `/conftomorrow` - Preview tomorrow's events\n• `/confthisweek` - View this week's schedule\n• `/confnextweek` - Preview next week's events"
        }
      },
      {
        "type": "section",
        "text": {
          "type": "mrkdwn",
          "text": "*Filter by Session Type:*\nAdd an acronym after any command to filter events\nExample: `/conftoday VEDA` or `/confnow SNWG`"
        }
      },
      {
        "type": "divider"
      },
      {
        "type": "header",
        "text": {
          "type": "plain_text",
          "text": "🗺️ Venue Resources",
          "emoji": true
        }
      },
      {
        "type": "section",
        "text": {
          "type": "mrkdwn",
          "text": "Use `/confmaps` to access:\n• Convention Center floor plans\n• Marriott Marquis layouts\n• Virtual venue tours\n• DC Metro map"
        }
      },
      {
        "type": "divider"
      },
      {
        "type": "context",
        "elements": [
          {
            "type": "mrkdwn",
            "text": "Need help? Message this bot directly or use any command for real-time assistance."
          }
        ]
      }
    ]
  };
      const options = {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify({
        user_id: userId,
        view: view
      }),
      headers: {
        'Authorization': `Bearer ${PropertiesService.getScriptProperties().getProperty('SLACK_BOT_TOKEN')}`
      },
      muteHttpExceptions: true
    };

    UrlFetchApp.fetch('https://slack.com/api/views.publish', options);
    Logger.log(`Home view published for user: ${userId}`);
  } catch (error) {
    Logger.log('Error publishing home view: ' + error.message);
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
