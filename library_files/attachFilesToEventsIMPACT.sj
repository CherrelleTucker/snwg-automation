// Purpose: populate relevant Sprint Review and PI Planning files in the description field of their Google calendar event. 

// source file: IMPACT PI calendar id c_e6e532cefc5ddfdd7f3c715e7a07326607cd240d951991f6a4e3b87653e67ef3@group.calendar.google.com

// IMPACT Sprint search folder: 1UmjkjY5RTRYFOQEt10mwU8trJQ389Jum
  // IMPACT Sprint subfolders to search: FY23 and FY24 if exist

// IMPACT PI Planning search folder: 169W64yI042Q24q4socXa4GhiQ7iY4a1f
  // IMPACT PI Planning subfolders to search: FY23 and FY24 if exist

// Attach Sprint Review files to sprint review events:
// Get Sprint Review Events from Calendar
// Identify Sprint Review Events (Check: Sprint reviews are never more than one day)
// Identify Sprint number in event Title. Event title format includes "PI" 'FY.Q' "Sprint" 'S', where FY is Fiscal Year, Q = Program Increment, and S = Sprint. The sprint number format is FY.PI.S
// Access IMPACT Sprint search folder and subfolders.
// File title formats vary, but always have the sprintNumber in FY.Q.S format. Match Sprint number from the event title to files with that number in the title. 
// If calendar event description already contains the information, stop.
// If calendar event description is empty, copy those file titles and file links to the calendar event description. 
// Repeat until all available events have been searched, then stop.

// Attach PI Planning files to their calendar event. 
// Get PI Planning Events from Calendar.
// There are three PI Planning events each PI. Each has a key phrase to match, in addition to the PI number. keyPhrase1 = "Welcome""Agenda" keyPhrase2 = "Review" keyPhrase3 = "Final""Retro""Retrospective"
// Identify PI Planning Events (Check: PI Planning events are never more than one day)
// Identify PI number in event Title. Event title format includes "PI FY.Pi", where FY is Fiscal Year, Q = Program Increment. The PI number format is FY.Q. 
// Access IMPACT PI Planning search folder and subfolders.
// File title formats vary, but always have piNumber in FY.Q format with no numbers following. Match PI number and keyPhrase from the event title to files with that number and phrase somewhere in the title. 
// If calendar event description already contains the information, stop.
// If calendar event description is empty, copy those file titles and file links to the calendar event description. 
// Repeat until all available events have been searched, then stop.


// Define your constants
var CALENDAR_ID = "c_e6e532cefc5ddfdd7f3c715e7a07326607cd240d951991f6a4e3b87653e67ef3@group.calendar.google.com";
var SPRINT_SEARCH_FOLDER_ID = "1UmjkjY5RTRYFOQEt10mwU8trJQ389Jum";
var PI_SEARCH_FOLDER_ID = "169W64yI042Q24q4socXa4GhiQ7iY4a1f";
var FY_FOLDERS = ['FY23', 'FY24'];

// Testing function. Delete before script deployment
function clearEventDescriptions() {
  // Get Calendar
  var calendar = CalendarApp.getCalendarById(CALENDAR_ID);

 // Get all events from a very early date to a very late date
  var events = calendar.getEvents(new Date(2000, 0, 1), new Date(2100, 0, 1));

  // Clear the description of each event
  events.forEach(function(event) {
    event.setDescription('');
  });
}

// Main function
function updateCalendarEvents() {
  Logger.log("Starting updateCalendarEvents");
  // Get Calendar
  var calendar = CalendarApp.getCalendarById(CALENDAR_ID);
  
  // Process Sprint Review Events
  Logger.log("Getting Sprint Review Events");
  var sprintEvents = getSprintReviewEvents(calendar);
  sprintEvents.forEach(function(event) {
  Logger.log("Processing Sprint Event: " + event.getTitle());
    processSprintEvent(event, SPRINT_SEARCH_FOLDER_ID);
  });
  
  // Process PI Planning Events
  var piEvents = getPIPlanningEvents(calendar);
  piEvents.forEach(function(event) {
    processPIEvent(event, PI_SEARCH_FOLDER_ID);
  });
}

// Returns all sprint review events
function getSprintReviewEvents(calendar) {
  var now = new Date();
  var past = new Date();
  past.setMonth(now.getMonth() - 6);
  var future = new Date();
  future.setMonth(now.getMonth() + 3);
  var events = calendar.getEvents(past, future);
  return events.filter(function(event) {
    return event.getTitle().includes("Sprint Review");
  });
}

// Process a single sprint event
function processSprintEvent(event, folderId) {
  Logger.log("Starting processSprintEvent for: " + event.getTitle());
  var description = event.getDescription();
  var title = event.getTitle();
  var match = title.match(/(\d{2}\.\d{1}) Sprint (\d{1})/); // Adjusted the regex pattern
  if (match) {
    var sprintNumber = match[1] + '.' + match[2]; // Construct the sprint number
    Logger.log("Sprint number found in title: " + sprintNumber);
    var files = getFilesFromFolder(folderId);
    Logger.log("Files found: " + files.map(function(file) { return file.getName(); }).join(", "));
    var relevantFiles = files.filter(function(file) {
      return file.getName().includes(sprintNumber);
    });
    if (relevantFiles.length > 0) {
      var relevantFileNames = relevantFiles.map(function(file) { return file.getName(); });
      if (!relevantFileNames.some(function(name) { return description.includes(name); })) {
        Logger.log("Relevant files found, updating description");
        event.setDescription(description + "\n" + formatFilesForDescription(relevantFiles));
      } else {
        Logger.log("Description already contains relevant file names");
      }
    } else {
      Logger.log("No relevant files found");
    }
  } else {
    Logger.log("No sprint number found in title");
  }
}

// Returns all PI planning events
function getPIPlanningEvents(calendar) {
  var now = new Date();
  var past = new Date();
  past.setMonth(now.getMonth() - 6);
  var future = new Date();
  future.setMonth(now.getMonth() + 3);
  var events = calendar.getEvents(past, future);
  return events.filter(function(event) {
    return event.getTitle().includes("PI Planning");
  });
}

// Process a single PI event
function processPIEvent(event, folderId) {
  Logger.log("Starting processPIEvent for: " + event.getTitle());
  var description = event.getDescription();
  var title = event.getTitle();
  var match = title.match(/\d{2}\.\d{1}/);  // Adjusted the regex pattern
  
  // Define key phrases
  var keyPhrases = ["Welcome", "Management", "Final"];
  
  // Define key phrase mapping
  var keyPhraseMapping = {
    "Welcome": ["Welcome", "Agenda"],
    "Management": ["Management"],
    "Final": ["Final", "Retro", "Retrospective"]
  };
  
  // Identify the key phrase in the event title
  var keyPhrase = keyPhrases.find(function(phrase) {
    return title.includes(phrase);
  });
  
  if (match && !description.includes(match[0]) && keyPhrase) {
    Logger.log("Match and key phrase found in title, getting files");
    var files = getFilesFromFolder(folderId);
    var relevantFiles = files.filter(function(file) {
      return file.getName().includes(match[0]) && keyPhraseMapping[keyPhrase].some(function(phrase) { return file.getName().includes(phrase); });
    });
    if (relevantFiles.length > 0) {
      var relevantFileNames = relevantFiles.map(function(file) { return file.getName(); });
      if (!relevantFileNames.some(function(name) { return description.includes(name); })) {
        Logger.log("Relevant files found, updating description");
        event.setDescription(description + "\n" + formatFilesForDescription(relevantFiles));
      } else {
        Logger.log("Description already contains relevant file names");
      }
    } else {
      Logger.log("No relevant files found");
    }
  } else {
    Logger.log("No match or key phrase found in title");
  }
}


// Utility function to get files from a folder and its subfolders
function getFilesFromFolder(folderId) {
  var folder = DriveApp.getFolderById(folderId);
  return getAllFilesInFolder(folder);
}

// Recursive helper function to get all files in a folder and its subfolders
function getAllFilesInFolder(folder) {
  var files = [];
  var iterator = folder.getFiles();
  while (iterator.hasNext()) {
    files.push(iterator.next());
  }
  var subfolders = folder.getFolders();
  while (subfolders.hasNext()) {
    files = files.concat(getAllFilesInFolder(subfolders.next()));
  }
  return files;
}

// Utility function to format file titles and links for event description
function formatFilesForDescription(files) {
  return files.map(function(file) {
    return file.getName() + ": " + file.getUrl();
  }).join("\n");
}
