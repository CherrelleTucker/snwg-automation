// Purpose: populate relevant Sprint Review and PI Planning files in the description field of their Google calendar event. 

// source file: IMPACT PI calendar id c_e6e532cefc5ddfdd7f3c715e7a07326607cd240d951991f6a4e3b87653e67ef3@group.calendar.google.com

// IMPACT Sprint search folder with subfolders: 1UmjkjY5RTRYFOQEt10mwU8trJQ389Jum
// IMPACT PI Planning search folder with subfolders: 169W64yI042Q24q4socXa4GhiQ7iY4a1f

// Global constants
var CALENDAR_ID = "c_e6e532cefc5ddfdd7f3c715e7a07326607cd240d951991f6a4e3b87653e67ef3@group.calendar.google.com";
var SPRINT_SEARCH_FOLDER_ID = "1UmjkjY5RTRYFOQEt10mwU8trJQ389Jum";
var PI_SEARCH_FOLDER_ID = "169W64yI042Q24q4socXa4GhiQ7iY4a1f";
var FY_FOLDERS = ['FY23', 'FY24'];

// Primary function to execute search and attach actions for PI and Sprint events
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

// Helper function to return all sprint review events from the past six months to the next three months
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

// Helper function to process a single sprint event by extracting the Sprint number from the event title, search for relevant files in the designated Sprint folder, and update the event's description if necessary.
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

// Helper function to return all PI planning events from the past six months to the next three months
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

// Helper function to process a single PI event by extracting the PI number from the event title, search for relevant files in the designated PI folder, and update the event's description if necessary.
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

// Helper wrapper function that gets all the files from the specified folder by passing its ID.Provides a convenient way to access files from a particular folder without having to call the recursive 
function getFilesFromFolder(folderId) {
  var folder = DriveApp.getFolderById(folderId);
  return getAllFilesInFolder(folder);
}

// Helper recursive function to iterate through files in the current folder and recursively goes through each subfolder to collect all the files.
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

// Helper function to format file titles and links for event description
function formatFilesForDescription(files) {
  return files.map(function(file) {
    return file.getName() + ": " + file.getUrl();
  }).join("\n");
}
