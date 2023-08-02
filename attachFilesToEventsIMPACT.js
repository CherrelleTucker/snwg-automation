// Purpose:  To automate and enhance the management of events in a Google Calendar associated with specific projects and sprints, namely Sprint Reviews and PI Planning Week events. It interacts with Google Calendar and Google Drive to find and populate relevant files in the description field of their Google calendar event in the shared IMPACT PI Calendar based on the event titles.

// To note: 
// This script is created to be a stand-alone script within the Google Apps Script environment. A stand-alone script is a script project that exists independently. This script  may be used as a library or included in other projects, however it is a self-contained script with its own set of functions and code logic. Stand-alone scripts can be created to perform specific tasks, automate workflows, manipulate data, or interact with various Google services like Google Sheets, Google Drive, Google Calendar, etc.

// Use instructions: 
// Open a new or existing Google Sheets file where you want to use the script.
// Click on "Extensions" in the top menu and select "Apps Script" from the dropdown menu. This will open the Google Apps Script editor in a new tab.
// Copy and paste the provided script into the Apps Script editor, replacing the existing code (if any).
// Before running the script, ensure that the script contains the correct Google Calendar ID and folder IDs for the "IMPACT Sprint" and "IMPACT PI Planning" search folders. These IDs are specified in the global constants CALENDAR_ID, SPRINT_SEARCH_FOLDER_ID, and PI_SEARCH_FOLDER_ID.
// Save the script file in the Apps Script editor by clicking on the floppy disk icon or pressing Ctrl + S (Windows/Linux) or Cmd + S (Mac).
// To run the script, click on the play button ▶️ or use the keyboard shortcut Ctrl + Enter (Windows/Linux) or Cmd + Enter (Mac) while the cursor is inside the updateCalendarEvents() function.
// The script will start processing the events in the specified Google Calendar. For "Sprint Review" events, it will extract the Sprint number from the event title and search for relevant files in the designated "IMPACT Sprint" search folder. If it finds relevant files, it will update the event's description with links to those files. Similarly, for "PI Planning" events, it will extract the PI number and identify key phrases from the event title. It will then search for relevant files in the designated "IMPACT PI Planning" search folder based on the extracted information and update the event's description with links to those files.
// After running the script, you can check the logs in the script editor to view the progress and outcomes of the processing for each event.

// Note:
// Make sure that the Google Calendar, as well as the folders mentioned in the global constants, have the necessary permissions so that the script can access and modify event descriptions and files within those folders. Also, ensure that the script is using the correct calendar ID and folder IDs for the "IMPACT" project. Modify the script as needed to fit your specific project's requirements

/////////////////////////////

// Global constants
var CALENDAR_ID = "c_e6e532cefc5ddfdd7f3c715e7a07326607cd240d951991f6a4e3b87653e67ef3@group.calendar.google.com"; // IMPACT PI calendar
var SPRINT_SEARCH_FOLDER_ID = "1UmjkjY5RTRYFOQEt10mwU8trJQ389Jum"; // IMPACT Sprint search folder with subfolders
var PI_SEARCH_FOLDER_ID = "169W64yI042Q24q4socXa4GhiQ7iY4a1f";
var FY_FOLDERS = ['FY23', 'FY24']; // IMPACT PI Planning search folder with subfolders

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