// Purpose: To automate the creation of slides for weekly project updates using a predefined template. This script calculates the current project iteration, adjusting for events like holidays, and populates the slide with data from Google Drive, including agendas from the current and previous weeks. Users can manually initiate slide creation via a custom Google Slides menu, and the script also interfaces with Google Sheets to manage open action items, streamlining the entire update process.

// Future development: use file name instead of hyperlink text
//  Consideration: limitations of Google Slides Object handling does not promote the hyperlinking of individual text strings within objects. 

// To use:
// 1. **Setting Up Google Drive Folders**: Before using the script, ensure you have folders set up in Google Drive to store the agendas and other related files. Note down the Folder IDs of these folders. You can find the Folder ID in the URL of the Google Drive folder.
// 2. **Accessing Google Apps Script**: Open your Google Drive. Click on "New" > "More" > "Google Apps Script" to access the Google Apps Script editor.
// 3. **Copying and Pasting Script**: In the script editor, erase any existing code and paste the provided script.
// 4. **Configuring Folder IDs**: Locate the lines with folder ID assignments, such as `var DMPRfolderID = '...';` and `var agendasFolderID = '...';`. Replace the predefined Folder IDs with those from your Google Drive folders, as noted in Step 1.
// 5. **Running the Script**: After saving your changes (using the floppy disk icon or pressing Ctrl + S/Cmd + S), close the script editor.
// 6. **Setting up a Custom Menu Trigger**: The script provides a custom menu in Google Slides to manually create a new slide. To use this, open your target Google Slides presentation and look for the "Update Slides" menu. Clicking on it will give an option to "Create a New Slide".
// 7. **Authorization and Permissions**: On the first run, you will be prompted for authorization. Navigate through the authorization process to grant the script the necessary permissions.
// 8. **Viewing Logs**: To monitor the script's progress or diagnose any issues, access the logs in the script editor by clicking on the bug icon in the toolbar.
// 9. **Review the Output**: The script will automatically create a new slide in the target presentation, fetching and populating relevant data. Ensure to inspect the generated slides to verify the accuracy of the data.
// The script is now set up to facilitate the streamlined creation of project update slides. Remember to use the custom menu in Google Slides whenever you need to generate a new update.

/////////////////////////////////

// Global variables
var quadTemplateId = '1iWm9oAJ-hFR_BJ-pEFGndhfFCLhWIigxJBvZM5xlf2M'; // SNWG MO Weekly Quad Chart Slide Template
var targetPresentationId = '1wV26VGjZJ_P7hsnxdkub6PltQ3ZkedYb8yDviEYE4aE'; // SNWG MO Weekly Quad Chart presentation
// var dmprfolderID = '1NCH6-V9pMA8pOivX0XD5ZtGLT-8OQZ6A'; // IMPACT/SNWG DMPR folder FY24; retired function
var agendasFolderID = '13zl2CvMNtDMFKcNZetAA00e5tkh3Eo_M'; // SNWG Meeting Notes folder 
var devSeedFolderID = '1Bvj1b1u2LGwjW5fStQaLRpZX5pmtjHSj'; // Assessment/DevSeed Weekly Meeting Agenda Folder
var assessmentHQFolderID = '1dmN0oYQZwGFu83BwOGT90I_GFtGH1aup'; // Assessment/HQ Weekly MEeting Agenda Folder
var operaFolderID = '1urLgf4tRUk4hmbWrxNYDidO7Qwb3BV3B'; // FY25 OPERA Meeting Notes Folder
var sheetID = '1uYgX660tpizNbIy44ddQogrRphfwZqn1D0Oa2RlSYKg'; // SNWG MO Action Tracking Sheet 
var teamCalendar = 'c_365230bc41700e58e23f74b286db1773d395e4bc6807c81a4c78658df5db423e@group.calendar.google.com'; // SNWG Team Google Calendar
var impactPiCalendar = 'c_e6e532cefc5ddfdd7f3c715e7a07326607cd240d951991f6a4e3b87653e67ef3@group.calendar.google.com'; // IMPACT PI calendar

// helper function to create custom menu for manually updating slides
function onOpen() {
  SlidesApp.getUi()
      .createMenu('Update Slides')
      .addItem('Create a New Slide','updatePresentationOptimized')
      .addToUi();
}

// function to get FY.PI.Sprint from IMPACT PI Calendar
function getPiFromImpactPiCalendar() {
    var calendar = CalendarApp.getCalendarById(impactPiCalendar);
    var currentWeekDates = getCurrentWeekDates();
    var events = calendar.getEvents(currentWeekDates.start, currentWeekDates.end);
    var piRegex = /PI \d{2}\.\d Sprint \d/; // Regex to match "PI YY.Q Sprint S" format

    for (var i = 0; i < events.length; i++) {
        var event = events[i];
        if (piRegex.test(event.getTitle())) {
            return event.getTitle();
        }
    }
    return "No PI Event Found"; // Return this if no matching event is found
}

// helper function to duplicate template slide with placeholders
function duplicateTemplateSlide(quadTemplateId) {
  try {
    var quadTemplateFile = DriveApp.getFileById(quadTemplateId);
    var newSlideFile = quadTemplateFile.makeCopy();
    var newSlide = SlidesApp.openById(newSlideFile.getId()).getSlides()[0];
    DriveApp.getFileById(newSlideFile.getId()).setTrashed(true);
    return newSlide;
  } catch (error) {
    Logger.log("Error duplicating template slide: " + error);
    throw error;
  }
}

// helper function to get current week Mon-Fri dates for header
function getCurrentWeekDates() {
  try {
    var today = new Date();
    var startOfWeek = new Date(today.setDate(today.getDate() - today.getDay() + 1));
    var endOfWeek = new Date(today.setDate(today.getDate() - today.getDay() + 5));
    return {
      start: startOfWeek,
      end: endOfWeek,
      formatted: Utilities.formatDate(startOfWeek, 'GMT', 'MM/dd/yy') + " - " + Utilities.formatDate(endOfWeek, 'GMT', 'MM/dd/yy')
    };
  } catch (error) {
    Logger.log("Error getting current week dates: " + error);
    throw error;
  }
}

// helper function to convert dates
function reformatWeekDates(weekDates) {
  var dates = weekDates.formatted.split(' - ');
  return {
    start: new Date(dates[0]),
    end: new Date(dates[1])
  };
}

// Helper function to extract date from filename
function extractDateFromFileName(fileName) {
  var datePattern = /^\d{4}-\d{2}-\d{2}/; // Regex pattern for YYYY-MM-DD
  var match = fileName.match(datePattern);

  if (match) {
    return new Date(match[0] + 'T00:00:00.000Z'); // Append time string
  } else {
    return null; // Return null if no date pattern is found
  }
}

// New helper function to format agenda links
function formatAgendaLink(file) {
  var doc = DocumentApp.openByUrl(file.url);
  return {
    text: doc.getName(),
    url: file.url
  };
}

// Helper function to find current week files, including from subfolders, with certain exclusions
function getFilesForCurrentWeek(folderIDs, currentWeekDates) {
  var filesWithLinks = [];
  var adjustedStartDate = new Date(currentWeekDates.start);
  adjustedStartDate.setDate(adjustedStartDate.getDate() - 1);
  var adjustedEndDate = new Date(currentWeekDates.end);
  adjustedEndDate.setDate(adjustedEndDate.getDate() + 1);

  folderIDs.forEach(function(folderID) {
    try {
      var foldersToProcess = [folderID];
      while (foldersToProcess.length > 0) {
        var currentFolderID = foldersToProcess.pop();
        var folder = DriveApp.getFolderById(currentFolderID);
        var files = folder.getFiles();
        
        while (files.hasNext()) {
          var file = files.next();
          var fileName = file.getName();
          var fileDate = extractDateFromFileName(fileName);
          
          if (fileDate && fileDate >= adjustedStartDate && fileDate <= adjustedEndDate) {
            filesWithLinks.push({
              name: fileName,
              url: file.getUrl()
            });
          }
        }

        var subfolders = folder.getFolders();
        while (subfolders.hasNext()) {
          var subfolder = subfolders.next();
          var subfolderName = subfolder.getName();
          if (!["2020", "2021", "FY21", "2022", "FY22", "2023", "FY23", "Archived Notes"].includes(subfolderName)) {
            foldersToProcess.push(subfolder.getId());
          }
        }
      }
    } catch (e) {
      Logger.log("Error accessing folder with ID: " + folderID + ". Error: " + e.message);
    }
  });
  return filesWithLinks;
}

// Helper function: Get Team Schedule events from the designated calendar
function getCalendarEventsForCurrentWeek() {
  var currentWeek = getCurrentWeekDates(); // Assuming this returns an object with 'start' and 'end' Date objects
  var calendar = CalendarApp.getCalendarById(teamCalendar); // Using global variable teamCalendar
  var events = calendar.getEvents(currentWeek.start, currentWeek.end);

  var eventDetails = [];
  for (var i = 0; i < events.length; i++) {
    var event = events[i];
    
    var startDate = event.getStartTime();
    var endDate = event.getEndTime();

    // Adjust for all-day events
    if (event.isAllDayEvent()) {
      endDate = new Date(endDate.getTime() - 24 * 60 * 60 * 1000); // Subtract one day from the end date
    }

    var formattedStartDate = Utilities.formatDate(startDate, Session.getScriptTimeZone(), "MMMM d");
    var formattedEndDate = Utilities.formatDate(endDate, Session.getScriptTimeZone(), "MMMM d");

    // Check if start date and end date are the same
    var dateRange = (formattedStartDate === formattedEndDate) ? formattedStartDate : (formattedStartDate + " - " + formattedEndDate);
    
    eventDetails.push({
      title: event.getTitle(),
      date: dateRange
    });
  }
  return eventDetails;
}

// helper function to replace slide placeholders
function replacePlaceholdersInSlide(slide, replacements) {
  try {
    var slideShapes = slide.getShapes();
    slideShapes.forEach(function (shape) {
      if (shape.getText) {
        var text = shape.getText().asString();
        for (var key in replacements) {
          if (text.includes(key)) {
            shape.getText().setText(text.replace(key, replacements[key]));
          }
        }
      }
    });
  } catch (error) {
    Logger.log("Error replacing placeholders in slide: " + error);
    throw error;
  }

    // Special handling for Team Schedules
  if (replacements['{{Team Schedules}}']) {
    var teamScheduleEvents = getCalendarEventsForCurrentWeek();
    var scheduleString = teamScheduleEvents.map(function(event) {
      return event.title + ": " + event.date;
    }).join('\n');
    replacements['{{Team Schedules}}'] = scheduleString;
  }
}

// Helper function to find a shape in a slide by its alt text description
function findShapeByAltText(slide, altText) {
  var shapes = slide.getShapes();
  for (var i = 0; i < shapes.length; i++) {
    var shape = shapes[i];
    if (shape.getDescription() === altText) {
      return shape;
    }
  }
  return null;
}

/*Retired function:  helper function to find current month's DMPR file
function findCurrentDMPRFile() {
  // Get the current year and month
  var date = new Date();
  var year = date.getFullYear();
  var month = ("0" + (date.getMonth() + 1)).slice(-2); // Convert to two digits

  // Create the search pattern for the file
  var searchPattern = year + "-" + month + "_SNWG_DMPR";

  // Get the files in the folder with the given folder ID
  var folder = DriveApp.getFolderById(dmprfolderID);
  var files = folder.getFiles();

  // Loop through the files and look for the one with the correct name format
  while (files.hasNext()) {
    var file = files.next();
    if (file.getName().startsWith(searchPattern)) {
      Logger.log("DMPR file found: " + file.getName()); // Log for debugging
      return file; // Return the found file
    }
  }

  // If no matching file is found, return null
  return null;
} */

/*retired function: helper function to find DMPR shape
function handleCurrentDMPRShape(currentDMPRShape) {
  if (currentDMPRShape) {
    Logger.log("Current DMPR shape found.");
    var dmprFile = findCurrentDMPRFile(); // Find the DMPR file
    if (dmprFile) {
      Logger.log("DMPR file found: " + dmprFile.getName());
  var year = new Date().getFullYear();
      var month = ("0" + (new Date().getMonth() + 1)).slice(-2);
      var dmprText = year + "-" + month + " DMPR"; // Format the text
      var textRange = currentDMPRShape.getText(); // Get the text range from the shape
      textRange.setText(dmprText); // Replace the text
      Logger.log("DMPR text set: " + dmprText);

      var dmprFileUrl = dmprFile.getUrl();
      Logger.log("Adding new hyperlink to DMPR file: " + dmprFileUrl);

      currentDMPRShape.setLinkUrl(dmprFileUrl); // Set the hyperlink on the shape itself
    } else {
      Logger.log("DMPR file not found!");
    }
  } else {
    Logger.log("Current DMPR shape not found!");
  }
}*/

/* Retired function: Get the current month's DMPR file and format the text
function getFormattedDmprText() {
    var dmprFile = findCurrentDMPRFile();
    if (dmprFile) {
        var year = new Date().getFullYear();
        var month = ("0" + (new Date().getMonth() + 1)).slice(-2);
        return year + "-" + month + " DMPR";
    }
    return ""; // Return an empty string if no DMPR file is found
}*/

// helper function to pull Open Action Items from the Action Tracking workbook
function updateOpenActionItems() {
  var sheet = SpreadsheetApp.openById(sheetID).getSheetByName('MO');

  if (sheet === null) {
    Logger.log('Could not find sheet with name MO');
    return [];
  }

  // Define the columns
  var statusColumn = 2; // Column B for status
  var assignedToColumn = 3; // Column C for assigned to
  var taskColumn = 4; // Column D for task

  // Get the values from the Sheet
  var dataRange = sheet.getRange(2, 1, sheet.getLastRow(), taskColumn).getValues(); // Adjust range to include all necessary columns
  var openActionItems = dataRange.filter(function(row) {
    return row[statusColumn - 1].toLowerCase() != 'done'; // Filter out rows where status is 'done'
  }).map(function(row) {
    return [row[assignedToColumn - 1], row[taskColumn - 1]]; // Return the assigned to and task columns
  });

  return openActionItems;
}

// function to fetch the list generated in updateOpenActionItems to update the table in the slide
function updateActionItemsTable(newSlide, assignedToTasks) {
  var tables = newSlide.getTables();

  if (tables.length > 0) {
    var table = tables[0];
    var currentNumRows = table.getNumRows();
    var numRowsNeeded = assignedToTasks.length;

    // Remove excess rows if needed
    while (currentNumRows > numRowsNeeded) {
      table.removeRow(currentNumRows - 1);
      currentNumRows--;
    }

    // Add new rows or update existing rows with action items
    for (var row = 0; row < numRowsNeeded; row++) {
      if (row >= currentNumRows) {
        table.appendRow(); // Append a new row if needed
      }
      table.getCell(row, 0).getText().setText(assignedToTasks[row][0]); // Assigned To
      table.getCell(row, 1).getText().setText(assignedToTasks[row][1]); // Task
    }
  } else {
    Logger.log("No table found for action items on the slide.");
  }
}

// Primary function to create and populate new slide
function updatePresentationOptimized() {
  Logger.log('Starting updatePresentationOptimized function');
  
  try {
    // Duplicate the template slide
    Logger.log('Duplicating template slide');
    var newSlide = duplicateTemplateSlide(quadTemplateId);

    // Get Current PI
    Logger.log('Getting current PI');
    var currentPI = getPiFromImpactPiCalendar();

    // Get the current week's dates
    var currentWeekDates = getCurrentWeekDates();
    var weekDatesString = currentWeekDates.formatted;

    // Fetch team schedule events
    Logger.log('Getting team schedule events');
    var teamScheduleEvents = getCalendarEventsForCurrentWeek();
    var teamSchedulesString = teamScheduleEvents.map(function(event) {
      return event.title + ": " + event.date;
    }).join('\n');

    // Fetch agenda files first
    Logger.log('Getting agenda files');
    var folderIDsToSearch = [
      agendasFolderID,
      devSeedFolderID
    ];
    
    var currentWeekAgendas = getFilesForCurrentWeek(folderIDsToSearch, currentWeekDates);
    Logger.log('Found ' + currentWeekAgendas.length + ' agenda files');
    
    // Update placeholders first
    var placeholders = {
      '{{Current Mon to Fri}}': weekDatesString,
      '{{Team Schedules}}': teamSchedulesString,
      '{{Adjusted PI}}': currentPI
    };
    
    Logger.log('Replacing placeholders');
    replacePlaceholdersInSlide(newSlide, placeholders);

    // Handle agenda items
    Logger.log('Processing agenda items');
    var agendaShape = findShapeByAltText(newSlide, "Meetings & Agendas");
    
    if (agendaShape) {
      var textRange = agendaShape.getText();
      var fullText = '';
      var agendaItems = [];
      
      // First, collect all items
      currentWeekAgendas.forEach(function(file, index) {
        try {
          var fileId = file.url.match(/[-\w]{25,}/);
          if (fileId) {
            var driveFile = DriveApp.getFileById(fileId[0]);
            var title = driveFile.getName();
            agendaItems.push({
              title: title,
              url: file.url,
              startIndex: fullText.length,
              endIndex: fullText.length + title.length
            });
            fullText += title;
            if (index < currentWeekAgendas.length - 1) {
              fullText += '\n';
            }
          }
        } catch (e) {
          Logger.log('Error building text: ' + e.toString());
        }
      });
      
      // Set the complete text
      textRange.setText(fullText);
      Logger.log('Text set successfully');
      
      // Apply links to each item's specific range
      agendaItems.forEach(function(item) {
        try {
          Logger.log('Applying link for: ' + item.title + ' at indices ' + item.startIndex + ':' + item.endIndex);
          var itemRange = textRange.getRange(item.startIndex, item.endIndex);
          var linkStyle = itemRange.getTextStyle();
          linkStyle.setForegroundColor('#0000FF');
          linkStyle.setUnderline(true);
          linkStyle.setLinkUrl(item.url);
        } catch (e) {
          Logger.log('Error applying link: ' + e.toString());
        }
      });
    }

    // Insert the new slide
    Logger.log('Inserting new slide');
    var targetPresentation = SlidesApp.openById(targetPresentationId);
    targetPresentation.insertSlide(0, newSlide);
    
    Logger.log('Slide update completed successfully');

  } catch (error) {
    Logger.log('Error in updatePresentationOptimized: ' + error.toString());
    throw error;
  }
}
