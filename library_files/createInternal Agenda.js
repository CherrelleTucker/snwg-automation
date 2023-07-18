// Purpose: Create new weekly agenda for the SNWG Internal Planning Meeting in the Weekly meeting folder
// Done: Duplicate template agenda
// Done: Document name populate with Monday following the creation
// Done: Populate previousAgenda, most recent Monthly status update, most recent OPERA tag up, next OPERA tag up, current DMPR
// Done: Internal date formats in Day Month, Date, Year format. 
// Done: call LibraryCurrentPICalculation to populate {{Current PI}}
 
function createNewInternalAgenda() {
  // Set the ID of the template document
  var templateId = "1tE6xNFeMLVpcGwWMB9GuYpGom5F4Bi_81dUsp_W3jDQ";

  // Make a copy of the template document
  var newDocument = DriveApp.getFileById(templateId).makeCopy();

  // Get the new document's ID
  var newDocumentId = newDocument.getId();

  // Open the new document
  var document = DocumentApp.openById(newDocumentId);

  // Set the new document's name with the current date (Monday following creation)
  var currentDate = getMondayFollowingDate(new Date());
  var newDocumentName = currentDate + " SNWG MO Internal Planning Meeting";
  document.setName(newDocumentName);

  // Replace placeholders with links and date
  var previousAgendaFolderId = "1WKYw4jnP6ejRkOLAIPoPvbEYClaLE4eR";
  var operaTagUpFolderId = "1M3EMWLCxhkqcPKLmH7grDu2zhSEuOvmc";
  var snwgMonthlyFolderId = "1HPjhc2LADvS9j3W_K3riq4RQPBngfqGY"; //folder ID is located in 
  var dmprFolderId = "1y2vjwf52HBJpTeSIg7sYPaLSSzGAnWZU";
  var nextOperaTagUpLink = getOperaTagUpLink(operaTagUpFolderId, newDocumentId, true);
  var previousOperaTagUpLink = getOperaTagUpLink(operaTagUpFolderId, newDocumentId, false);
  
  var previousAgendaLink = getMostRecentFileLink(previousAgendaFolderId, newDocumentId, "{{Link to Previous Agenda}}");
  var previousOperaTagUpLink = getMostRecentFileLink(operaTagUpFolderId, newDocumentId, "{{link to last OPERA tag up}}");
  var snwgMonthlyLink = getMostRecentFileLink(snwgMonthlyFolderId, newDocumentId, "{{link to last SNWG/NASA monthly}}");
  var dmprLink = getDMPRLink(dmprFolderId, newDocumentId, currentDate);
  var currentPI = LibraryCurrentPICalculation.getCurrentPI(); // Call the library function to get the current PI
    

  var documentBody = document.getBody();
  
  // Replace the placeholders with hyperlinks
  replaceWithHyperlink(documentBody, "{{Link to Previous Agenda}}", previousAgendaLink);
  replaceWithHyperlink(documentBody, "{{link to last OPERA tag up}}", previousOperaTagUpLink);
  replaceWithHyperlink(documentBody, "{{link to next OPERA tag up}}", nextOperaTagUpLink);
  replaceWithHyperlink(documentBody, "{{link to last SNWG/NASA monthly}}", snwgMonthlyLink);
  replaceWithHyperlink(documentBody, "{{link to current DMPR}}", dmprLink);

  documentBody.replaceText("{{Current PI}}", currentPI);  // Replace the placeholder with the actual current PI
  
 // Format the date as "Day, Month Date, Year" and replace the placeholder
var dateForInternal = new Date(currentDate);
dateForInternal.setDate(dateForInternal.getDate() + 1);
var formattedDate = Utilities.formatDate(dateForInternal, Session.getScriptTimeZone(), "EEEE, MMMM dd, yyyy");
documentBody.replaceText("{{Internal Date}}", formattedDate);  
}

  function getMondayFollowingDate(date) {
    // Find the next Monday from the given date
    var day = date.getDay();
    var diff = (day === 0 ? 1 : 8) - day;
    var nextMonday = new Date(date.getTime() + diff * 24 * 60 * 60 * 1000);
  
    // Format the date as "YYYY-MM-DD"
    var formattedDate = Utilities.formatDate(nextMonday, Session.getScriptTimeZone(), "yyyy-MM-dd");
  
    return formattedDate;
 
  }

 function getOperaTagUpLink(folderId, excludeId, isFuture) {
  var folder = DriveApp.getFolderById(folderId);
  var files = folder.getFiles();
  var latestFile = null;
  var latestFileTitle = '';
  var today = new Date();
  today.setHours(0,0,0,0);  // set time to 00:00:00.000

  while (files.hasNext()) {
    var file = files.next();
    var fileTitle = file.getName();
    var fileId = file.getId();
    var fileDate = file.getLastUpdated();

    // Skip files based on isFuture flag
    if (isFuture && fileDate < today) {
        continue;
    }
    if (!isFuture && fileDate > today) {
        continue;
    }

    if (!fileTitle.includes('Template') && fileId !== excludeId) {
      if (latestFile === null || fileTitle > latestFileTitle) {
        latestFile = file;
        latestFileTitle = fileTitle;
      }
    }
  }

  if (latestFile !== null) {
    return latestFile.getUrl();
  } else {
    return '';
  }
}

//Create Custom Get Action menu on documnent open
function onOpen() {
  DocumentApp.getUi()
  .createMenu('Action Tracking')
  .addItem('Run Function', 'ActionItemTablePopulationv10')
  .addToUi();
}

// Function to get the link of the most recent OPERA file excluding future files
function getMostRecentFileLink(folderId, excludeId, placeholderText) {
  var folder = DriveApp.getFolderById(folderId);
  var files = folder.getFiles();
  var newestFile = null;
  var newestFileName = '';
  var todayStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");

  while (files.hasNext()) {
    var file = files.next();
    var fileName = file.getName();
    var fileId = file.getId();

    // Skip files whose title date is in the future
    if (fileName.localeCompare(todayStr) > 0) {
      continue;
    }

    // If this file is more recent than our current newest file, and is not the template or excluded file, update newestFile and newestFileName
    if (fileName.localeCompare(newestFileName) > 0 && fileId !== excludeId && !fileName.includes("Template")) {
      newestFile = file;
      newestFileName = fileName;
    }
  }

  if (newestFile !== null) {
    var placeholderLink = "{{Link to Previous Agenda}}";
    if (placeholderText === placeholderLink && newestFile.getId() === excludeId) {
      return "";
    } else {
      return newestFile.getUrl();
    }
  } else {
    return "";
  }
}

// Function to get the link of the next OPERA file (i.e., future file)
function getOperaTagUpLink(folderId, excludeId, isFuture) {
  var folder = DriveApp.getFolderById(folderId);
  var files = folder.getFiles();
  var relevantFile = null;
  var relevantFileName = '';
  var todayStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");

  while (files.hasNext()) {
    var file = files.next();
    var fileName = file.getName();
    var fileId = file.getId();

    if (isFuture) {
      // Skip files whose title date is in the past or today
      if (fileName.localeCompare(todayStr) <= 0) {
        continue;
      }

      // If this file is sooner than our current relevant file, and is not the template or excluded file, update relevantFile and relevantFileName
      if (relevantFileName === '' || fileName.localeCompare(relevantFileName) < 0) {
        relevantFile = file;
        relevantFileName = fileName;
      }
    } else {
      // Skip files whose title date is in the future
      if (fileName.localeCompare(todayStr) > 0) {
        continue;
      }

      // If this file is more recent than our current relevant file, and is not the template or excluded file, update relevantFile and relevantFileName
      if (fileName.localeCompare(relevantFileName) > 0) {
        relevantFile = file;
        relevantFileName = fileName;
      }
    }
  }

  if (relevantFile !== null) {
    return relevantFile.getUrl();
  } else {
    return '';
  }
}
  
  function getDMPRLink(folderId, excludeId, currentDate) {
    var folder = DriveApp.getFolderById(folderId);
    var files = folder.getFiles();
    var dmprFile = null;
    var dmprFileMonth = currentDate.substring(0, 7);
  
    while (files.hasNext()) {
      var file = files.next();
      var fileName = file.getName();
  
      if (fileName.startsWith(dmprFileMonth) && file.getId() !== excludeId && !fileName.includes("Template")) {
        dmprFile = file;
        break;
      }
    }
  
    if (dmprFile !== null) {
      return dmprFile.getUrl();
    } else {
      return "";
    }
  }
 function replaceWithHyperlink(documentBody, placeholderText, url) {
  var foundElement = documentBody.findText(placeholderText);
  
  if (foundElement) {
    var startOffset = foundElement.getStartOffset();
    var endOffset = foundElement.getEndOffsetInclusive();
    var textElement = foundElement.getElement().asText();
  
    // Check if url is not empty
    if (url !== '') {
        // Extract file ID from the URL
        var fileId = url.split('/')[5];
        // Get the file
        var file = DriveApp.getFileById(fileId);
        // Get the file name
        var fileName = file.getName();
    } else {
        var fileName = 'Not found';
    }
      
    textElement.deleteText(startOffset, endOffset);
    textElement.insertText(startOffset, fileName).setLinkUrl(startOffset, startOffset + fileName.length - 1, url);
  }
}