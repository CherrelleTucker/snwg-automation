// Purpose: create a new SNWG Internal Weekly Planning Meeting agenda based on the template, replacing placeholders with appropriate data.

////////////////////////////////////////////////

// Primary function to create and populate a new Internal Planning meeting agenda
function createNewInternalAgenda() {
  var templateId = "1tE6xNFeMLVpcGwWMB9GuYpGom5F4Bi_81dUsp_W3jDQ";
  var newDocument = createCopyOfTemplate(templateId);
  var newDocumentId = newDocument.getId();
  var document = DocumentApp.openById(newDocumentId);

  // Get the date for the next Monday following the current date
  var currentDate = getMondayFollowingDate(new Date());
  var newDocumentName = currentDate + " SNWG MO Internal Planning Meeting";
  document.setName(newDocumentName);

  // IDs of folders where specific files are stored
  var previousAgendaFolderId = "1WKYw4jnP6ejRkOLAIPoPvbEYClaLE4eR";
  var operaTagUpFolderId = "1M3EMWLCxhkqcPKLmH7grDu2zhSEuOvmc";
  var snwgMonthlyFolderId = "1HPjhc2LADvS9j3W_K3riq4RQPBngfqGY";
  var dmprFolderId = "1y2vjwf52HBJpTeSIg7sYPaLSSzGAnWZU";

  // Get links to specific files from their respective folders
  var nextOperaTagUpLink = getMostRecentFileLink(operaTagUpFolderId, newDocumentId, true);
  var previousAgendaLink = getMostRecentFileLink(previousAgendaFolderId, newDocumentId, false);
  var previousOperaTagUpLink = getMostRecentFileLink(operaTagUpFolderId, newDocumentId, false);
  var snwgMonthlyLink = getMostRecentFileLink(snwgMonthlyFolderId, newDocumentId, false);
  var dmprLink = getDMPRLink(dmprFolderId, currentDate);

  // Get the current PI (Program Increment) for the document using the PI calculator library script
  var currentPI = currentPIcalculator.getCurrentPI();
  var documentBody = document.getBody();

  // Replace placeholders in the document body with the obtained links and PI information
  replaceWithHyperlink(documentBody, "{{Link to Previous Agenda}}", previousAgendaLink);
  replaceWithHyperlink(documentBody, "{{link to last OPERA tag up}}", previousOperaTagUpLink);
  replaceWithHyperlink(documentBody, "{{link to next OPERA tag up}}", nextOperaTagUpLink);
  replaceWithHyperlink(documentBody, "{{link to last SNWG/NASA monthly}}", snwgMonthlyLink);
  replaceWithHyperlink(documentBody, "{{link to current DMPR}}", dmprLink);

  documentBody.replaceText("{{Current PI}}", currentPI); 

  replaceWithFormattedDate(documentBody, "{{Internal Date}}", currentDate);
}

// Helper Function: Create a copy of a template document
function createCopyOfTemplate(templateId) {
  return DriveApp.getFileById(templateId).makeCopy();
}

// Helper Function: Get the Monday following a given date
function getMondayFollowingDate(date) {
  var day = date.getDay();
  var diff = (day === 0 ? 1 : 8) - day;
  var nextMonday = new Date(date.getTime() + diff * 24 * 60 * 60 * 1000);
  var formattedDate = Utilities.formatDate(nextMonday, Session.getScriptTimeZone(), "yyyy-MM-dd");
  return formattedDate;
}

// Helper Function: Get the link of the most recent or next file from a folder
function getMostRecentFileLink(folderId, excludeId, isFuture) {
  var folder = DriveApp.getFolderById(folderId);
  var files = folder.getFiles();
  var relevantFile = null;
  var relevantFileName = '';
  var todayStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");

  while (files.hasNext()) {
    var file = files.next();
    var fileName = file.getName();
    var fileId = file.getId();
    relevantFile = getRelevantFile(isFuture, todayStr, fileName, fileId, excludeId, relevantFile, relevantFileName);
    if (relevantFile) relevantFileName = fileName;
  }

  return relevantFile ? relevantFile.getUrl() : '';
}

// Helper Function: Get the relevant file based on the specified conditions
function getRelevantFile(isFuture, todayStr, fileName, fileId, excludeId, currentRelevantFile, currentRelevantFileName) {
  var fileIsRelevant = (isFuture && fileName.localeCompare(todayStr) > 0) || (!isFuture && fileName.localeCompare(todayStr) <= 0);
  var fileIsMoreRecent = currentRelevantFileName === '' || (isFuture ? fileName.localeCompare(currentRelevantFileName) < 0 : fileName.localeCompare(currentRelevantFileName) > 0);
  if (fileIsRelevant && fileIsMoreRecent && fileId !== excludeId && !fileName.includes("Template")) {
    return DriveApp.getFileById(fileId);
  }
  return currentRelevantFile;
}


// Helper Function: Get the link of the DMPR file corresponding to a given month
function getDMPRLink(folderId, currentDate) {
  var folder = DriveApp.getFolderById(folderId);
  var files = folder.getFiles();
  var dmprFile = null;
  var dmprFileMonth = currentDate.substring(0, 7);

  while (files.hasNext()) {
    var file = files.next();
    var fileName = file.getName();
    if (fileName.startsWith(dmprFileMonth) && !fileName.includes("Template")) {
      dmprFile = file;
      break;
    }
  }

  return dmprFile ? dmprFile.getUrl() : '';
}

// Helper Function: Replace placeholder text in a document with a hyperlink
function replaceWithHyperlink(documentBody, placeholderText, url) {
  var foundElement = documentBody.findText(placeholderText);
  if (foundElement) {
    var startOffset = foundElement.getStartOffset();
    var endOffset = foundElement.getEndOffsetInclusive();
    var textElement = foundElement.getElement().asText();
    if (url !== '') {
      var fileId = url.split('/')[5];
      var file = DriveApp.getFileById(fileId);
      var fileName = file.getName();
    } else {
      var fileName = 'Not found';
    }
    textElement.deleteText(startOffset, endOffset);
    textElement.insertText(startOffset, fileName).setLinkUrl(startOffset, startOffset + fileName.length - 1, url);
  }
}

// Helper Function: Replace a placeholder with a formatted date
function replaceWithFormattedDate(documentBody, placeholderText, currentDate) {
  var dateForInternal = new Date(currentDate);
  dateForInternal.setDate(dateForInternal.getDate() + 1);
  var formattedDate = Utilities.formatDate(dateForInternal, Session.getScriptTimeZone(), "EEEE, MMMM dd, yyyy");
  documentBody.replaceText(placeholderText, formattedDate);
}
