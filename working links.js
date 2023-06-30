function copyNewestDocument() {
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
    var snwgMonthlyFolderId = "1HPjhc2LADvS9j3W_K3riq4RQPBngfqGY";
    var dmprFolderId = "1y2vjwf52HBJpTeSIg7sYPaLSSzGAnWZU";
    var nextOperaTagUpLink = getNextOperaTagUpLink(operaTagUpFolderId, newDocumentId);
  
    var previousAgendaLink = getMostRecentFileLink(previousAgendaFolderId, newDocumentId, "{{Link to Previous Agenda}}");
    var operaTagUpLink = getMostRecentFileLink(operaTagUpFolderId, newDocumentId, "{{link to most recent OPERA tag up}}");
    var snwgMonthlyLink = getMostRecentFileLink(snwgMonthlyFolderId, newDocumentId, "{{link to most recent SNWG/NASA monthly}}");
    var dmprLink = getDMPRLink(dmprFolderId, newDocumentId, currentDate);
  
    var documentBody = document.getBody();
    
    // Replace the placeholders with hyperlinks
    replaceWithHyperlink(documentBody, "{{Link to Previous Agenda}}", previousAgendaLink);
    replaceWithHyperlink(documentBody, "{{link to most recent OPERA tag up}}", operaTagUpLink);
    replaceWithHyperlink(documentBody, "{{link to next OPERA tag up}}", nextOperaTagUpLink);
    replaceWithHyperlink(documentBody, "{{link to most recent SNWG/NASA monthly}}", snwgMonthlyLink);
    replaceWithHyperlink(documentBody, "{{link to current DMPR}}", dmprLink);
    
    documentBody.replaceText("{{Internal Date}}", currentDate);
  
    // Save and close the document
    document.saveAndClose();
    
    // Do any additional processing or modifications to the new document as needed
  }
  
  function getNextOperaTagUpLink(folderId, excludeId) {
    var folder = DriveApp.getFolderById(folderId);
    var files = folder.getFiles();
    var latestFile = null;
    var latestFileTitle = '';
  
    while (files.hasNext()) {
      var file = files.next();
      var fileTitle = file.getName();
      var fileId = file.getId();
  
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
  
  function getMondayFollowingDate(date) {
    // Find the next Monday from the given date
    var day = date.getDay();
    var diff = (day === 0 ? 1 : 8) - day;
    var nextMonday = new Date(date.getTime() + diff * 24 * 60 * 60 * 1000);
  
    // Format the date as "YYYY-MM-DD"
    var formattedDate = Utilities.formatDate(nextMonday, Session.getScriptTimeZone(), "yyyy-MM-dd");
  
    return formattedDate;
  }
  
  function getMostRecentFileLink(folderId, excludeId, placeholderText) {
    var folder = DriveApp.getFolderById(folderId);
    var files = folder.getFiles();
    var newestFile = null;
    var newestFileDate = new Date(0);
  
    while (files.hasNext()) {
      var file = files.next();
      var fileDate = file.getLastUpdated();
      var fileId = file.getId();
      var fileName = file.getName();
  
      if (fileDate > newestFileDate && fileId !== excludeId && !fileName.includes("Template")) {
        newestFile = file;
        newestFileDate = fileDate;
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
      var file = DriveApp.getFileById(url.split('/')[5]);
      var fileName = file.getName();
      
      textElement.deleteText(startOffset, endOffset);
      textElement.insertText(startOffset, fileName).setLinkUrl(startOffset, startOffset + fileName.length - 1, url);
    }
  }
  