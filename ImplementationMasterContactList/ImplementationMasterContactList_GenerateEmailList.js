// Function to open the dialog box in Google Sheets
function openDialog() {
  const htmlOutput = HtmlService.createHtmlOutputFromFile("EmailDialog")
    .setWidth(800)
    .setHeight(700);
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, "Generate Email List");
}

// Function to generate the email list based on user-selected criteria
function generateEmailList(data) {
  const sheet = getSheetByName("All_Contacts");
  const rows = getSheetData(sheet);
  const logMessages = [];

  const indices = getColumnIndices();
  logMessages.push(`Fixed Column Indices: ${JSON.stringify(indices)}`);

  // Log the data received
  logMessages.push(`Received Data: ${JSON.stringify(data)}`);

  const filteredEmails = filterRows(rows, data, indices, logMessages)
    .map(row => row[indices.emailIndex]);

  const uniqueEmails = getUniqueSortedEmails(filteredEmails);
  logMessages.push(`Unique Emails: ${uniqueEmails}`);

  Logger.log(logMessages.join("\n")); // Log detailed messages for better debugging

  return {
    uniqueEmails: uniqueEmails.length > 0 ? uniqueEmails : ["No Contacts found matching the requested criteria"],
    logMessages
  };
}

// Function to create a Google Doc with the email list based on the criteria
function openEmailListDoc(emailList, criteria) {
  // Split emailList string by the separator ("; ") to create an array
  const emailArray = emailList.split("; ").filter(email => email.trim() !== "");

  const title = createDocTitle(criteria);
  const doc = DocumentApp.create(title);
  const body = doc.getBody();

  addDocContent(body, emailArray, criteria);
  return { title, url: doc.getUrl() };
}

// Helper function to get sheet by name
function getSheetByName(name) {
  return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
}

// Helper function to get sheet data
function getSheetData(sheet) {
  return sheet.getDataRange().getValues();
}

// Helper function to define column indices based on headers
function getColumnIndices() {
  return {
    roleIndex: 3,
    yearIndex: 10,
    contactLevelIndex: 4,
    cycleIndex: 0,
    emailIndex: 5,
    thematicAreaStartIndex: 16, // Column Q (index starts at 0)
    thematicAreaEndIndex: 24    // Column Y
  };
}

// Helper function to filter rows based on the criteria
function filterRows(rows, data, indices, logMessages) {
  return rows.slice(1).filter(row => {
    logMessages.push(`Row data: ${JSON.stringify(row)}`);

    // Make all matching case-insensitive
    const roleMatch = !data.role.length || data.role.some(role => role.toLowerCase() === row[indices.roleIndex].toLowerCase());
    const yearMatch = !data.year.length || data.year.includes(String(row[indices.yearIndex]));
    const contactLevelMatch = !data.contactLevel.length || data.contactLevel.some(level => level.toLowerCase() === row[indices.contactLevelIndex].toLowerCase());
    const cycleMatch = !data.cycle.length || data.cycle.some(cycle => row[indices.cycleIndex].toLowerCase().includes(cycle.toLowerCase()));

    // Modified logic for SNWG Project to handle "SNWG MO"
    const projectMatch = !data.project.length || data.project.some(project => {
      if (project.toLowerCase() === "snwg mo") {
        // Match any of "SNWG Assessment", "SNWG SEP", or "SNWG Implementation"
        return ["snwg assessment", "snwg sep", "snwg implementation"].includes(row[indices.cycleIndex].toLowerCase());
      } else {
        return row[indices.cycleIndex].toLowerCase() === project.toLowerCase();
      }
    });

    // Check Thematic Areas (Columns Q to Y)
    const thematicAreaMatch = !data.thematicAreas.length || data.thematicAreas.every(area => {
      const thematicIndex = getThematicAreaIndex(area);
      return row[indices.thematicAreaStartIndex + thematicIndex] === true;
    });

    logMessages.push(`roleMatch: ${roleMatch}, yearMatch: ${yearMatch}, contactLevelMatch: ${contactLevelMatch}, cycleMatch: ${cycleMatch}, projectMatch: ${projectMatch}, thematicAreaMatch: ${thematicAreaMatch}`);

    return roleMatch && yearMatch && contactLevelMatch && cycleMatch && projectMatch && thematicAreaMatch;
  });
}

// Helper function to get thematic area index for corresponding slice index (Q-Y)
function getThematicAreaIndex(thematicArea) {
  const thematicAreas = [
    'Atmospheric Composition',
    'Carbon Cycle & Ecosystems',
    'Disaster Response',
    'Earth Surface & Interior',
    'Land Cover / Land Use Change',
    'Ocean & Cryosphere',
    'Water & Energy Cycle',
    'Weather & Atmospheric Dynamics',
    'Other / Infrastructure'
  ];
  return thematicAreas.indexOf(thematicArea);
}

// Helper function to remove duplicate emails and sort them alphabetically
function getUniqueSortedEmails(emails) {
  return [...new Set(emails)].sort();
}

// Helper function to create the document title
function createDocTitle(criteria) {
  const role = criteria.role ? criteria.role.join(", ") : "Any Role";
  const year = criteria.year ? criteria.year.join(", ") : "Any Year";
  const contactLevel = criteria.contactLevel ? criteria.contactLevel.join(", ") : "Any Contact Level";
  const cycle = criteria.cycle ? criteria.cycle.join(", ") : "Any Cycle";
  const project = criteria.project ? criteria.project.join(", ") : "Any Project";
  const thematicAreas = criteria.thematicAreas ? criteria.thematicAreas.join(", ") : "Any Thematic Area";

  let title = `${role}_${year}_${contactLevel}_${cycle}_${project}_${thematicAreas}_Contact List`;
  return title.replace(/_{2,}/g, "_").replace(/_$/, "");
}

// Helper function to add content to the document
function addDocContent(body, emailList, criteria) {
  body.appendParagraph("Email Contact List").setHeading(DocumentApp.ParagraphHeading.HEADING1);
  body.appendParagraph(`Generated on: ${new Date().toLocaleString()}`);
  body.appendParagraph("\nCriteria:");
  body.appendParagraph(`Role: ${criteria.role ? criteria.role.join(", ") : "Any Role"}`);
  body.appendParagraph(`Year: ${criteria.year ? criteria.year.join(", ") : "Any Year"}`);
  body.appendParagraph(`Contact Level: ${criteria.contactLevel ? criteria.contactLevel.join(", ") : "Any Contact Level"}`);
  body.appendParagraph(`Cycle: ${criteria.cycle ? criteria.cycle.join(", ") : "Any Cycle"}`);
  body.appendParagraph(`Project: ${criteria.project ? criteria.project.join(", ") : "Any Project"}`);
  body.appendParagraph(`Thematic Areas: ${criteria.thematicAreas ? criteria.thematicAreas.join(", ") : "Any Thematic Area"}`);
  body.appendParagraph("\nEmail List:");
  body.appendParagraph(emailList.join("; "));
}

// Test function for logging and debugging purposes
function testGenerateEmailList() {
  // Simulate data collection from the HTML dialog box
  const testCriteria = {
    contactLevel: ["2. Secondary/Optional"],
    thematicAreas: ["Water & Energy Cycle"]
  };

  // Log collected data to ensure it matches what is passed from the HTML dialog
  Logger.log(`Test Criteria: ${JSON.stringify(testCriteria)}`);

  // Call the generateEmailList function with the test criteria
  const result = generateEmailList(testCriteria);
  Logger.log(`Unique Emails: ${result.uniqueEmails}`);
  result.logMessages.forEach(msg => Logger.log(msg));

  // Call the openEmailListDoc function to simulate document creation
  const docResult = openEmailListDoc(result.uniqueEmails, testCriteria);
  Logger.log(`Google Doc Created: Title - ${docResult.title}, URL - ${docResult.url}`);
}

// Test function to validate Google Doc creation
function testCreateGoogleDoc() {
  Logger.log("Running test for Google Doc creation...");
  const testEmailList = "test1@example.com; test2@example.com; test3@example.com";
  
  // Call the openEmailListDoc function directly
  const docResult = openEmailListDoc(testEmailList, {
    role: ["Test Role"],
    year: ["2022"],
    contactLevel: ["2. Secondary/Optional"],
    cycle: "C2",
    project: "C2 - Commercial Solution",
    thematicAreas: ["Disaster Response"]
  });

  // Log the document result
  Logger.log(`Google Doc Created: Title - ${docResult.title}, URL - ${docResult.url}`);
}
