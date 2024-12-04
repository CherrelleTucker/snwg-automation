function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Contact Tools')
    .addItem('Show Consolidated Contact Info (Exit Views and Filters Before Selecting)', 'showContactInfo')
    .addItem("Create targeted email list", "openDialog")
    .addToUi();
}

function showContactInfo() {
  const ui = SpreadsheetApp.getUi();
  const range = SpreadsheetApp.getActiveSpreadsheet().getActiveRange();

  // Ensure exactly one cell is selected in column C
  if (range.getNumColumns() === 1 && range.getNumRows() === 1 && range.getColumn() === 3) { // Column C is index 2
    const contactName = range.getCell(1, 1).getValue(); // Get value from column C

    const contactInfo = getContactInfo(contactName);
    const htmlOutput = HtmlService.createHtmlOutput(`
      <style>
        body {
          font-family: Arial, sans-serif;
          font-size: 14px;
          line-height: 1.5;
          padding: 0px;
          background-color: #f9f9f9;
        }
        h1 {
          font-size: 20px;
        }
        ul {
          list-style-type: square;
          padding-left: 20px;
        }
        li {
          margin-bottom: 8px;
        }
        .role-group {
          margin-bottom: 24px; /* Increased spacing */
        }
        .important {
          font-weight: bold;
        }
        .role-item:nth-child(odd) {
          background-color: #f0f0f0; /* Light gray for odd items */
        }
        .role-item:nth-child(even) {
          background-color: #ffffff; /* White for even items */
        }
        .section {
          margin-bottom: 24px; /* Add spacing between sections */
        }
      </style>
      <body>
        <h1>Contact Information for ${contactName}</h1>
        ${contactInfo}
      </body>
    `)
    .setWidth(800)
    .setHeight(700);

    ui.showModalDialog(htmlOutput, `Contact Information for ${contactName}`);
  } else {
    ui.alert('Please select a single cell in column C.');
  }
}

function getContactInfo(contactName) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('All_Contacts'); // Replace with your actual sheet name
  if (!sheet) {
    SpreadsheetApp.getUi().alert('The sheet named "All_Contacts" was not found. Please check the sheet name.');
    return 'Sheet not found.';
  }
  const data = sheet.getDataRange().getValues();
  const headers = data[1]; // Headers are in the second row (index 1)
  const result = {};
  const roleYearMap = new Map(); // To keep track of roles with combined years

  // Initialize a result object for the specific columns we are interested in
  const columnsToInclude = [0, 1, 5, 6, 7, 8, 12]; // Removed 'Assessment Solution' column and retained other necessary columns

  columnsToInclude.forEach(index => {
    result[headers[index]] = new Map(); // Use a Map to maintain the original value and its lowercased version
  });

  // Iterate through the rows, starting from the third row to skip headers
  for (let i = 2; i < data.length; i++) { // Data starts from row 3 (index 2)
    if (data[i][2] === contactName) { // Column C is index 2
      for (const index of columnsToInclude) {
        const header = headers[index];
        let value = data[i][index];
        if (value !== undefined && value !== null) {
          // Ensure value is a string for trimming and comparison
          value = String(value).trim();
          const normalizedValue = value.toLowerCase(); // Normalize for comparison
          if (!result[header].has(normalizedValue)) { // Check if the lowercased value is already stored
            result[header].set(normalizedValue, value); // Store both the lowercased key and original value
          }
        }
      }
      // Collect roles separately and combine roles with multiple years
      const role = data[i][3]; // Column D is index 3
      const year = data[i][10]; // Column K is index 10
      const assessmentSolution = data[i][1]; // Column B is index 1
      if (role && year && assessmentSolution) {
        const roleKey = `${role} in ${assessmentSolution}`;
        if (!roleYearMap.has(roleKey)) {
          roleYearMap.set(roleKey, new Set());
        }
        roleYearMap.get(roleKey).add(year);
      }
    }
  }

  // Format the roles into strings with combined years
  let roles = [];
  roleYearMap.forEach((years, roleKey) => {
    const combinedYears = Array.from(years).sort().join(' and ');
    roles.push(`<div class="role-item">- ${roleKey} in ${combinedYears}</div>`);
  });

  // Format the result object into a string for display with extra spacing
  let formattedResult = '';

  // Adding roles section if available with banding for visibility
  if (roles.length > 0) {
    formattedResult += `<div class="section"><strong>Roles:</strong><br>`;
    formattedResult += roles.join('') + '</div>';
  }

  // Adding other contact information
  for (const [key, values] of Object.entries(result)) {
    if (values.size > 0 && key !== 'Assessment Solution') { // Skip 'Assessment Solution' key
      const uniqueValues = Array.from(values.values()); // Retrieve the original values from the Map
      formattedResult += `<div class="section"><strong>${key}:</strong> ${uniqueValues.join(', ')}</div>`;
    }
  }

  return formattedResult || 'No contact found.';
}
