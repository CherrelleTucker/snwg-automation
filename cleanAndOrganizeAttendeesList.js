// Purpose: 
// Clean and organize the attendees' section in a Google Document. It restructures names from "Last Name, First Name Middle Initial" to "First Name Last Name" format, removes duplicate entries, and eliminates any email addresses from the list. Additionally, it maps specific email addresses to their corresponding names based on a provided mapping and sorts the names alphabetically. The script aims to enhance the presentation and readability of the attendees' section by ensuring it contains unique and properly formatted names.

// To Note:
// This script is developed for use as either as a Google Apps Script container script or as a Google Apps Script library script. 
// Google Apps Script container script: a script that is bound to a specific file, such as a Google Sheets, Google Docs, or Google Forms file. This container script acts as the file's custom script, allowing users to extend the functionality of the file by adding custom functions, triggers, and menu items to enhance its behavior and automation.

// Google Apps Script library script: a self-contained script that contains reusable functions and can be attached to multiple projects or files. By attaching the library script to different projects, developers can access and use its functions across various files, enabling code sharing and improving code maintenance and version control.

// To Use As container script:
// Open your Google Document: Open the Google Document where you want to extract action items and populate them in an action tracking table.
// Open the Script Editor: Click on "Extensions" in the top menu and then select "Apps Script." This will open the Google Apps Script editor in a new tab.
// Copy the Script: Copy the entire script provided above and paste it into the Google Apps Script editor.
// Save the Script: Click on the floppy disk icon or press "Ctrl + S" (Windows) or "Cmd + S" (Mac) to save the script.
// Execute the function organizeAttendees
// Enable Permissions:The script will request permission to access your Google Document. Click "Continue" and grant the necessary permissions.

// Please ensure that you are familiar with the Google Apps Script environment and have appropriate access to edit the Google Document before running the script. Also, review and customize the script as per your specific requirements before using it.

//////////////////////////////////////////////////

// To use this script as a library script:
// 1. Obtain the script ID of the inDocActionItems library script.
//  script ID: 1kBbrOJCXewvSixfq1yR8d-lEtgDG5yzD9-pqPeuC9ugLka7gQULwkBH_ <-- verify current library script id by checking in Project settings (gear icon).
// 2. Open the container document where you want to use this script.
// 3. Click on the "Project settings" gear icon in the script editor.
// 4. In the "Script ID" field, replace the existing script ID with the script ID of the inDocActionItems library script.
// 5. Click "Save" to update the script ID.

//////////////////////////////////

// Helper function to reorganize names from "Last Name, First Name Middle Initial" to "First Name Last Name"
function reorganizeNames(names) {
  return names.map(function (name) {
    var nameParts = name.split('(');
    var nameOnly = nameParts[0].trim();
    var nameComponents = nameOnly.split(',').map(function (part) {
      return part.trim();
    });

    if (nameComponents.length >= 2) {
      var firstName = nameComponents[1];
      var lastName = nameComponents[0];

      // If a middle initial is present, remove it
      var firstNameParts = firstName.split(' ');
      if (firstNameParts.length === 2) {
        firstName = firstNameParts[0];
      }

      return firstName + ' ' + lastName;
    }

    return nameOnly;
  });
}

// Helper function to remove email addresses from the attendees' section
function removeEmails(attendees) {
  return attendees.map(function (name) {
    return name.replace(/<[^>]*>|(\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,}\b)[;,]*/g, '').trim();
  });
}

// Helper function to map email addresses to names
function mapEmailsToNames(attendees) {
  var emailToName = {
    'name@doman.com': 'First Last',
    'name@doman.com': 'First Last',
    'name@doman.com': 'First Last'
  };

  var mappedNames = {}; // Object to store the mapped names
  var mappedEmails = {}; // Object to store the mapped email addresses
  var reorganizedNames = reorganizeNames(attendees); // Reorganize the names to "First Name Last Name" format

  return reorganizedNames.map(function (name) {
    var emailMatch = name.match(/[\w.-]+@[\w.-]+\.\w+/); // Extract the email address from the name

    if (emailMatch && emailMatch[0] in emailToName) {
      var mappedName = emailToName[emailMatch[0]];

      // Check if the mapped name has already been used in the mapping or reorganized list
      var isMappedNameUsed = mappedNames.hasOwnProperty(mappedName);
      var isReorganizedNameUsed = reorganizedNames.indexOf(mappedName) !== -1;

      // Check if the email address has already been mapped to a name
      var isEmailMapped = mappedEmails.hasOwnProperty(emailMatch[0]);

      if (isMappedNameUsed || isReorganizedNameUsed || isEmailMapped) {
        // If the name or email address has already been used, return nothing (remove the duplicate entry)
        return '';
      } else {
        // If it hasn't been used, add it to the mappedNames and mappedEmails objects and return the mapped name
        mappedNames[mappedName] = true;
        mappedEmails[emailMatch[0]] = true;
        return mappedName;
      }
    } else {
      return name;
    }
  }).filter(Boolean); // Remove any empty strings from the resulting array
}

// Helper function to remove duplicates from an array
function removeDuplicates(arr) {
  var uniqueNames = new Set();
  var resultArr = [];
  for (var i = 0; i < arr.length; i++) {
    var name = arr[i].toLowerCase(); // Convert to lowercase
    if (!uniqueNames.has(name)) {
      uniqueNames.add(name);
      resultArr.push(arr[i]);
    }
  }
  return resultArr;
}

// Primary function to create a cleaned attendees list
function organizeAttendees() {
  // Step 1: Get the active Google Document and its body
  var document = DocumentApp.getActiveDocument();
  var body = document.getBody();
  
  // Step 2: Find the section with the phrase "Attendees: "
  var attendeesPhrase = 'Attendees: ';
  var attendeesText = body.findText(attendeesPhrase);

  // Step 3: Check if the attendees section is found
  if (attendeesText) {
    var attendeesElement = attendeesText.getElement();
    var attendeesString = attendeesElement.asText().getText().substring(attendeesPhrase.length);

    // Step 4: Check if there are any attendees
    if (!attendeesString.trim()) {
      Logger.log('No attendees found.');
      return; // Stop further processing
    }

    // Step 5: Split the attendees section into an array of names
    var attendeesArray = attendeesString.split(';').map(function (name) {
      return name.trim();
    });

    // Step 6: Log the found attendees (optional)
    Logger.log('Attendees found:');
    Logger.log(attendeesArray);

    // Step 7: Remove duplicates from the attendees list
    var uniqueAttendees = removeDuplicates(attendeesArray);

    // Step 8: Reorganize names to "First Name Last Name" format
    var reorganizedAttendees = reorganizeNames(uniqueAttendees);

    // Step 9: Check if there are any email addresses in the attendees list
    var hasEmailAddresses = reorganizedAttendees.some(function (name) {
      return name.match(/[\w.-]+@[\w.-]+\.\w+/);
    });

    // Step 10: Process attendees only if there are email addresses
    if (hasEmailAddresses) {
      // Step 11: Map email addresses to names using the provided mapping
      var attendeesWithNames = mapEmailsToNames(reorganizedAttendees);

      // Step 12: Sort the names alphabetically
      attendeesWithNames.sort();

      // Step 13: Check if any changes were made to the attendee list
      var commaSeparatedAttendees = attendeesWithNames.join(', ');
      if (commaSeparatedAttendees === attendeesString) {
        Logger.log('Attendees are already sorted, formatted, and duplicates are removed.');
        return; // Stop further processing
      }

      // Step 14: Update the attendees section with the cleaned and organized list
      attendeesElement.asText().setText(attendeesPhrase + commaSeparatedAttendees);

      // Step 15: Log the updated attendees (optional)
      Logger.log('Attendees sorted, formatted, and duplicates removed:');
      Logger.log(commaSeparatedAttendees);
    } else {
      Logger.log('No email addresses found. Skipping further processing.');
    }
  } else {
    Logger.log('Attendees section not found.');
  }
}
