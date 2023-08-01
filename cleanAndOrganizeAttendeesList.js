// Purpose: to process a list of attendees copied from an email's required: and optional: list, replace specific email addresses with associated names, remove additional email addresses, unneccessary text & duplicates, then alphabetize by first name 

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
    'raw0037@uah.edu': 'Rachel Wyatt',
    'ksv0003@uah.edu': 'Katrina Virts',
    'jeanne.leroux@nsstc.uah.edu': "Jeanne' le Roux",
    'jeanne.leroux@uah.edu': "Jeanne' le Roux",
    'jr0020@uah.edu': "Jeanne' le Roux",
    'al0001@uah.edu': 'Anita Leroy',
    'olofsson76@gmail.com': 'Pontus Olofsson',
    'Charley@wayforagers.org': 'Charley Haley',
    'nasaldh@gmail.com': 'Larry Hill'    
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
