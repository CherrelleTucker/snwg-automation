// Helper function to get today's date
function getToday() {
    return new Date();
  }
  
  // Helper function to calculate the next most recent due date
  function getNextDueDate() {
    const startDate = new Date("December 4, 2023 12:00:00");
    const today = getToday();
  
    // Calculate the difference in days from the start date
    const daysDiff = Math.floor((today - startDate) / (24 * 60 * 60 * 1000));
  
    // Calculate the number of bi-weeks since the start date
    const biWeeksSinceStart = Math.floor(daysDiff / 14);
  
    // Calculate the next due date
    const nextDueDate = new Date(startDate);
    nextDueDate.setDate(startDate.getDate() + (biWeeksSinceStart + 1) * 14);
    nextDueDate.setHours(12, 0, 0, 0);
  
    return nextDueDate;
  }
  
  // Helper function to format a date in the standard format
  function formatDate(date) {
    const options = { weekday: 'long', year: 'numeric', month: 'long', day: 'numeric' };
    return date.toLocaleDateString('en-US', options);
  }
  
  // Helper function to show a dialog with the next due date
  function showNextDueDateDialog() {
    const nextDueDate = getNextDueDate();
    const today = getToday();
  
    const daysDifference = Math.floor((nextDueDate - today) / (24 * 60 * 60 * 1000));
  
    if (daysDifference === 0) {
      // If the closest due date is today, show "TODAY AT NOON" with red color, bold font weight, and larger font size; Verified Working
      return `<p style="font-size: 24px; font-weight: bold; color: red;">TODAY AT NOON</p>`;
    } else if (daysDifference >= 8 && daysDifference <= 14) {
      // If today is within the week AFTER the closest Due Date, show the date in black, the phrase "Due in 2 Weeks: " and no font weight
      return `<p style="font-size: 18px; color: black;">Due in 2 Weeks: ${formatDate(nextDueDate)}</p>`;
    } else if (daysDifference >= 1 && daysDifference <= 7) {
      // If today is within the week BEFORE the closest due date, show the date in blue, "Next Week:" and no font weight; Verified Working
      return `<p style="font-size: 18px; color: blue;">Next Week: ${formatDate(nextDueDate)}</p>`;
    }
  }
  
  
  // Primary function to execute all helper actions
  function showBlurbsDueDialog() {
    const message = showNextDueDateDialog();
  
    // Create a custom dialog box with formatted HTML content
    const htmlOutput = HtmlService.createHtmlOutput(message)
      .setTitle('Blurbs Due')
      .setWidth(300)
      .setHeight(100);
  
    // Show the dialog box
    DocumentApp.getUi().showModalDialog(htmlOutput, 'Blurbs Due');
  }
  
  // Function to add a custom menu to the document toolbar
  function onOpen() {
    DocumentApp.getUi()
      .createMenu('When are Blurbs Due?')
      .addItem('Check Due Date', 'showBlurbsDueDialog')
      .addToUi();
  }
  
  // Function to simulate blurbs being due today (for testing purposes)
  function testBlurbsDueToday() {
    const originalGetToday = getToday;
    getToday = () => new Date('December 18, 2023 10:00:00'); // Set a specific date and time for testing
    showBlurbsDueDialog();
    getToday = originalGetToday; // Restore the original getToday function
  }
  
  // Function to simulate blurbs being due the week before (for testing purposes)
  function testDateTheWeekAfterBlurbsAreDue() {
    const originalGetToday = getToday;
    getToday = () => new Date('December 6, 2023 10:00:00'); // Set a specific date and time for testing (2 days after)
    showBlurbsDueDialog();
    getToday = originalGetToday; // Restore the original getToday function
  }
  
  // Function to simulate blurbs being due within the week after (for testing purposes)
  function testDateOneWeekBeforeBlurbsAreDue() {
    const originalGetToday = getToday;
    getToday = () => new Date('December 15, 2023 10:00:00'); // Set a specific date and time for testing (3 days before)
    showBlurbsDueDialog();
    getToday = originalGetToday; // Restore the original getToday function
  }
  
  // Run the onOpen function when the document is opened
  onOpen();
  