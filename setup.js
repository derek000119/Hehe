// Step 1: Create the Google Form

function createFeedbackForm() {
  var form = FormApp.create('EchoHR: Employee Feedback');
  
  form.setDescription('Your feedback is valuable to us. Please share your thoughts anonymously.');
  
  form.addScaleItem()
    .setTitle('How satisfied are you with your work environment?')
    .setBounds(1, 5)
    .setLabels('Not satisfied', 'Very satisfied');
    
  form.addParagraphTextItem()
    .setTitle('What challenges are you currently facing at work?');
    
  form.addParagraphTextItem()
    .setTitle('Do you have any suggestions for improvement?');
    
  form.addMultipleChoiceItem()
    .setTitle('Which area does your feedback primarily relate to?')
    .setChoiceValues(['Work-life balance', 'Management', 'Resources', 'Team dynamics', 'Career growth', 'Other']);
    
  form.setConfirmationMessage('Thank you for your feedback! Your input helps us improve our workplace.');
  
  Logger.log('Form created with ID: ' + form.getId());
  return form;
}

// Step 2: Set up the Google Sheet

function setupFeedbackSheet() {
  var sheet = SpreadsheetApp.create('EchoHR Feedback Responses');
  var form = FormApp.openById("1CgxghpH6COd37Tisv6rf63LJ23uJNm_lKCitXT1RyD0"); // Replace with actual form ID
  form.setDestination(FormApp.DestinationType.SPREADSHEET, sheet.getId());
  
  Logger.log('Sheet created with ID: ' + sheet.getId());
  return sheet;
}

// Step 3: Create a trigger for weekly reminders

function createWeeklyTrigger() {
  ScriptApp.newTrigger('sendWeeklyReminder')
    .timeBased()
    .onWeekDay(ScriptApp.WeekDay.FRIDAY)
    .atHour(10)
    .create();
}

// Step 4: Function to send weekly reminders

function sendWeeklyReminder() {
  var form = FormApp.openById('1CgxghpH6COd37Tisv6rf63LJ23uJNm_lKCitXT1RyD0'); // Replace with actual form ID
  var formUrl = form.getPublishedUrl();
  
  var recipients = getEmployeeEmails(); // Implement this function to fetch employee emails
  
  for (var i = 0; i < recipients.length; i++) {
    MailApp.sendEmail({
      to: recipients[i],
      subject: 'Weekly Feedback Reminder - EchoHR',
      body: "Don't forget to submit your weekly feedback. Your input is valuable to us!\n\n" +
            "Click here to submit your feedback: " + formUrl
    });
  }
}

// Helper function to get employee emails (implement according to your system)
function getEmployeeEmails() {
  // This is a placeholder. In a real scenario, you'd fetch emails from your employee database
  return ['johnson38831688@gmail.com', 'leecheng_jun@yahoo.com', 'mrtkz0529@gmail.com'];
}

// Run this function to set up the entire system
function myFunction() {
  var form = createFeedbackForm();
  var sheet = setupFeedbackSheet();
  createWeeklyTrigger();
  
  Logger.log('EchoHR system set up successfully!');
  Logger.log('Form ID: ' + form.getId());
  Logger.log('Sheet ID: ' + sheet.getId());  //1KWkpd4ZV-QzagcsM8JjKzT3eIUAdPq7pGO3KAAwX1xA
}