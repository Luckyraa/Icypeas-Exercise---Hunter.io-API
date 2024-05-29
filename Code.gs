// Function to run when the Google Sheet is opened
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Hunter.io')
    .addItem('Set API Key', 'showApiKeyDialog')
    .addItem('Find Emails', 'processEmails')
    .addToUi();
}

// Function to show a dialog for entering the Hunter.io API key
function showApiKeyDialog() {
  var htmlOutput = HtmlService.createHtmlOutputFromFile('ApiKeyDialog')
      .setWidth(400)
      .setHeight(300);
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Enter Hunter.io API Key');
}

// Function to save the entered API key
function saveApiKey(apiKey) {
  PropertiesService.getUserProperties().setProperty('HUNTER_API_KEY', apiKey);
  SpreadsheetApp.getUi().alert('API Key saved successfully!');
}

// Function to call the Hunter API and retrieve an email address
function FindEmail(firstName, lastName, company) {
  Logger.log('FindEmail called with: ' + firstName + ' ' + lastName + ' ' + company);
  var apiKey = PropertiesService.getUserProperties().getProperty('HUNTER_API_KEY');
  if (!apiKey) {
    Logger.log('API key not found');
    return 'Error: API key not found. Please set the API key using the Hunter.io menu.';
  }

  if (!firstName || !lastName || !company) {
    Logger.log('Missing input data');
    return 'Error: Missing input data';
  }

  var url = 'https://api.hunter.io/v2/email-finder?domain=' + encodeURIComponent(company) + 
            '&first_name=' + encodeURIComponent(firstName) + 
            '&last_name=' + encodeURIComponent(lastName) + 
            '&api_key=' + apiKey;

  Logger.log('Request URL: ' + url);
  
  try {
    var response = UrlFetchApp.fetch(url);
    var result = JSON.parse(response.getContentText());
    Logger.log('API response: ' + response.getContentText());
    
    if (result.data && result.data.email) {
      return result.data.email;
    } else if (result.errors) {
      Logger.log('API error: ' + result.errors[0].details);
      return 'Error: ' + result.errors[0].details;
    } else {
      return 'No email found';
    }
  } catch (e) {
    Logger.log('Exception during API call: ' + e.message);
    return 'Error: Exception during API call: ' + e.message;
  }
}

// Function to process each row in the sheet and find emails
function processEmails() {
  var ui = SpreadsheetApp.getUi();
  ui.alert('Email search process started. This may take a few moments.');

  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var dataRange = sheet.getDataRange();
  var data = dataRange.getValues();
  var emailsFound = 0;
  var errors = 0;

  for (var i = 1; i < data.length; i++) {
    var firstName = data[i][0];
    var lastName = data[i][1];
    var company = data[i][2];
    
    if (firstName && lastName && company) {
      var email = FindEmail(firstName, lastName, company);
      sheet.getRange(i + 1, 5).setValue(email); // Assuming emails go to the E column
      if (email.startsWith('Error:')) {
        errors++;
      } else {
        emailsFound++;
      }
    } else {
      Logger.log('Skipping row ' + (i + 1) + ' due to missing data');
      sheet.getRange(i + 1, 5).setValue('Error: Missing data');
      errors++;
    }
  }

  ui.alert('Email search completed.\nEmails found: ' + emailsFound + '\nErrors: ' + errors);
}
