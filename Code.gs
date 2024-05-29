function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Hunter.io')
    .addItem('Open Settings', 'showSidebar')
    .addToUi();
}

function showSidebar() {
  var html = HtmlService.createHtmlOutputFromFile('Sidebar')
      .setTitle('Hunter.io Settings')
      .setWidth(300);
  SpreadsheetApp.getUi().showSidebar(html);
}

function saveApiKey(apiKey) {
  PropertiesService.getUserProperties().setProperty('HUNTER_API_KEY', apiKey);
  SpreadsheetApp.getUi().alert('API Key saved successfully!');
}



/**
 * Custom function to find an email using Hunter.io API with support for company names.
 * @param {string} firstName - The first name of the person.
 * @param {string} lastName - The last name of the person.
 * @param {string} company - The company name or company email.
 * @return {string} - The found email address or an error message.
 * @customfunction
 */
function FindEmail(firstName, lastName, company) {
  var apiKey = PropertiesService.getUserProperties().getProperty('HUNTER_API_KEY');
  if (!apiKey) {
    return 'Error: API key not found. Please set the API key using the Hunter.io menu.';
  }

  if (!firstName || !lastName || !company) {
    return 'Error: Missing input data';
  }

  var domain = isValidDomain(company) ? company : getDomainFromCompanyName(company);
  if (!domain) {
    return 'Error: Invalid company name or domain.';
  }

  var url = 'https://api.hunter.io/v2/email-finder?domain=' + encodeURIComponent(domain) + 
            '&first_name=' + encodeURIComponent(firstName) + 
            '&last_name=' + encodeURIComponent(lastName) + 
            '&api_key=' + apiKey;

  try {
    var response = UrlFetchApp.fetch(url);
    var result = JSON.parse(response.getContentText());

    if (result.data && result.data.email) {
      return result.data.email;
    } else if (result.errors) {
      return 'Error: ' + result.errors[0].details;
    } else {
      return 'No email found';
    }
  } catch (e) {
    return 'Error: Exception during API call: ' + e.message;
  }
}

/**
 * Helper function to validate if a string is a valid domain.
 * @param {string} domain - The domain to validate.
 * @return {boolean} - True if the domain is valid, false otherwise.
 */
function isValidDomain(domain) {
  var domainPattern = /^[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$/;
  return domainPattern.test(domain);
}

/**
 * Helper function to get the domain from a company name using Hunter.io API.
 * @param {string} companyName - The company name to get the domain for.
 * @return {string|null} - The domain if found, null otherwise.
 */
function getDomainFromCompanyName(companyName) {
  var apiKey = PropertiesService.getUserProperties().getProperty('HUNTER_API_KEY');
  var url = 'https://api.hunter.io/v2/domain-search?company=' + encodeURIComponent(companyName) + '&api_key=' + apiKey;

  try {
    var response = UrlFetchApp.fetch(url);
    var result = JSON.parse(response.getContentText());

    if (result.data && result.data.domain) {
      return result.data.domain;
    } else {
      return null;
    }
  } catch (e) {
    return null;
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


function testFindEmail() {
  var email = FindEmail('Florent', 'Gaujal', 'nsigma.fr');
  Logger.log('Test Email: ' + email);
}


