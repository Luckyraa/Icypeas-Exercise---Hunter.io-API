// Function to display the add-on option in the spreadsheet
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('My Hunter.io add-on')
    .addItem('Open Sidebar', 'showSidebar')
    .addToUi();
}


// Function to display the Sidebar
function showSidebar() {
  var html = HtmlService.createHtmlOutputFromFile('Sidebar')
      .setTitle('Hunter.io Sidebar')
      .setWidth(300);
  SpreadsheetApp.getUi().showSidebar(html);
}


// Function to save API key
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
    return 'Error: API key not found.';
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
    } else if (result.errors && result.errors.length > 0) {
      return 'Error: ' + result.errors[0].details;
    } else {
      return 'No email found';
    }
  } catch (e) {
    return 'Error: ' + e.message;
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
    } else if (result.errors && result.errors.length > 0) {
      return null;
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


// Function to test each case, especially the error cases 
function testFindEmail() {
  var originalApiKey = PropertiesService.getUserProperties().getProperty('HUNTER_API_KEY');
  
  // Set the API key for all tests except the one testing missing API key
  if (!originalApiKey) {
    PropertiesService.getUserProperties().setProperty('HUNTER_API_KEY', 'your_actual_api_key_here');
  }

  var testCases = [
    {
      description: "Valid inputs with a company domain",
      firstName: "Sacha",
      lastName: "Edery",
      company: "nsigma.fr",
      expectedResult: "sacha.edery@nsigma.fr"
    },
    {
      description: "Valid inputs with a company name",
      firstName: "Sacha",
      lastName: "Edery",
      company: "Nsigma",
      expectedResult: "sacha.edery@nsigma.fr"
    },
    {
      description: "Invalid domain",
      firstName: "Alice",
      lastName: "Johnson",
      company: "invalid_domain",
      expectedResult: "Error: Invalid company name or domain."
    },
    {
      description: "Invalid company name",
      firstName: "Bob",
      lastName: "Brown",
      company: "NonExistentCompany",
      expectedResult: "Error: Invalid company name or domain."
    },
    {
      description: "Missing first name",
      firstName: "",
      lastName: "White",
      company: "example.com",
      expectedResult: "Error: Missing input data"
    },
    {
      description: "Missing last name",
      firstName: "Charlie",
      lastName: "",
      company: "example.com",
      expectedResult: "Error: Missing input data"
    },
    {
      description: "Missing company",
      firstName: "David",
      lastName: "Green",
      company: "",
      expectedResult: "Error: Missing input data"
    }
  ];

  // Run tests with API key set
  var results = testCases.map(function(testCase) {
    var result = FindEmail(testCase.firstName, testCase.lastName, testCase.company);
    return {
      description: testCase.description,
      result: result,
      passed: result === testCase.expectedResult || result.startsWith("Expected email address")
    };
  });

  // Test case for API key not set
  PropertiesService.getUserProperties().deleteProperty('HUNTER_API_KEY');
  var noApiKeyTest = {
    description: "API key not set",
    firstName: "Eve",
    lastName: "Black",
    company: "example.com",
    expectedResult: "Error: API key not found."
  };
  var noApiKeyResult = FindEmail(noApiKeyTest.firstName, noApiKeyTest.lastName, noApiKeyTest.company);
  results.push({
    description: noApiKeyTest.description,
    result: noApiKeyResult,
    passed: noApiKeyResult === noApiKeyTest.expectedResult
  });

  // Restore the original API key if it was set
  if (originalApiKey) {
    PropertiesService.getUserProperties().setProperty('HUNTER_API_KEY', originalApiKey);
  }

  // Log results
  results.forEach(function(testResult) {
    Logger.log("Test: " + testResult.description);
    Logger.log("Result: " + testResult.result);
    Logger.log("Passed: " + testResult.passed);
    Logger.log("--------------------");
  });
}
