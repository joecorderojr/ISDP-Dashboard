function doGet() {
  getLoginUser();

  // Index is the name of your HTML file
  var tmp = HtmlService.createTemplateFromFile("index");

  document_list = loadDocumentData();
  tmp.list = document_list;

  initializedDocumentVariables(); 
  getDataPrivacyTrainingData2025();
  getDataPrivacyTrainingData2026();
  getCyberSecurityTrainingData2025();
  getCyberSecurityTrainingData2026();
  getCyberSecurityTrainingData2024();
  getPhishingResult();
  getISO27001Data();
  getNPCChecklist();
  getTPSData();
  getVASummaryData();
  getSOC2TYPE2ChartData();
  getISOComplianceStatus();

  //var tempDashboardTableList = getDashboardData();
  //tmp.dashboardTableData = tempDashboardTableList;
  //Logger.log(tmp.dashboardTableData)

  return tmp.evaluate().setSandboxMode(HtmlService.SandboxMode.IFRAME);
}

function include(filename) {
  try {
    // 1. Create a template object from the file
    const template = HtmlService.createTemplateFromFile(filename);
    
    // 2. Evaluate the template (process scriptlets) and get the content
    return template.evaluate().getContent();
  } catch (e) {
    Logger.log(`Error including file: ${filename}. Details: ${e.stack}`);
    return `<div style="border: 2px solid red; padding: 10px; background-color: #381e1eff;">
              <strong>Error including ${filename}:</strong>
              <pre>${e.stack}</pre>
            </div>`;
  }
}

function include_picker(filename,datePickerId) {
  try {
    const template = HtmlService.createTemplateFromFile(filename);
    template.datePickerId = datePickerId;
    template.rowId = "picker_" + datePickerId;
    return template.evaluate().getContent();
  } catch (e) {
    Logger.log(`Error including file: ${filename}. Details: ${e.stack}`);
    return `<div style="border: 2px solid red; padding: 10px; background-color: #381e1eff;">
              <strong>Error including ${filename}:</strong>
              <pre>${e.stack}</pre>
            </div>`;
  }
}

function include_object(filename, obj) {
  try {
    const template = HtmlService.createTemplateFromFile(filename);
    template.obj = obj;
    return template.evaluate().getContent();
  } catch (e) {
    Logger.log(`Error including file: ${filename}. Details: ${e.stack}`);
    return `<div style="border: 2px solid red; padding: 10px; background-color: #381e1eff;">
              <strong>Error including ${filename}:</strong>
              <pre>${e.stack}</pre>
            </div>`;
  }
}

function formatDateTimeJS(dateObj) {
  const dateTimeObject = new Date(dateObj); // Gets the current date and time
  
  // Gets the date components based on the local time zone of the server
  const year = dateTimeObject.getFullYear();
  const month = dateTimeObject.getMonth() + 1; // getMonth() returns 0-11, so add 1
  const day = dateTimeObject.getDate();
  
  // Pads single-digit month/day with a leading zero (e.g., 5 -> 05)
  const paddedMonth = String(month).padStart(2, '0');
  const paddedDay = String(day).padStart(2, '0');

  if(isNaN(paddedMonth)) 
    return "";

  const dateOnlyString = `${paddedMonth}/${paddedDay}/${year}`; 

  return dateOnlyString;
}

function duplicateSpreadsheet() {
  const ARCHIVE_FOLDER_NAME = "Archive";
  
  try {
    // 1. Get the current active spreadsheet object and its details
    const ss = SpreadsheetApp.openByUrl(url);
    const spreadsheetName = ss.getName();
    const spreadsheetId = ss.getId();
    
    // 2. Generate a timestamp for the new file name
    const timestamp = new Date().toLocaleDateString('en-US', { 
      year: 'numeric', 
      month: '2-digit', 
      day: '2-digit', 
      hour: '2-digit', 
      minute: '2-digit'
    }).replace(/,/g, ''); 

    const newFileName = `${spreadsheetName} - COPY ${timestamp}`;

    // --- NEW LOGIC START ---
    
    // 3. Find or Create the Archive Folder
    let archiveFolder;
    // Search for the folder named 'Archive' in the root of Drive
    const folders = DriveApp.getFoldersByName(ARCHIVE_FOLDER_NAME); 
    
    if (folders.hasNext()) {
      // If found, use the first matching folder
      archiveFolder = folders.next();
      Logger.log("Found existing folder: %s", ARCHIVE_FOLDER_NAME);
    } else {
      // If not found, create the folder in the root of Drive
      archiveFolder = DriveApp.createFolder(ARCHIVE_FOLDER_NAME);
      Logger.log("Created new folder: %s", ARCHIVE_FOLDER_NAME);
    }
    
    // 4. Get the File object of the current spreadsheet
    const originalFile = DriveApp.getFileById(spreadsheetId);

    // 5. Create the duplicate file INSIDE the Archive folder
    // The makeCopy(name, destinationFolder) signature is used here.
    const newFile = originalFile.makeCopy(newFileName, archiveFolder);
    
    // --- NEW LOGIC END ---

    //Logger.log("Original File Name: %s", spreadsheetName);
    //Logger.log("New File Name: %s", newFileName);
    //Logger.log("Successfully duplicated file to folder: %s (ID: %s)", ARCHIVE_FOLDER_NAME, newFile.getId());

  } catch (e) {
    Logger.log("Failed to duplicate file: %s", e.toString());
  }
}

function mapAllHeaders(headers) {
  const headerMap = {};
  
  headers.forEach((header, index) => {
    const standardizedHeader = header.trim().toUpperCase();
    
    // Initialize the array if the header is encountered for the first time
    if (!headerMap[standardizedHeader]) {
      headerMap[standardizedHeader] = [];
    }
    
    // Push the current index
    headerMap[standardizedHeader].push(index);
  });
  
  return headerMap;
}

function filterEmptyRows(data) {
  if (!Array.isArray(data) || data.length === 0) {
    return [];
  }

  return data.filter(row => {
    // We use .some() to check if at least ONE cell in the row is NOT an empty string ("").
    // .some() returns true as soon as it finds a match, making it very efficient.
    return row.some(cellValue => {
      // Check if the cell value is not empty, null, or undefined, 
      // and if it's a string, check that it's not just whitespace.
      if (typeof cellValue === 'string') {
        return cellValue.trim() !== "";
      }
      // For numbers, dates, or boolean values, simply check for non-falsy values.
      return cellValue; 
    });
  });
}

function sortArrayByProperty(data, propertyName, ascending = true) {
  if (!Array.isArray(data) || data.length < 2) {
    return data; // Cannot sort empty or single-element arrays
  }

  // Create a shallow copy to prevent modifying the original array (best practice)
  const sortedData = [...data];

  sortedData.sort((a, b) => {
    // Safely retrieve property values and ensure they are treated as strings for comparison
    const valA = String(a[propertyName] || '').toLowerCase();
    const valB = String(b[propertyName] || '').toLowerCase();
    let comparison = 0;

    if (valA < valB) {
      comparison = -1;
    } else if (valA > valB) {
      comparison = 1;
    } else {
      comparison = 0;
    }

    // Apply ascending/descending logic
    return ascending ? comparison : comparison * -1;
  });

  return sortedData;
}

function updateSheetRow(data, row) {
  const ss = SpreadsheetApp.openByUrl(url);
  const sheet = ss.getSheetByName("Updated Masterlist");

  const newValues = sheet.getRange(row, 1, 1, 7).getValues()[0];

  newValues[0] = data.unit,
  newValues[1] = data.controlNo,
  newValues[2] = data.title,
  newValues[3] = data.type,
  newValues[4] = data.description,
  newValues[5] = data.preparedBy, 
  newValues[6] = data.effectiveDate,
  newValues[7] = data.ammendmentDate,  
  newValues[8] = data.latestVersion,  
  newValues[9] = data.documentLink,  
  newValues[10] = data.retentionPeriod,
  newValues[11] = data.lastReviewedDate,
  newValues[13] = data.remarks

  sheet.getRange(row, 1, 1, 14).setValues([newValues]);
}

function parseDate(dateStr) {
  if (!dateStr) return '';

  // Handles MM/DD/YYYY
  if (dateStr.includes('/')) {
    const [m, d, y] = dateStr.split('/');
    return new Date(y, m - 1, d);
  }

  // Handles YYYY-MM-DD
  return new Date(dateStr);
}

function addNewRow(data) {
  const ss = SpreadsheetApp.openByUrl(url);
  const sheet = ss.getSheetByName("Updated Masterlist");

  const TEMPLATE_ROW = 5;
  const newRow = sheet.getLastRow() + 1;

  const effectiveDate = parseDate(data.effectiveDate);
  const ammendmentDate = parseDate(data.ammendmentDate);
  const lastReviewedDate = parseDate(data.lastReviewedDate);

  sheet.getRange(newRow, 1, 1, 12).setValues([[
    data.unit,
    data.controlNo,
    data.title,
    data.type,
    data.description,
    data.preparedBy, 
    effectiveDate,
    ammendmentDate,  
    data.latestVersion,  
    data.documentLink,  
    data.retentionPeriod,
    lastReviewedDate,
  ]]);

  // Format date values
  sheet.getRange(newRow, 7)
    .setNumberFormat('mmmm d, yyyy');
  sheet.getRange(newRow, 8)
    .setNumberFormat('mmmm d, yyyy');
  sheet.getRange(newRow, 12)
    .setNumberFormat('mmmm d, yyyy');

  sheet.getRange(TEMPLATE_ROW, 13)
    .copyTo(sheet.getRange(newRow, 13), { contentsOnly: false });

  sheet.getRange(newRow, 14).setValue(data.remarks);
}