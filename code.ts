// Google Apps Script to handle RSVP and Wishes form submissions for the Wedding Invitation
// Save this code in your Google Apps Script editor (Extensions > Apps Script)

function doGet(e) {
  return handleResponse(e);
}

function doPost(e) {
  return handleResponse(e);
}

function handleResponse(e) {
  try {
    const params = e.parameter;
    const formName = params.formName;
    
    if (!formName) {
      return createJsonResponse({ error: 'Missing formName parameter' }, 400);
    }
    
    if (formName === 'rsvp') {
      return handleRsvp(params);
    } else if (formName === 'wish') {
      return handleWish(params);
    } else {
      return createJsonResponse({ error: 'Invalid formName' }, 400);
    }
    
  } catch (error) {
    return createJsonResponse({ error: error.toString() }, 500);
  }
}

function handleRsvp(params) {
  const sheetName = 'RSVP';
  const headers = ['Timestamp', 'Name', 'Guests', 'Dietary Notes'];
  const rowData = [
    new Date(),
    params['Name'] || '',
    params['Guests'] || '',
    params['Dietary Notes'] || ''
  ];
  
  return appendToSheet(sheetName, headers, rowData);
}

function handleWish(params) {
  const sheetName = 'Wishes';
  const headers = ['Timestamp', 'Name', 'Message'];
  const rowData = [
    new Date(),
    params['Name'] || '',
    params['Message'] || ''
  ];
  
  return appendToSheet(sheetName, headers, rowData);
}

function appendToSheet(sheetName, headers, rowData) {
  // Gets the spreadsheet the script is attached to
  const doc = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = doc.getSheetByName(sheetName);
  
  // Auto-generate sheet and headers if it doesn't exist
  if (!sheet) {
    sheet = doc.insertSheet(sheetName);
    
    // Add headers
    sheet.appendRow(headers);
    
    // Style headers to make them stand out
    const headerRange = sheet.getRange(1, 1, 1, headers.length);
    headerRange.setFontWeight('bold');
    headerRange.setBackground('#F3F4F6');
    headerRange.setFontFamily('Montserrat');
    
    // Freeze the top row so headers stay visible when scrolling
    sheet.setFrozenRows(1);
    
    // Adjust column widths automatically
    for (let i = 1; i <= headers.length; i++) {
      sheet.autoResizeColumn(i);
    }
  }
  
  // Append the new form data
  sheet.appendRow(rowData);
  
  return createJsonResponse({ 
    success: true, 
    message: `Successfully added to ${sheetName}` 
  });
}

function createJsonResponse(data, statusCode = 200) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * Run this function manually once from the Apps Script editor 
 * to instantly set up both sheets and their headers.
 * (Optional - the sheets will also be created automatically upon first submission)
 */
function setupSheets() {
  const doc = SpreadsheetApp.getActiveSpreadsheet();
  
  // Setup RSVP sheet
  if (!doc.getSheetByName('RSVP')) {
    const rsvpSheet = doc.insertSheet('RSVP');
    const rsvpHeaders = ['Timestamp', 'Name', 'Guests', 'Dietary Notes'];
    rsvpSheet.appendRow(rsvpHeaders);
    rsvpSheet.getRange('A1:D1').setFontWeight('bold').setBackground('#F3F4F6');
    rsvpSheet.setFrozenRows(1);
  }
  
  // Setup Wishes sheet
  if (!doc.getSheetByName('Wishes')) {
    const wishSheet = doc.insertSheet('Wishes');
    const wishHeaders = ['Timestamp', 'Name', 'Message'];
    wishSheet.appendRow(wishHeaders);
    wishSheet.getRange('A1:C1').setFontWeight('bold').setBackground('#F3F4F6');
    wishSheet.setFrozenRows(1);
  }
  
  // Clean up default 'Sheet1' if it's empty
  const sheet1 = doc.getSheetByName('Sheet1');
  if (sheet1 && sheet1.getLastRow() === 0) {
    doc.deleteSheet(sheet1);
  }
}
