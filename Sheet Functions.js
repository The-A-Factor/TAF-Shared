//~▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨- Configuration -▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨

const MAX_ROWS = 5244; // Maximum number of data rows to keep
const HEADER_ROW = 0; // Row 1 contains headers
const DATA_START_ROW = 2; // Data starts from row 2

//~▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨

/*
 Automatic Trigger
 Run this once to install the trigger that monitors for changes
*/
function installTrigger() {
  const triggers = ScriptApp.getScriptTriggers();
  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === 'manageDeviceLogRows') {
      ScriptApp.deleteTrigger(trigger);
    }
  });
  
  // Create new trigger
  ScriptApp.newTrigger('manageDeviceLogRows')
    .timeBased()
    .everyMinutes(5) // Check every minute (you can adjust this)
    .create();
    
  console.log('Automatic Trigger installed successfully - will check every 5 Minutes');
}

/*
 Remove the automatic trigger if you want to stop it
*/
function removeTrigger() {
  const triggers = ScriptApp.getScriptTriggers();
  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === 'manageDeviceLogRows') {
      ScriptApp.deleteTrigger(trigger);
      console.log('Trigger removed');
    }
  });
}

function manageDeviceLogRows() {
  const sheet = ss.getSheetByName("Log");
  const lastRow = sheet.getLastRow();
  
  // Calculate how many data rows we have (excluding header)
  const dataRows = lastRow - HEADER_ROW;
  console.log(`Current data rows: ${dataRows}, Max allowed: ${MAX_ROWS}`);
  
  // If we exceed the maximum, delete rows from the bottom
  if (dataRows > MAX_ROWS) {
    const rowsToDelete = dataRows - MAX_ROWS;
    const startDeleteRow = lastRow - rowsToDelete + 1;
    console.log(`Deleting ${rowsToDelete} rows starting from row ${startDeleteRow}`);
    
    // Delete the excess rows
    sheet.deleteRows(startDeleteRow, rowsToDelete);
    console.log(`Successfully deleted ${rowsToDelete} old entries`);
  }
}