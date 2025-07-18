function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
   .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

//~▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨ WITS 'Only change the info in this box ▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨

// To get the spreadsheet ID it's in the URL after /d/**SPREADSHEET_ID**/edit#
const spreadsheetId = '1TdUa0AX6iKz_WIg_1IEGH_3vdBmClmBIzOfY5a-XD68'; // Specificlyy References the entire spreadsheet
const ss = SpreadsheetApp.openById(spreadsheetId);
const referenceRow = 4700; // New Badge & Associate info gets added after this row

//~▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨

function checkDeviceStatus(deviceID) {
  const sheet = ss.getSheetByName("Log");
  const data = sheet.getDataRange().getValues();
  const headers = data[0];

  const deviceIndex = headers.indexOf("Device ID");
  const outTimestampIndex = headers.indexOf("Check-Out Timestamp");
  const nameIndex = headers.indexOf("Name");
  const checkedOutIndex = headers.indexOf("Checked Out?");
  const checkedInIndex = headers.indexOf("Checked In?");

  //for (let i = data.length - 1; i > 0; i--) { // Checks logs from ▽ to △
  for (let i = 1; i < data.length; i++) { // Checks logs from △ to ▽
    const row = data[i];
    if (row[deviceIndex] === deviceID) {
      const checkedOut = row[checkedOutIndex];
      const checkedIn = row[checkedInIndex];

      if (checkedOut === "Yes" && checkedIn === "Yes") {
        return {
          status: "checked_in", 
          message: "This device is in stock, check-out?"
          };
      } else if (checkedOut === "Yes" && (!checkedIn || checkedIn === "")) {
        const name = row[nameIndex] || "Last User of device";
        const outTime = row[outTimestampIndex]
          ? Utilities.formatDate(new Date(row[outTimestampIndex]), Session.getScriptTimeZone(), "M/d/yyyy h:mm:ss a")
          : "Unknown time";
        return {
          status: "checked_out", 
          message: `This device has been checked out by ${name} on ${outTime}.`};
      }
    }
  }
  
  return {
    status: "never_used", 
    message: "This device has not been logged yet. Check-out?"
    };
}

function checkIfBadgeExists(badgeID) {
  const sheet = ss.getSheetByName("BQ Associates");
  const data = sheet.getRange("C2:C" + sheet.getLastRow()).getValues().flat();
  return data.includes(badgeID);
}

function registerAndHandleNewBadge(badgeID, deviceID, action, userID, department) {
  const sheet = ss.getSheetByName("BQ Associates");

  //const referenceRow = 4900; // New Badge & Associate info gets added after this row (This was made into a gloabl variable)
  const insertRow = referenceRow + 1;
  const F_Associate1 = `=IFERROR(FILTER(AssociateWhEMID,E${insertRow}=AssociateUserID),"Incorrect User ID")`;
  const F_Associate2 = `=IFERROR(FILTER(AssociateName,E${insertRow}=AssociateUserID),"Incorrect User ID")`;
  const F_Associate3 = `=IFERROR(FILTER(AssociateDetails,E${insertRow}=AssociateUserID),"Incorrect User ID")`;
  sheet.insertRowBefore(insertRow);

  sheet.getRange(insertRow, 1).setFormula(F_Associate1)
  sheet.getRange(insertRow, 3).setValue(badgeID); // Assuming col C = Badge ID
  sheet.getRange(insertRow, 4).setFormula(F_Associate2)
  sheet.getRange(insertRow, 5).setValue(userID);  // Assuming col E = User ID
  sheet.getRange(insertRow, 6).setFormula(F_Associate3)

  // Proceed with check out
  handleDeviceAction(badgeID, deviceID, action, department);
}

function handleDeviceAction(badgeID, deviceID, action, department) {
  Logger.log(`handleDeviceAction called with badgeID: ${badgeID}, deviceID: ${deviceID}, action: ${action}`);
  const sheet = ss.getSheetByName("Log");
  const data = sheet.getDataRange().getValues();
  const now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "M/d/yyyy h:mm:ss a");
  
  // List of columns on the Log sheet
  const headers = data[0];
  const badgeIndex = headers.indexOf("Badge ID");
  const nameIndex = headers.indexOf("Name");
  const deviceIndex = headers.indexOf("Device ID");
  const departmentIndex = headers.indexOf("Department"); // New
  const checkedOutIndex = headers.indexOf("Checked Out?");
  const outTimestampIndex = headers.indexOf("Check-Out Timestamp");
  const checkedInIndex = headers.indexOf("Checked In?");
  const inTimestampIndex = headers.indexOf("Check-In Timestamp");
  const statusIndex = headers.indexOf("Status");
  const usageDurationIndex = headers.indexOf("Usage Duration"); // New
  const deviceStateIndex = headers.indexOf("Device State"); // NEW: Device State column

  //for (let i = data.length - 1; i > 0; i--) { // Checks logs from ▽ to △
  for (let i = 1; i < data.length; i++) { // Checks logs from △ to ▽
    const row = data[i];
    if (row[deviceIndex] === deviceID) {
      const checkedOut = row[checkedOutIndex];
      const checkedIn = row[checkedInIndex];

      if (action === "Check In" && checkedOut === "Yes" && !checkedIn) {
        sheet.getRange(i + 1, checkedInIndex + 1).setValue("Yes");

        const inTimestampCell = sheet.getRange(i + 1, inTimestampIndex + 1);
        inTimestampCell.setValue(now);
        inTimestampCell.setNumberFormat("M/d/yyyy h:mm:ss AM/PM");
        sheet.getRange(i + 1, statusIndex + 1).setValue("Complete");

        // Calculate duration
        const outTime = new Date(row[outTimestampIndex]);
        const inTime = new Date(now);
        const durationMs = inTime - outTime;
        const hours = Math.floor(durationMs / (1000 * 60 * 60));
        const minutes = Math.floor((durationMs % (1000 * 60 * 60)) / (1000 * 60));
        const formattedDuration = `${hours}h ${minutes}m`;
        sheet.getRange(i + 1, usageDurationIndex + 1).setValue(formattedDuration);

        return;
      }

      if (action === "Check Out" && checkedOut === "Yes" && !checkedIn) {
        Logger.log("Device already checked out, skipping...");
        // Device is already checked out and not yet returned
        return;
      }

      break; // We've found the latest record, move to create a new one if needed
    }
  }

  if (action === "Check Out") {
    Logger.log("Appending new row at the top for check out");

    // Always insert a new row at row 2 (just below the header)
    sheet.insertRows(2, 1); // Insert 1 row above row 2
    const targetRow = 2;

    Logger.log(`Logging new check out at row ${targetRow}`);

    sheet.getRange(targetRow, 1, 1, 1).setValue(badgeID); // Column A
    sheet.getRange(targetRow, 3, 1, 1).setValue(deviceID); // Column C
    sheet.getRange(targetRow, 4, 1, 1).setValue(department); // Column D
    sheet.getRange(targetRow, 5, 1, 1).setValue("Yes"); // Checked Out?
    sheet.getRange(targetRow, 6, 1, 1).setValue(now); // Check-Out Timestamp
    sheet.getRange(targetRow, 9, 1, 1).setValue("Out"); // Status
    sheet.getRange(targetRow, outTimestampIndex + 1).setNumberFormat("M/d/yyyy h:mm:ss AM/PM");

    // Set "Pending..." in Usage Duration column
    sheet.getRange(targetRow, usageDurationIndex + 1).setValue("Pending...");

    // NEW: Set default "Device State" to "In Use" for new devices
    if (deviceStateIndex >= 0) {
      sheet.getRange(targetRow, deviceStateIndex + 1).setValue("In Use");
      
      // Set up dropdown validation for Device State column
      const deviceStateOptions = ["On order", "In Stock", "In Transit", "In Use", "Consumed", "In Maintenance", "Retired", "Missing"];
      const deviceStateRange = sheet.getRange(targetRow, deviceStateIndex + 1);
      const deviceStateValidation = SpreadsheetApp.newDataValidation()
        .requireValueInList(deviceStateOptions, true)
        .build();
      deviceStateRange.setDataValidation(deviceStateValidation);
    }

    // Set data validation for the "Department" column
    const departmentOptions = ["Receiving", "Midile Mile", "Picking", "Putaway", "HDO", "Deluxing", "QC", "IC"]; // EDropdown Options
    const departmentRange = sheet.getRange(targetRow, departmentIndex + 1); // Target the "Department" Column | last row
    const departmentValidation = SpreadsheetApp.newDataValidation()
      .requireValueInList(departmentOptions, true)
      .build();
    departmentRange.setDataValidation(departmentValidation);

    // Apply the dynamic formula for new rows in column B
    const formula = `=IFERROR(IF(ISBLANK(A${targetRow}),"",VLOOKUP(A${targetRow},BQ_Badge,2,false)),"New Badge...")`;
    const Cell = sheet.getRange(targetRow, 2); // Column B
      
    Cell.setFormula(formula); // Column B
  }
}

function getUniqueDevicesFromLog() {
  const sheet = ss.getSheetByName("Log");
  const data = sheet.getRange("C2:C" + sheet.getLastRow()).getValues().flat();
  const devices = [...new Set(data.filter(id => id))].sort();
  return devices;
}

//-▢▢▢▢▢▢▢▢▢▢▢▢▢▢▢▢▢▢▢▢▢▢▢▢▢▢▢▢▢▢▢▢▢▢▢▢▢▢▢▢▢▢▢▢▢▢▢▢▢▢▢▢▢▢▢▢▢▢▢▢▢▢▢▢▢▢▢▢▢▢▢▢▢▢▢▢▢▢▢▢▢▢▢▢▢▢▢▢▢▢▢▢▢▢▢▢▢▢▢▢▢▢▢▢▢▢▢▢

function getNexusData() {
  try {
    var sheet = ss.getSheetByName("BQ Devices");
    
    // Get all data from the sheet
    var data = sheet.getDataRange().getValues();
    
    // Return the data (first row typically contains headers)
    return data;
  } catch (error) {
    console.error('Error fetching Nexus data:', error);
    return [];
  }
}

// NEW: Helper function to set up Device State column validation for existing rows
function setupDeviceStateValidation() {
  const sheet = ss.getSheetByName("Log");
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const deviceStateIndex = headers.indexOf("Device State");
  
  if (deviceStateIndex >= 0) {
    const lastRow = sheet.getLastRow();
    const deviceStateOptions = ["On order", "In Stock", "In Transit", "In Use", "Consumed", "In Maintenance", "Retired", "Missing"];
    
    // Apply validation to all rows with data (excluding header)
    if (lastRow > 1) {
      const deviceStateRange = sheet.getRange(2, deviceStateIndex + 1, lastRow - 1, 1);
      const deviceStateValidation = SpreadsheetApp.newDataValidation()
        .requireValueInList(deviceStateOptions, true)
        .build();
      deviceStateRange.setDataValidation(deviceStateValidation);
    }
  }
}