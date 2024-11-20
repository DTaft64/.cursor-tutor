// Script to Generate Unique Names and Attendance Counts
//
// Cache frequently accessed sheets
const CACHE = {
  sheets: null,
  formSheet: null,
  outputSheet: null
};

function initializeCache() {
  if (!CACHE.sheets) {
    CACHE.sheets = SpreadsheetApp.getActiveSpreadsheet();
    CACHE.formSheet = CACHE.sheets.getSheetByName("Form Responses 1");
    CACHE.outputSheet = CACHE.sheets.getSheetByName("Attendance Output DO NOT TOUCH");
  }
}

function generateUniqueNames() {
  initializeCache();
  
  // Get all data in one batch operation
  const [header, ...rows] = CACHE.formSheet.getDataRange().getValues();
  
  // Process data more efficiently
  const fridayAttendees = rows.reduce((acc, [timestamp, lastName, firstName]) => {
    const date = new Date(timestamp);
    if (date.getDay() === 5) { // Friday check
      const dateStr = date.toDateString();
      const key = `${dateStr}-${firstName}-${lastName}`;
      
      if (!acc.seen.has(key)) {
        acc.seen.add(key);
        const nameKey = `${firstName} ${lastName}`;
        acc.counts[nameKey] = (acc.counts[nameKey] || 0) + 1;
        acc.names[nameKey] = { firstName, lastName };
      }
    }
    return acc;
  }, { seen: new Set(), counts: {}, names: {} });

  // Create output data in one pass
  const outputData = Object.entries(fridayAttendees.counts)
    .map(([nameKey, count]) => [
      fridayAttendees.names[nameKey].firstName,
      fridayAttendees.names[nameKey].lastName,
      count
    ])
    .sort((a, b) => b[2] - a[2]);

  // Batch write operations
  writeOutputData(outputData);
  updateDescription();
}

function writeOutputData(outputData) {
  const headers = [["Unique First Names", "Unique Last Names", "# of times attended"]];
  
  // Clear and write all data in minimal operations
  CACHE.outputSheet.clear();
  
  // Write headers with formatting
  const headerRange = CACHE.outputSheet.getRange(1, 1, 1, 3);
  headerRange
    .setValues(headers)
    .setFontWeight("bold")
    .setVerticalAlignment("middle")
    .setFontSize(12)
    .setHorizontalAlignment("center");
  
  // Write data if exists
  if (outputData.length > 0) {
    CACHE.outputSheet.getRange(2, 1, outputData.length, 3)
      .setValues(outputData)
      .setFontSize(12)
      .setHorizontalAlignment("center");
  }
}

function updateDescription() {
  const description = [
    "This script processes data from a Google Sheet ('Form Responses 1') to extract and analyze attendance records.",
    "It identifies unique first and last name pairs from timestamps, filters entries to include only those recorded on Fridays,",
    "and excludes duplicate entries for the same name on the same Friday.",
    "The script then counts how many times each unique name appears across all Fridays",
    "and sorts the results in descending order of attendance count.",
    "Finally, it writes the unique first names, last names, and their corresponding counts into a second sheet ('Attendance Output DO NOT TOUCH'),",
    "with headers formatted in bold."
  ];

  const descriptionRange = CACHE.outputSheet.getRange(4, 4, description.length, 1);
  descriptionRange
    .setValues(description.map(line => [line]))
    .setWrapStrategy(SpreadsheetApp.WrapStrategy.OVERFLOW)
    .setFontSize(12);
}
// END OF SCRIPT