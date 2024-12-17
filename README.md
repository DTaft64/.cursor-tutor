[READ ME for Google Sheets Attendance Tracker.md](https://github.com/user-attachments/files/17823421/READ.ME.for.Google.Sheets.Attendance.Tracker.md)

# Google Sheets Attendance Tracker

## Overview
This Google Apps Script processes form responses to track unique attendance for Friday events. It automatically filters, counts, and formats attendance data from a form response sheet into a clean, readable output.

## Features
- Filters for Friday attendance only
- Removes duplicate entries for the same person on the same day
- Counts total attendance per person
- Sorts by attendance frequency
- Formatted output with consistent styling

## Sheet Requirements
The script requires two sheets in your Google Spreadsheet:
1. **Form Responses 1**: Contains the raw form submission data with:
   - Column A: Timestamp
   - Column B: Last Name
   - Column C: First Name

2. **Attendance Output DO NOT TOUCH**: Destination sheet for processed data

## Setup The Sheet
1. Install the apps script code.
   - Open your Google Sheet
   - Go to Extensions > Apps Script
   - Copy the entire code into the script editor
   - Save the script
   - Run the `generateUniqueNames()` function
2. Configure the spreadsheet.
   - Title the Google forum sheet "Form Responses 1"
   - Title the output sheet "Attendance Output DO NOT TOUCH"
3. Make a "run" button.
   - Open the output sheet "Attendance Output DO NOT TOUCH"
   - Make a run butten by making a drawing and asigning a script to it. Insert> drawing> (make a drawing)> done> click the three dots in the upper right> inserted drawing and select "Assign script" and type the script name "generateUniqueNames"


## Output Format
The script generates:
- Centered, 12-point font text
- Bold headers
- Three columns:
  - Unique First Names
  - Unique Last Names
  - Number of times attended
- Description of the script's functionality (in column D)

## Functions

### `initializeCache()`
Initializes cached references to spreadsheet objects to improve performance.

### `generateUniqueNames()`
Main function that processes the form data and generates the output.

### `writeOutputData(outputData)`
Handles formatting and writing the processed data to the output sheet.

### `updateDescription()`
Adds explanatory text about the script's functionality.

## Performance Optimizations
- Sheet reference caching
- Batch operations for reading/writing
- Efficient data processing using reduce()
- Minimized API calls

## Maintenance
- Sheet names must match exactly: "Form Responses 1" and "Attendance Output DO NOT TOUCH"
- The script assumes timestamps are in Column A, last names in Column B, and first names in Column C
- Any changes to the form structure will require corresponding updates to the script

## Support
For issues or questions, please contact me at dtaft on discord
