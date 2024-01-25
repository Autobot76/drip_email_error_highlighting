// The purpose of this is to contain all the calls required
// to check schedule in one function instead of many.
function checkSchedule()
{
  highlightInvalidTime();

  // These functions apply to both the Schedule and Subscriptions Sheets.
  highlightEmptyCampaign();

  highlightInvalidTemplate();
}

// This function will highlight a row in red and leave a note
// if it has an invalid time format.
function highlightInvalidTime()
{
  // This is used to verify that it is the Schedule sheet.
  let curSheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

  // If it is the Schedule sheet verify the time.
  if(curSheet.getName() === "Schedule")
  {
    let sheet = SpreadsheetApp.getActiveSheet();
    let range = SpreadsheetApp.getActive().getRange('C2:C');
    let rangeVals = range.getValues();
    let lastRow = range.getLastRow();

    // Get the max number of columns for this sheet.
    let maxCol = sheet.getMaxColumns();

    for(j = 0; j < lastRow - 1; j++)
    {
      // Clear any out of date notes.
      sheet.getRange(j+2, 3).clearNote();

      // Clear out-of-date highlighting.
      // sheet.getRange(j+2, 1, 1, maxCol).setBackground("#ffffff");

      // Emptiness check.
      if(rangeVals[j][0] === "")
      {
        // j+2 is the row we are starting our range from. j starts at zero but to getRange() it starts at 1.
        // 3 is the starting column of our range. It is 3 because the fourth column is the time.
        // sheet.getRange(j+2, 3).setBackground("#cc4125");
        // sheet.getRange(j+2, 3).setNote("Warning: Time cannot be empty.");

        // As discussed, if the time is left empty a default time of 10AM shall be used.
        sheet.getRange(j+2, 3).setValue("10:00:00 AM");
      }
      // Take advantage of the fact that if a user enters the time in our accepted format
      // Google Sheets automatically will make that into a Date object. So, if it is just
      // a String the user has made a mistake.
      else if(typeof rangeVals[j][0] === "string")
      {
        sheet.getRange(j+2, 1, 1, maxCol).setBackground("#cc4125");
        sheet.getRange(j+2, 3).setNote("Warning: Time is not in a valid format.");
      }
      // A number is not accepted it should be a Date object.
      else if(typeof rangeVals[j][0] === "number")
      {
        sheet.getRange(j+2, 1, 1, maxCol).setBackground("#cc4125");
        sheet.getRange(j+2, 3).setNote("Warning: Time is not in a valid format.");
      }
    }
  }
}

// This function will check for invalid Templates in this Schedule sheet.
// This should not be confused with highlightInvalidTemplateName().
// There is no need to check if a Template already exists because
// there will already be a data validation rule in place.
function highlightInvalidTemplate()
{
  let curSheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

  if(curSheet.getName() === "Schedule")
  {
    let sheet = SpreadsheetApp.getActiveSheet();
    let range = SpreadsheetApp.getActive().getRange('B2:B');
    let rangeVals = range.getValues();
    let lastRow = range.getLastRow();

    // Adapted from: https://yagisanatode.com/2017/12/13/google-apps-script-iterating-through-ranges-in-sheets-the-right-and-wrong-way/
    // Loop through all the Template names in the sheet and check that they are not empty.
    for(j = 0; j < lastRow - 1; j++)
    {
      // Get the max number of columns for this sheet.
      let maxCol = sheet.getMaxColumns();

      // Clear any out-of-date notes.
      sheet.getRange(j+2, 2).clearNote();

      // Emptiness check.
      if(rangeVals[j][0] === "")
      {
        sheet.getRange(j+2, 1, 1, maxCol).setBackground("#cc4125");

        // Add a note telling the user that the Template cannot be empty.
        sheet.getRange(j+2, 2).setNote("Error: Template cannot be empty.");
      }
    }
  }
}