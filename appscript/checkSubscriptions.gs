// The purpose of this is to contain all the calls required
// to check subscriptions in one function instead of many.
function checkSubscriptions()
{
  // Shared with checkSchedule
  highlightEmptyCampaign();

  highlightInvalidLastEmailDay();
  highlightInvalidSubscribeDate();

  // Shared with checkSchedule
  dupCampaigns();

  highlightEmptyUser();
  highlightInvalidLastEmailDate();
}

// This function will highlight all invalid LastEmailDay rows in the Subscription Sheet.
// Currently, anything that is not a number is considered invalid.
// The offending row will be highlighted in red and a note will be left in the offending
// cell.
function highlightInvalidLastEmailDay()
{
  // Get the current sheet. This is used to verify that it is the Subscriptions sheet.
  let curSheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

  // Currently if it is the Subscriptions sheet then verify the LastEmailDay
  if(curSheet.getName() === "Subscriptions")
  {
    let sheet = SpreadsheetApp.getActiveSheet();
    let range = SpreadsheetApp.getActive().getRange('F2:F');
    let rangeVals = range.getValues();
    let lastRow = range.getLastRow();

    // Adapted from: https://yagisanatode.com/2017/12/13/google-apps-script-iterating-through-ranges-in-sheets-the-right-and-wrong-way/
    // Loop through all the LastEmailDay in the sheet and check that they are a number.
    for(j = 0; j < lastRow - 1; j++)
    {
      // Get the max number of columns for this sheet.
      let maxCol = sheet.getMaxColumns();

      // Clear any out-of-date notes.
      // 6 because LastEmailDay is in column F which is 6.
      sheet.getRange(j+2, 6).clearNote();

      // If it is not a number highlight the row red and leave a note.
      if(isNaN(rangeVals[j][0]))
      {
        sheet.getRange(j+2, 1, 1, maxCol).setBackground("#cc4125");
        sheet.getRange(j+2, 6).setNote("Error: LastEmailDay must be a number.");
      }
      // Check if highlighting can be cleared.
      checkNoErrors(sheet.getRange(j+2, 1, 1, maxCol));
    }
  }
}

// This function will check and highlight the SubscribeDate that is empty
// or not a date. It will leave a note and highlight the cell in red.
function highlightInvalidSubscribeDate()
{
  // Verify that it is the Subscriptions sheet.
  let curSheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

  // If it is the Subscriptions sheet verify the SubscribeDate.
  if(curSheet.getName() === "Subscriptions")
  {
    let sheet = SpreadsheetApp.getActiveSheet();
    let range = SpreadsheetApp.getActive().getRange('D2:D');
    let rangeVals = range.getValues();
    let lastRow = range.getLastRow();

    for(j = 0; j < lastRow - 1; j++)
    {
      // Clear any out of date notes.
      sheet.getRange(j+2, 4).clearNote();

      // Emptiness check.
      if(rangeVals[j][0] === "")
      {
        // 4 is the column for SubscribeDate.
        sheet.getRange(j+2, 4).setBackground("#cc4125");
        sheet.getRange(j+2, 4).setNote("Error: SubscribeDate cannot be empty.");
      }
      // Take advantage of the fact that if a user enters the Date in our accepted format
      // Google Sheets automatically will make that into a Date object. So, if it is a
      // String or a number the user has made a mistake.
      else if(typeof rangeVals[j][0] === "string" || typeof rangeVals[j][0] === "number")
      {
        sheet.getRange(j+2, 4).setBackground("#cc4125");
        sheet.getRange(j+2, 4).setNote("Error: SubscribeDate is not in a valid format.");
      }
      // Check if highlighting can be cleared.
      checkNoErrors(sheet.getRange(j+2, 1, 1, sheet.getMaxColumns()));
    }
  }
}

// This function will highlight empty User fields in Subscriptions
// and leave a note.
function highlightEmptyUser()
{
  let curSheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  if(curSheet.getName() === "Subscriptions")
  {
    let sheet = SpreadsheetApp.getActiveSheet();
    let range = SpreadsheetApp.getActive().getRange('B2:B');
    let rangeVals = range.getValues();
    let lastRow = range.getLastRow();

    // Loop through all the Users in the sheet and check that they are not empty.
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
        sheet.getRange(j+2, 2).setNote("Warning: User cannot be empty.");
      }

      // Check if highlighting can be cleared.
      checkNoErrors(sheet.getRange(j+2, 1, 1, maxCol));
    }
  }
}

// NOTE: The date in Sheets is in MM/DD/YYYY

// This function will highlight all invalid LastEmailDate cells in the Subscription Sheet.
// Currently, anything that is not a date is considered invalid.
// The offending cell will be highlighted in yellow and a note will be left in the offending
// cell.

// NOTE: Empty is not considered invalid here and instead the program will automatically
// set this field's value to that of SubscribeDate.
function highlightInvalidLastEmailDate()
{
  // Verify that it is the Subscriptions sheet.
  let curSheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

  // Verify the LastEmailDate
  if(curSheet.getName() === "Subscriptions")
  {
    let sheet = SpreadsheetApp.getActiveSheet();
    let range = SpreadsheetApp.getActive().getRange('E2:E');
    let rangeVals = range.getValues();
    let lastRow = range.getLastRow();

    // Adapted from: https://yagisanatode.com/2017/12/13/google-apps-script-iterating-through-ranges-in-sheets-the-right-and-wrong-way/
    // Loop through all the LastEmailDay in the sheet and check that they are a number.
    for(j = 0; j < lastRow - 1; j++)
    {
      // Clear any out-of-date notes.
      // 5 - LastEmailDate is in column E which is 5.
      sheet.getRange(j+2, 5).clearNote();

      // If LastEmailDate is empty fill it with whatever is in SubscribeDate (SD).
      // SD cannot be empty.
      // Some checks should first be done to ensure that whatever is in SD is valid
      // before being copied here.
      if(rangeVals[j][0] === '')
      {
        // Check SubscribeDate is not empty and is a Date object.
        if(sheet.getRange(j+2, 4).getValue() !== "" && typeof sheet.getRange(j+2, 4).getValue() !== "string" && typeof sheet.getRange(j+2, 4).getValue() !== "number")
        {
          // Get the "Index" for SubscribeDate and set the value of LastEmailDate to that.
          sheet.getRange(j+2, 5).setValue(sheet.getRange(j+2, 4).getValue());
        }
        else
        {
          sheet.getRange(j+2, 5).setBackground("#ffea00");
          sheet.getRange(j+2, 5).setNote("Error: Cannot copy from SubscribeDate since it is invalid.");
        }
      }
      // If it is not a Date highlight the row red and leave a note.
      else if(typeof rangeVals[j][0] === "string")
      {
        sheet.getRange(j+2, 5).setBackground("#cc4125");
        sheet.getRange(j+2, 5).setNote("Error: LastEmailDate must be a date.");
      }
      // Check if highlighting can be cleared.
      checkNoErrors(sheet.getRange(j+2, 1, 1, sheet.getMaxColumns()));
    }
  }
}