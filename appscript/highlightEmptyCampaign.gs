// This function will highlight any empty Campaign field
// in the Schedule AND Subscriptions sheet.
function highlightEmptyCampaign()
{
  let curSheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  if(curSheet.getName() === "Schedule" || curSheet.getName() === "Subscriptions")
  {
    let sheet = SpreadsheetApp.getActiveSheet();
    let range = SpreadsheetApp.getActive().getRange('A2:A');
    let rangeVals = range.getValues();
    let lastRow = range.getLastRow();

    // Loop through all the Campaigns in the spread sheet and check that they are not empty.
    for(j = 0; j < lastRow - 1; j++)
    {
      // Get the max number of columns for this sheet.
      let maxCol = sheet.getMaxColumns();

      // Clear any out of date notes.
      sheet.getRange(j+2, 1).clearNote();

      // Emptiness check.
      if(rangeVals[j][0] === "")
      {
        // j+2 is the row we are starting our range from.
        // 1 is the starting column of our range of what is going to be highlighted.
        // 1 is the number of rows, we are only interested in the row with the invalid data.
        // maxCol is the number of columns to include, we want to highlight the whole row.
        sheet.getRange(j+2, 1, 1, maxCol).setBackground("#cc4125");

        sheet.getRange(j+2, 1).setNote("Error: Campaign cannot be empty.");
      }
      // Check if highlighting can be cleared.
      checkNoErrors(sheet.getRange(j+2, 1, 1, maxCol));
    }
  }
}