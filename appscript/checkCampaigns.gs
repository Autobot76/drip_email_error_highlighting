// The purpose of this is to contain all the calls required
// to check campaigns in one function instead of many.
function checkCampaigns()
{
  highlightEmptyCampaignIds();
  testDuplicatesIds();
  highlightInvalidSender();
}

// This function will highlight a row in red if a Campaign ID is not present
// and leave a note stating that the Campaign ID is empty.
function highlightEmptyCampaignIds()
{
  // Get the current sheet. This is used to verify that it is the Campaigns sheet.
  let curSheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

  // If it is the Campaigns sheet then check if campaign ID is empty.
  if(curSheet.getName() === "Campaigns")
  {
    let sheet = SpreadsheetApp.getActiveSheet();
    let range = SpreadsheetApp.getActive().getRange('A2:A');
    let rangeVals = range.getValues();
    let lastRow = range.getLastRow();

    // Loop through all the campaign IDs in the spread sheet and check that they are not empty.
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

        // Add a note telling the user that the Campaign ID cannot be empty.
        sheet.getRange(j+2, 1).setNote("Error: Campaign ID cannot be empty.");
      }
      else
      {
        // If a proper campaign ID has been entered we no longer need to highlight this row.
        sheet.getRange(j+2, 1, 1, maxCol).setBackground("#ffffff");
      }
    }
  }
}

// The idea is to get the list of Campaign ID names as an array and find all the
// indexes that are duplicates.

// NOTE: Capitalization is considered to be a difference and this will not raise
// an error.
function testDuplicatesIds()
{
  // Get the current sheet. This is used to verify that it is the Campaigns sheet.
  let curSheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  if(curSheet.getName() === "Campaigns")
  {
    // The range is from A2 (where the IDs start) to the end of column A.
    let allIds = SpreadsheetApp.getActive().getRange('A2:A').getValues();
    // Calling getValues() will return: [[id1], [id2], ...] flat() removes
    // the inner square brackets.
    let ids = allIds.flat();

    // Reference:
    // https://stackoverflow.com/questions/18417728/get-the-array-index-of-duplicates
    // This will return the value that has been duplicated and the indexes of the
    // duplicates.
    Array.prototype.getDuplicates = function ()
    {
      var duplicates = {};
      for (var i = 0; i < this.length; i++)
      {
        if(duplicates.hasOwnProperty(this[i]))
        {
          duplicates[this[i]].push(i);
        }
        else if (this.lastIndexOf(this[i]) !== i)
        {
          duplicates[this[i]] = [i];
        }
      }
      return duplicates;
    };

    let dupArr = ids.getDuplicates();

    let sheet = SpreadsheetApp.getActiveSheet();
    // Get the max number of columns for this sheet.
    let maxCol = sheet.getMaxColumns();

    // After finding the index of duplicates we must highlight them.
    for(const dupes in dupArr)
    {
      for(let i = 0; i < dupArr[dupes].length; i++)
      {
        // Loops through each duplicate email and lists their index.
        sheet.getRange(dupArr[dupes][i]+2, 1, 1, maxCol).setBackground("#cc4125");
        sheet.getRange(dupArr[dupes][i]+2, 1).setNote("Error: Campaign IDs cannot be duplicated.");
      }
    }
  }
}

// This function will highlight a row that contains an invalid
// sender email address in the Campaigns sheet and leave a note
// in the offending cell.
function highlightInvalidSender()
{
  let curSheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  if(curSheet.getName() === "Campaigns")
  {
    let sheet = SpreadsheetApp.getActiveSheet();
    let range = SpreadsheetApp.getActive().getRange('B2:B');
    let rangeVals = range.getValues();
    let lastRow = range.getLastRow();
    // The top level domains an email address is allowed to have.
    const allowedTlds = [".com", ".org", ".net", ".edu", ".gov"];

    // Loop through all the Senders and check that they are not empty.
    for(j = 0; j < lastRow - 1; j++)
    {
      let maxCol = sheet.getMaxColumns();

      // Clear any out-of-date notes.
      sheet.getRange(j+2, 2).clearNote();

      // Emptiness check.
      if(rangeVals[j][0] === "")
      {
        sheet.getRange(j+2, 1, 1, maxCol).setBackground("#cc4125");
        sheet.getRange(j+2, 2).setNote("Error: Sender email address cannot be empty.");
      }
      // Check for invalid email syntax.
      else if(!(rangeVals[j][0].includes('@') && rangeVals[j][0].includes('.') && allowedTlds.some(substring => rangeVals[j][0].includes(substring))))
      {
        sheet.getRange(j+2, 1, 1, maxCol).setBackground("#cc4125");
        sheet.getRange(j+2, 2).setNote("Error: Sender email address format is invalid.");
      }

      // Check if highlighting can be cleared.
      checkNoErrors(sheet.getRange(j+2, 1, 1, maxCol));
    }
  }
}