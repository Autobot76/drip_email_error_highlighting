// The purpose of this is to contain all the calls required
// to check templates in one function instead of many.
function checkTemplates()
{
  highlightInvalidTemplateName();
  highlightInvalidSendGridId();
  dupTemplateNames();
}

// Highlight all TemplateName that is invalid in red.
function highlightInvalidTemplateName()
{
  // Get the current sheet. This is used to verify that it is the Templates sheet.
  let curSheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

  // Currently if it is the Templates sheet then verify the template names.
  if(curSheet.getName() === "Templates")
  {
    let sheet = SpreadsheetApp.getActiveSheet();
    let range = SpreadsheetApp.getActive().getRange('A2:A');
    let rangeVals = range.getValues();
    let lastRow = range.getLastRow();

    // Adapted from: https://yagisanatode.com/2017/12/13/google-apps-script-iterating-through-ranges-in-sheets-the-right-and-wrong-way/
    // Loop through all the Template names in the sheet and check that they are not empty.
    for(j = 0; j < lastRow - 1; j++)
    {
      // Get the max number of columns for this sheet.
      let maxCol = sheet.getMaxColumns();

      // Clear any out-of-date notes.
      sheet.getRange(j+2, 1).clearNote();

      // Emptiness check.
      if(rangeVals[j][0] === "")
      {
        // The cells that are empty will now be highlighted.
        // j+2 is the row we are starting our range from.
        // 1 is the starting column of our range. It is always 1 because the first column is the TemplateName.
        // 1 is the number of rows, we are only interested in the row with the invalid data.
        // maxCol is the number of columns to include, we want to highlight the whole row.
        sheet.getRange(j+2, 1, 1, maxCol).setBackground("#cc4125");

        // Add a note telling the user that the TemplateName cannot be empty.
        sheet.getRange(j+2, 1).setNote("Error: TemplateName cannot be empty.");
      }
      checkNoErrors(sheet.getRange(j+2, 1, 1, maxCol));
    }
  }
}

// Highlight all invalid SendGridID in red and leave a note.
function highlightInvalidSendGridId()
{
  // Get the current sheet. This is used to verify that it is the Templates sheet.
  let curSheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

  // Currently if it is the Templates sheet then verify the SendGrid IDs.
  if(curSheet.getName() === "Templates")
  {
    let sheet = SpreadsheetApp.getActiveSheet();
    let range = SpreadsheetApp.getActive().getRange('B2:B');
    let rangeVals = range.getValues();
    let lastRow = range.getLastRow();

    // Adapted from: https://yagisanatode.com/2017/12/13/google-apps-script-iterating-through-ranges-in-sheets-the-right-and-wrong-way/
    // Loop through all the SendGrid IDs in the sheet and check that they are not empty.
    for(j = 0; j < lastRow - 1; j++)
    {
      // Clear any out-of-date notes.
      sheet.getRange(j+2, 2).clearNote();

      // Emptiness check.
      if(rangeVals[j][0] === "")
      {
        // The cells that are empty will now be highlighted.
        // j+2 is the row we are starting our range from.
        // 2 is the starting column of our range. It is always 2 because the 2nd column is the SendGrid ID.
        sheet.getRange(j+2, 2).setBackground("#cc4125");

        // Add a note telling the user that the SendGrid ID cannot be empty.
        sheet.getRange(j+2, 2).setNote("Error: SendGrid ID cannot be empty.");
      }
      else
      {
        // Check if highlighting can be removed for this row.
        checkNoErrors(sheet.getRange(j+2, 1, 1, sheet.getMaxColumns()));
      }
    }
  }
}

// This function will highlight any duplicate TemplateName in red and add a note.
function dupTemplateNames()
{
  let curSheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  // Verify that it is the Templates sheet.
  if(curSheet.getName() === "Templates")
  {
    // The range is from A2 (where the TemplateName start) to the end of column A.
    let allTN = SpreadsheetApp.getActive().getRange('A2:A').getValues();

    // Calling getValues() will return: [[1], [2], ...] flat() removes
    // the inner square brackets.
    let tn = allTN.flat();

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

    let dupArr = tn.getDuplicates();

    let sheet = SpreadsheetApp.getActiveSheet();
    // Get the max number of columns for this sheet.
    let maxCol = sheet.getMaxColumns();

    // After finding the index of duplicates we must highlight them.
    for(const dupes in dupArr)
    {
      for(let i = 0; i < dupArr[dupes].length; i++)
      {
        // Loops through each duplicate TemplateNames and lists their index.
        sheet.getRange(dupArr[dupes][i]+2, 1, 1, maxCol).setBackground("#cc4125");
        sheet.getRange(dupArr[dupes][i]+2, 1).setNote("Error: TemplateName cannot be duplicated.");
      }
    }
  }
}