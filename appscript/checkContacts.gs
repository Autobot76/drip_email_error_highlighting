// The purpose here is to try to declutter all the calls
// in triggers.gs
function checkContacts()
{
  highlightInvalidEmails();
  testDuplicatesEmail();
  highlightInvalidTz();
  checkNames();
  checkEmailNameValid();
}

// Purpose: Check to see if red highlighting from Email to LastName
// can be removed.
function checkEmailNameValid()
{
  let curSheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  if(curSheet.getName() === "Contacts")
  {
    let sheet = SpreadsheetApp.getActiveSheet();
    let range = SpreadsheetApp.getActive().getRange('A2:A');
    let lastRow = range.getLastRow();
    let valid = true;

    for(j = 0; j < lastRow - 1; j++)
    {
      // Get all the notes for this row, excluding the Timezone column.
      let curRowNotesArr = sheet.getRange(j+2, 1, 1, 3).getNotes();

      // curRowNotesArr's data is actually in the inner array.
      // Loop through all the notes from Email to LastName, if even 1 is present
      // this whole data set is considered invalid.
      // If there are no notes present just remove the red highlighting.
      curRowNotesArr[0].forEach(function (note) {
        if(note !== '')
        {
          valid = false;
        }
      });

      if(valid)
      {
        sheet.getRange(j+2, 1, 1, 3).setBackground("#ffffff");
      }

      // Prepare for the next round of looping.
      valid = true;
    }
  }
}

// Purpose: Check if the FirstName & LastName field is empty, if so
// leave a note and highlight the row.
function checkNames()
{
  let curSheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  if(curSheet.getName() === "Contacts")
  {
    let sheet = SpreadsheetApp.getActiveSheet();
    let range = SpreadsheetApp.getActive().getRange('B2:C');
    let rangeVals = range.getValues();
    let lastRow = range.getLastRow();

    for(j = 0; j < lastRow - 1; j++)
    {
      // Check FirstName
      sheet.getRange(j+2, 2).clearNote();
      if(rangeVals[j][0] === "")
      {
        sheet.getRange(j+2, 1, 1, 3).setBackground("#cc4125");
        sheet.getRange(j+2, 2).setNote("Error: FirstName cannot be empty.");
      }

      // Check LastName
      sheet.getRange(j+2, 3).clearNote();
      if(rangeVals[j][1] === "")
      {
        sheet.getRange(j+2, 1, 1, 3).setBackground("#cc4125");
        sheet.getRange(j+2, 3).setNote("Error: LastName cannot be empty.");
      }
    }
  }
}

// Purpose: Highlight all rows in red that contain an invalid email address.
function highlightInvalidEmails()
{
  // Get the current sheet. This is used later on to verify that it is the Contacts sheet.
  let curSheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

  // Currently if it is the contact sheet verify the emails.
  if(curSheet.getName() === "Contacts")
  {
    // Emptiness check works.
    let sheet = SpreadsheetApp.getActiveSheet();
    let range = SpreadsheetApp.getActive().getRange('A2:A');
    let rangeVals = range.getValues();
    let lastRow = range.getLastRow();

    // The top level domains an email address is allowed to have.
    const allowedTlds = [".com", ".org", ".net", ".edu", ".gov"];

    // Adapted from: https://yagisanatode.com/2017/12/13/google-apps-script-iterating-through-ranges-in-sheets-the-right-and-wrong-way/
    // Loop through all the emails in the spread sheet and check that they are not empty.
    for(j = 0; j < lastRow - 1; j++)
    {
      // Get the max number of columns for this sheet.
      let maxCol = sheet.getMaxColumns();

      // Clear any out of date notes.
      sheet.getRange(j+2, 1).clearNote();

      if(rangeVals[j][0] === "")
      {
        // The cells that are empty will now be highlighted.
        // j+2 is the row we are starting our range from.
        // 1 is the starting column of our range. It is always one because the first column is the emails.
        // 1 is the number of rows, we are only interested in the row with the invalid data.
        // maxCol is the number of columns to include, we want to highlight the whole row.
        sheet.getRange(j+2, 1, 1, maxCol).setBackground("#cc4125");

        // Add a note telling the user that the email cannot be empty.
        sheet.getRange(j+2, 1).setNote("Error: Email cannot be empty.");
      }
      // Check that the email contains @ and . which is common for all emails addresses.
      // The last part of this statement checks if the inputted email address contains the allowed TLDs.
      // Reference:
      // https://stackoverflow.com/questions/5582574/how-to-check-if-a-string-contains-text-from-an-array-of-substrings-in-javascript (Praveena)
      else if(!(rangeVals[j][0].includes('@') && rangeVals[j][0].includes('.') && allowedTlds.some(substring => rangeVals[j][0].includes(substring))))
      {
        sheet.getRange(j+2, 1, 1, maxCol).setBackground("#cc4125");
        sheet.getRange(j+2, 1).setNote("Error: Invalid email address was provided.");
      }

      // Check if highlighting can be cleared.
      checkNoErrors(sheet.getRange(j+2, 1, 1, maxCol));
    }
  }
}

// The idea is to get the list of emails as an array and find all the indexes
// that are duplicates.
function testDuplicatesEmail()
{
  // Get the current sheet. This is used to verify that it is the Contacts sheet.
  let curSheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  if(curSheet.getName() === "Contacts")
  {
    // The range is from A2 (where emails start) to the end of column A.
    let allEmails = SpreadsheetApp.getActive().getRange('A2:A').getValues();
    // Calling getValues() will return: [[email1], [email2], ...] flat() removes
    // the inner square brackets.
    let em = allEmails.flat();

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

    let dupArr = em.getDuplicates();

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
        sheet.getRange(dupArr[dupes][i]+2, 1).setNote("Warning: Emails cannot be duplicated.");
      }
    }
  }
}

// Highlight all Timezone cells in yellow that contain invalid time zones.
function highlightInvalidTz()
{
  // Time zone should be formatted as: xx:xx+/-
  // https://www.ge.com/digital/documentation/meridium/V36160/Help/Master/Subsystems/AssetPortal/Content/Time_Zone_Mappings.htm

  // 12:00- to 13:00+

  // Get the current sheet. This is used later on to verify that it is the Contacts sheet.
  let curSheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

  // Currently if it is the contact sheet verify the time zone.
  // This will change later on.
  if(curSheet.getName() === "Contacts")
  {
    // Cases:
    // 14:00-
    // 20:00-
    // According to regex this is fine but in reality it isn't.
    // These time zones don't exist.
    // Implement a new approach by splitting the string by the :
    // const validTzRe = new RegExp(/[0-9]?[0-9]:[0-9][0-9](\-|\+)/);

    const colonRe = new RegExp(/:/);

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
        // j+2 is the row we are starting our range from. j starts at zero but to getRange() it starts at 1.
        // 4 is the starting column of our range. It is always 4 because the fourth column is the timezone.
        sheet.getRange(j+2, 4).setBackground("#ffea00");
        sheet.getRange(j+2, 4).setNote("Warning: Time zone should not be empty.");
      }
      // Check if a colon exists. If it doesn't it's not a time zone.
      else if(!colonRe.test(rangeVals[j][0]))
      {
        sheet.getRange(j+2, 4).setBackground("#ffea00");
        sheet.getRange(j+2, 4).setNote("Warning: Time zone is not in a valid format.");
      }
      // If it is not of type string then it's wrong.
      else if(typeof rangeVals[j][0] !== "string")
      {
        sheet.getRange(j+2, 4).setBackground("#ffea00");
        sheet.getRange(j+2, 4).setNote("Warning: Time zone is not in a valid format.");
      }
      else
      {
        // Split the string by the :
        let curStrArr = rangeVals[j][0].split(":");

        // There cannot be more than 2 digits to the left of the colon.
        // There cannot be more than 3 digits to the right of the colon. +1 for +/-
        if(curStrArr[0].length > 2.0 || curStrArr[1].length > 3.0)
        {
          sheet.getRange(j+2, 4).setBackground("#ffea00");
          sheet.getRange(j+2, 4).setNote("Warning: Time zone is not in a valid format.");
        }
        // If there are two digits to the left of the colon,
        // check that both of them are numbers and not something else.
        else if(curStrArr[0].length === 2.0)
        {
          if(isNaN(parseInt(curStrArr[0].charAt(0))) || isNaN(parseInt(curStrArr[0].charAt(1))))
          {
            sheet.getRange(j+2, 4).setBackground("#ffea00");
            sheet.getRange(j+2, 4).setNote("Warning: Time zone is not in a valid format.");
          }
          // If there are 2 numbers, both numbers cannot be greater than 13 for positives.
          else if(parseInt(curStrArr[0]) > 13)
          {
            sheet.getRange(j+2, 4).setBackground("#ffea00");
            sheet.getRange(j+2, 4).setNote("Warning: Time zone is not in a valid format.");
          }
          // If the first 2 digits to the right of the colon is not a number then it is invalid.
          else if(isNaN(parseInt(curStrArr[1].charAt(0))) || isNaN(parseInt(curStrArr[1].charAt(1))))
          {
            sheet.getRange(j+2, 4).setBackground("#ffea00");
            sheet.getRange(j+2, 4).setNote("Warning: Time zone is not in a valid format.");
          }
          // This case addresses 13:01, 13:02 and so on, which is an invalid time zone.
          else if(parseInt(curStrArr[0]) === 13 && parseInt(curStrArr[1].charAt(0)) !== 0 || parseInt(curStrArr[1].charAt(1)) !== 0)
          {
            sheet.getRange(j+2, 4).setBackground("#ffea00");
            sheet.getRange(j+2, 4).setNote("Warning: Time zone is not in a valid format.");
          }
        }
        // If the digits to the left of the colon is not a number then it is invalid.
        else if(isNaN(parseInt(curStrArr[0])))
        {
          sheet.getRange(j+2, 4).setBackground("#ffea00");
          sheet.getRange(j+2, 4).setNote("Warning: Time zone is not in a valid format.");
        }
        // If the first 2 digits to the right of the colon is not a number then it is invalid.
        else if(isNaN(parseInt(curStrArr[1].charAt(0))) || isNaN(parseInt(curStrArr[1].charAt(1))))
        {
          sheet.getRange(j+2, 4).setBackground("#ffea00");
          sheet.getRange(j+2, 4).setNote("Warning: Time zone is not in a valid format.");
        }
        // Check that the left most digit to the right of the colon is in range 0 to 5
        // There is no need to check for the second digit because it cannot be greater
        // than 9 anyways due to how numbers work.
        else if(parseInt(curStrArr[1].charAt(0)) > 5)
        {
          sheet.getRange(j+2, 4).setBackground("#ffea00");
          sheet.getRange(j+2, 4).setNote("Warning: Time zone is not in a valid format.");
        }
      }
      // Check if highlighting can be cleared.
      checkNoErrors(sheet.getRange(j+2, 1, 1, sheet.getMaxColumns()));
    }
  }
}