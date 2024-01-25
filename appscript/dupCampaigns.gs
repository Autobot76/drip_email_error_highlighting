// This function will check for duplicates in the Subscriptions
// AND Schedule sheet for the Campaign field. This should not be
// confused with dupCampaignIds() which checks the CampaignID in
// the Campaigns sheet.
function dupCampaigns()
{
  let curSheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  // Verify that it is the Subscriptions OR the Schedule sheet.
  if(curSheet.getName() === "Subscriptions")
  {
    // The range is from A2 (where the Campaigns start) to the end of column A.
    let allCamps = SpreadsheetApp.getActive().getRange('A2:A').getValues();

    // Calling getValues() will return: [[1], [2], ...] flat() removes
    // the inner square brackets.
    let camps = allCamps.flat();

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

    let dupArr = camps.getDuplicates();

    let sheet = SpreadsheetApp.getActiveSheet();
    // Get the max number of columns for this sheet.
    let maxCol = sheet.getMaxColumns();

    // After finding the index of duplicates we must highlight them.
    for(const dupes in dupArr)
    {
      for(let i = 0; i < dupArr[dupes].length; i++)
      {
        // Loops through each duplicate Campaigns and lists their index.
        sheet.getRange(dupArr[dupes][i]+2, 1, 1, maxCol).setBackground("#cc4125");

        // Note: The validation rule does gets triggered when a note is added.
        // If the user selects two of the same Campaign that is valid it will
        // show an additional note of saying this campaign does not exist, even
        // when it does.

        sheet.getRange(dupArr[dupes][i]+2, 1).setNote("Warning: Campaign cannot be duplicated.");
      }
    }
  }
}