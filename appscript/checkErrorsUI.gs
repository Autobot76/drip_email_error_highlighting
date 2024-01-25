// Add a special menu to Sheets that will allow the user to manually
// check for invalid data input, in case the system hasn't done so
// already.
function checkErrorsUI()
{
  // Reference: https://developers.google.com/apps-script/reference/base/menu
  let ui = SpreadsheetApp.getUi();
  ui.createMenu('Error Checker')
    .addItem('Check', 'onEdit')
    .addToUi();
}