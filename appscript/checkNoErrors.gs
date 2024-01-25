// The purpose of this function is to check if all errors (i.e. highlighting) has been
// addressed in a row and if it has remove the highlighting from the row.
// curRow is a Range (https://developers.google.com/apps-script/reference/spreadsheet/range)
function checkNoErrors(curRow)
{
  // Try to prove this to be false. A single note will make this invalid.
  let valid = true;

  let curRowArr = curRow.getNotes();
  // Loop through each row item and even if there is one note then it is deemed invalid
  // if there are no notes then it is valid so clear highlighting.
  curRowArr[0].forEach(function (note) {
    Logger.log(note);
    if(note !== '')
    {
      valid = false;
    }
  });

  // If everything is valid the highlighting is no longer needed.
  if(valid)
  {
    curRow.setBackground("#ffffff");
  }
}