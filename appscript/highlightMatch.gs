//add custom menu to trigger search
function addTriggerSearchMenu()
{
  var ui = SpreadsheetApp.getUi();
  // Or DocumentApp or FormApp.
  ui.createMenu('Highlight tool')
      .addItem('Find and Highlight', 'startFind')
      .addToUi();
}

function startFind()
{
  var ui = SpreadsheetApp.getUi();

  //ask for user input a keyword to search for
  var result = ui.prompt("Please enter a column to search for");

  var button = result.getSelectedButton();

  if(button != ui.Button.OK)
  {
    return;
  }

  //first: get column index
  let columnString = result.getResponseText();

  let columnIndex = findColumnIndex(columnString.toLowerCase());

  if(columnIndex < 0)
  {
    ui.alert("Can't find column " + "\"" + columnString + "\"" + " in this sheet");
    return;
  }

  console.log("find column index: " + columnIndex);

  //second: get keyword
  //ask for user input a keyword to search for
  var result = ui.prompt("Please enter a keyword to search for");

  var button = result.getSelectedButton();

  if(button != ui.Button.OK)
  {
    return;
  }

  let searchWord = result.getResponseText();
  highlightMatch(searchWord, columnIndex);
}

//find column index by input name
function findColumnIndex(columnName)
{

  //get current sheet
  var curSheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  //get last column index
  var maxCol = curSheet.getLastColumn();
  //current column j+1: range from 1 to last
  var sourceRange = curSheet.getRange(1, 1, 1, maxCol);
  //get display values to compare string
  var sourceValues = sourceRange.getDisplayValues();

  // search every cell in row 1 from A1 to the last column
    for (var i = 0; i < maxCol; i++) {
      let currentValue = sourceValues[0][i];
      if (currentValue.toLowerCase() == columnName) 
      {
        // return the column number if we find it
        return(i+1);
      }
    }
    // return -1 if it doesn't exist
    return(-1);
}

//sheet content changed, update match
function updateMatch()
{
  //get current sheet
  var curSheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
   
  //get cached keyword
  let metadatas = curSheet.getDeveloperMetadata();
  if(metadatas.length > 0)
  {
    var historyKeyword = '';
    var historyIndex = -1;
    
    metadatas.forEach(
      //Loop through every element of the meta data
      function(thisMetaData) {
        let key = thisMetaData.getKey();
        let v = thisMetaData.getValue();
        console.log("history keyword:" + historyKeyword);
        if (key == 'historyKeyword')
        {
          historyKeyword = v;
        } else if (key == 'historyColumn')
        {
          historyIndex = parseInt(v);
        }
      }
    )

    if (historyKeyword.length > 0 && historyIndex >= 0)
    {
      highlightMatch(historyKeyword, historyIndex);
    }
  }
}

function highlightMatch(searchWord, columnIndex)
{
    //get current sheet
    var curSheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

    //get cached keyword
    let metadatas = curSheet.getDeveloperMetadata();
    if (metadatas.length > 0) {
        metadatas.forEach(
        //Loop through every element of the meta data
        function(thisMetaData) {
            let key = thisMetaData.getKey();
            if (key == 'historyKeyword') {
                thisMetaData.remove();
            } else if (key == 'historyColumn') {
                thisMetaData.remove();
            }
        })
    }
    //cache the new keyword
    curSheet.addDeveloperMetadata("historyKeyword", searchWord, SpreadsheetApp.DeveloperMetadataVisibility.DOCUMENT);
    curSheet.addDeveloperMetadata("historyColumn", columnIndex, SpreadsheetApp.DeveloperMetadataVisibility.DOCUMENT);

    //search and highlight
    searchBy(searchWord, columnIndex);
    updateMatch();
}

//has to be placed in the onEdit(), so it can do the search and matching while editing is made

//need to check on current highlighing conditions and remove existing highlight if highlight needs to be performed at the location

function searchBy(keyword, columnIndex)
{
  
  //get current sheet
  var curSheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
   
  //get last row index
  var maxRow = curSheet.getLastRow();
  //get last column index
  var maxCol = curSheet.getLastColumn();
  
  //cache matcheRanges first
  var matcheRanges = []

  //current column j+1: range from 1 to last
    var sourceRange = curSheet.getRange(1, columnIndex, maxRow);
    //get value of these cells
    //get display values to compare string
    var sourceValues = sourceRange.getDisplayValues();
    // console.log(sourceValues);
  
    //start from row 2, ignore first header row
    for(var i = 1; i < maxRow; i++)
    {
      //cell string value in row i, lastColumn j
      let currentValue = sourceValues[i][0];

      let currentRowRange = curSheet.getRange(i+1, 1, 1, maxCol);

      //compare between lowercased string
      if(currentValue.toLowerCase().includes(keyword.toLowerCase()))
      {
        //note that range start from 1
        matcheRanges.push(currentRowRange);
      } 
      else
      {
        if (didMarkHighlighted(currentRowRange))
        {
          // ⚠️clean row highlight background at columnIndex⚠️
          console.log('backup highlighted row: ' + i+1);
          currentRowRange.setNote(null);
          currentRowRange.setBackground('white');
        }
      }
    }

  //after clean all history highlighted cells, then set match cells
  matcheRanges.forEach(
    //Loop through every element of the meta data
    function(matchRange) {
      if(shouldMarkHighlighted(matchRange))
      {
        //set match color background: green
        matchRange.setBackground('green');
      }
    }
  )
}

function shouldMarkHighlighted(range)
{
  range.getNotes()[0].forEach(
    function(note) {
      if(note == 'Mark Highlighted')
      {
        return false;
      }
    }
  )

  var bgColors = range.getBackgrounds();
  for (var i in bgColors)
  {
    for (var j in bgColors[i])
    {

      //if any cell background color is not white, will not mark highlight
      var color = bgColors[i][j];
      if(color != '#ffffff')
      {
        return false;
      }
    }
  }

  return true;
}

function didMarkHighlighted(range)
{
  var did = false;
  range.getNotes()[0].forEach(
    function(note) {
      if(note == 'Mark Highlighted')
      {
        did = true;
      }
    }
  )

  return did;
}