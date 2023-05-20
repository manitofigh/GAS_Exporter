function moveData() {
  var inputURL = "REPLACE WITH INPUT URL";
  var outputURL = "REPLACE WITH OUTPUT URL" ;

  var inputSpreadsheet = SpreadsheetApp.openByUrl(inputURL);
  var outputSpreadsheet = SpreadsheetApp.openByUrl(outputURL);

  var inputSheet = inputSpreadsheet.getSheetByName("InputV2M");
  var outputSheet = outputSpreadsheet.getSheetByName("Sheet1");

  var numberOfBotNotesToKeep = 3;

  function getColByName(name){
    var lastColumn = outputSheet.getLastColumn();
    var headers = outputSheet.getRange(1,1,1,lastColumn).getValues();
    // Logger.log("headers: "+JSON.stringify(headers));
    var colindex = headers[0].indexOf(name);
    return colindex+1;
  }

// function checkNote(cellName){
//   var tempCellNote = cellName.getNote();
//   var botNotes = [];
//   var newNote = '';

//   // Search for sentences starting with "#!" and ending with "!#"
//   var regex = /#!(.*?)!#/g;
//   var match = regex.exec(tempCellNote);
//   while (match != null) {
//     botNotes.push(match[0]); // push the entire match (including #! and !#) to the array
//     match = regex.exec(tempCellNote);
//   }

//   // Remove oldest bot notes if array exceeds the limit
//   if (botNotes.length > numberOfBotNotesToKeep) {
//     botNotes = botNotes.slice(botNotes.length - numberOfBotNotesToKeep);
//   }

//   // Add any text that doesn't start with "#!" or end with "!#"
//   var otherNotes = tempCellNote.split(/#!.*?!#/g);
//   for (var i = 0; i < otherNotes.length; i++) {
//     if (otherNotes[i].trim() !== '') {
//       newNote += otherNotes[i].trim() + '\n'; // add the non-bot note to the newNote string with a line break
//     }
//   }

//   // Join bot notes with line breaks and add them to the newNote string
//   newNote += botNotes.join('\n');

//   cell.setNote(newNote);
// }

function checkNote(cellName) {
  var tempCellNote = cellName.getNote();
  var botNotes = [];
  var nonBotNotes = [];

  // Separate bot and human notes
  var regex = /#!(.*?)!#/g;
  var lastMatchEndIndex = 0;
  var match = regex.exec(tempCellNote);
  while (match != null) {
    var nonBotNote = tempCellNote.slice(lastMatchEndIndex, match.index);
    if (nonBotNote.trim()) {
      nonBotNotes.push(nonBotNote);
    }
    botNotes.push(match[0]);
    lastMatchEndIndex = regex.lastIndex;
    match = regex.exec(tempCellNote);
  }
  var nonBotNote = tempCellNote.slice(lastMatchEndIndex);
  if (nonBotNote.trim()) {
    nonBotNotes.push(nonBotNote);
  }

  // Remove oldest bot notes if array exceeds the limit
  if (botNotes.length > numberOfBotNotesToKeep) {
    botNotes = botNotes.slice(botNotes.length - numberOfBotNotesToKeep);
  }

  // Join bot notes with line breaks, but keep the original order
  var newNote = '';
  var i = 0;
  var j = 0;
  while (i < botNotes.length || j < nonBotNotes.length) {
    if (botNotes[i] && (!nonBotNotes[j] || tempCellNote.indexOf(botNotes[i]) < tempCellNote.indexOf(nonBotNotes[j]))) {
      newNote += botNotes[i];
      i++;
    } else if (nonBotNotes[j]) {
      newNote += nonBotNotes[j];
      j++;
    }
  }
  cellName.setNote(newNote.trim());
}

  asinColumn = getColByName("ASIN");  
  unavailableColumn = getColByName("Unavailable");
  procurementColumn = getColByName("Ready for procurement")
  supplier1Column = getColByName("Supplier 1");
  supplier2Column = getColByName("Supplier 2");
  supplier3Column = getColByName("Supplier 3");
  supplier4Column = getColByName("Supplier 4");
  supplier5Column = getColByName("Supplier 5");
  supplier6Column = getColByName("Supplier 6");
  supplier7Column = getColByName("Supplier 7");
  supplier8Column = getColByName("Supplier 8");
  supplier9Column = getColByName("Supplier 9");
  supplier10Column = getColByName("Supplier 10");

  var range = inputSheet.getRange("H1:H1040"); // set the range to scan
  var infoRange = inputSheet.getRange("O1:O1040"); // set the range for the information column
  var values = range.getValues();
  var infoValues = infoRange.getValues();
  var offset = range.getRowIndex();

  for (var i = 0; i < values.length; i++) { // loop for each row
    
    var now = new Date();
    var dateTime = Utilities.formatDate(now, "EST", "MM-dd-yyyy HH:mm");
    Logger.log("Currently running at row number ("+ (i+offset) +")");
  
    if (values[i][0] == "") { // not available
          var tempRowValue = i + offset; // Saves row index of empty cells
          var allSupplierCellsAreRed = true;
          
          // change color of Amazon.com supplier cells based on
          // product availability and update "Unavailable" field
          for (var k = supplier1Column; k <= supplier10Column; k++) { // for each supplier
            var cell = outputSheet.getRange(tempRowValue, k);
            // checkNote(cell); // activate if you want to manually test it cell by cell.
            if(cell.getValue() == "Amazon.com" && cell.getBackground() != "#f4cccc") { // if it's unavailable but the cell color isn't red:
                cell.setBackground("#f4cccc");
                var prevNote = cell.getNote();
                if (prevNote != "") { cell.setNote(prevNote + "\n"); }
                cell.setNote("#!Currently unavailable. (" + dateTime + ")!#");
                checkNote(cell);
            }
            //if a red cell is still unavailable and has had no note previously.
            else if (cell.getValue() == "Amazon.com" && cell.getBackground() == "#f4cccc" && cell.getNote() == "") { 
              checkNote(cell)
              cell.setNote("#!Still unavailable (" + dateTime +")!#")
              checkNote(cell);
              }

            // check if at least one supplier is not red
            if(allSupplierCellsAreRed == true && cell.getValue() != "" && cell.getBackground() != "#f4cccc") {
              allSupplierCellsAreRed = false;
              //Logger.log("found white supplier cell, " + cell.getValue() + ", at ["+tempRowValue+"]["+k+"]");
            }
          }

          // set "Unavailable" to yes if all supplier cells are red
          if(allSupplierCellsAreRed) {
            outputSheet.getRange(tempRowValue, unavailableColumn).setValue("yes");
            //set "Ready for procurement" to "no (unavailable)" if all the suppliers are is red and the RFP column is already set to "yes".
            if (outputSheet.getRange(tempRowValue, procurementColumn).getValue() == "yes")
             {outputSheet.getRange(tempRowValue, procurementColumn).setValue("No (Unavailable)")}
          }
          else if (allSupplierCellsAreRed == false && outputSheet.getRange(tempRowValue, unavailableColumn).getValue() == "yes"){
            outputSheet.getRange(tempRowValue, unavailableColumn).clearContent();
            if (outputSheet.getRange(tempRowValue, procurementColumn).getValue() == "No (Unavailable)")
             {outputSheet.getRange(tempRowValue, procurementColumn).setValue("yes")}
          }
      }
    else if (values[i][0] != ""){ // available
      var tempRowValue = i + offset; // Saves row index of empty cell
      var infoCell = infoValues[i][0];
      for (var k = supplier1Column; k <= supplier10Column; k++) {
        var cell = outputSheet.getRange(tempRowValue, k);
        if(cell.getValue() == "Amazon.com" && cell.getBackground() === "#f4cccc") { //It was previously unavailable but now is available.
          cell.setBackground("#FFFFFF");
          var prevNote = cell.getNote()
          cell.setNote(prevNote + "\n" + "#!Previously unavailable. Now in stock. [" + infoCell +"] " + "(" + dateTime + ")!#");
          checkNote(cell);
        }
      }
    
    }
  }
}
