`
Overview:
The script consists of two main functions:

1. sortColumnsByRow9: This function sorts the columns based on the values in row 9 and reorders them accordingly.
2. onOpen: This function adds a custom menu to the Google Sheets UI to trigger the sortColumnsByRow9 function.

#Code Explanation
'sortColumnsByRow9' Function
This function performs the following steps:
  1. Retrieve the Active Sheet: Accesses the active sheet of the active spreadsheet.
  2. Determine Number of Columns and Rows: Calculates the number of columns (excluding column A) and the total number of rows.
  3. Read Data and Formulas: Reads all data from columns B onwards and captures the formulas in row 9.
  4. Capture Headers and Row 9 Values: Stores the headers and the values/formulas in row 9 for sorting.
  5. Create Columns Array: Creates an array of objects, each representing a column with its header, row 9 value, data, and formula.
  6. Sort Columns: Sorts the columns based on the values in row 9 in descending order.
  7. Rearrange and Write Data: Rearranges the data in sorted order and writes it back to the sheet.
  8. Reapply Headers and Formulas: Reapplies the headers and formulas to the sorted columns.
`

function sortColumnsByRow9() {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var numColumns = sheet.getLastColumn() - 1; // Columns B onwards
    var numRows = sheet.getMaxRows(); // Get the number of rows used
  
    // Read all data and formulas at once
    var dataRange = sheet.getRange(1, 2, numRows, numColumns); // Columns B onwards
    var allData = dataRange.getValues();
    var row9Formulas = [];
  
    // Capture row 9 formulas and headers
    var headers = sheet.getRange(1, 2, 1, numColumns).getValues()[0];
    for (var i = 0; i < numColumns; i++) {
      var formula = sheet.getRange(9, i + 2).getFormula();
      row9Formulas.push(formula);
    }
  
    // Create an array to hold the columns and their row 9 values
    var columns = [];
  
    for (var i = 0; i < numColumns; i++) {
      var colValues = [];
      for (var j = 0; j < numRows; j++) {
        colValues.push(allData[j][i]);
      }
      var row9Value = colValues[8]; // Row 9 value (index 8 for zero-based index)
      columns.push({ index: i, value: row9Value, data: colValues, header: headers[i], formula: row9Formulas[i] });
    }
  
    // Sort columns based on row 9 values
    columns.sort(function(a, b) {
      return b.value - a.value;
    });
  
    // Create a new array to hold the sorted data
    var sortedData = [];
    for (var i = 0; i < numRows; i++) {
      sortedData.push(new Array(numColumns));
    }
  
    // Rearrange columns in the sortedData array
    for (var j = 0; j < columns.length; j++) {
      var colData = columns[j].data;
      for (var row = 0; row < numRows; row++) {
        sortedData[row][j] = colData[row];
      }
    }
  
    // Write the sorted data back to the sheet
    dataRange.setValues(sortedData);
  
    // Reapply the headers
    var headerRange = sheet.getRange(1, 2, 1, numColumns);
    headerRange.setValues([columns.map(function(col) { return col.header; })]);
  
    // Reapply the formulas in row 9
    for (var k = 0; k < columns.length; k++) {
      var row9Range = sheet.getRange(9, k + 2);
      row9Range.setFormula(columns[k].formula);
    }
  }
  
  function onOpen() {
    var ui = SpreadsheetApp.getUi();
    ui.createMenu('Custom Scripts')
      .addItem('Sort Columns by Row 9', 'sortColumnsByRow9')
      .addToUi();
  }