/*
* Pretty Good Monte Carlo - Monte Carlo for Google Sheets using Google Apps Script (no remote data sharing)
* Copyright (C) 2024  Steven Suppe
*
* This program is free software: you can redistribute it and/or modify
* it under the terms of the GNU General Public License as published by
* the Free Software Foundation, either version 3 of the License, or
* (at your option) any later version.
*
* This program is distributed in the hope that it will be useful,
* but WITHOUT ANY WARRANTY; without even the implied warranty of
* MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
* GNU General Public License for more details.
*
* You should have received a copy of the GNU General Public License
* along with this program.  If not, see <http://www.gnu.org/licenses/>.
*/

function getNamedRanges() {
  // Get the active spreadsheet
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

  // Create an empty dictionary to store the named ranges
  var namedRanges = {};

  // Iterate through all the sheets in the spreadsheet
  spreadsheet.getSheets().forEach(function(sheet) {
    // Get the named ranges for the current sheet
    var sheetNamedRanges = sheet.getNamedRanges();

    // Iterate through the named ranges and add them to the dictionary
    sheetNamedRanges.forEach(function(namedRange) {
      // Get the A1 notation with sheet name prepended
      var range = namedRange.getRange();
      var a1Notation = sheet.getName() + "!" + range.getA1Notation();

      namedRanges[namedRange.getName()] = a1Notation;
    });
  });

  // Return the dictionary of named ranges
  return namedRanges;
}

function dictToList(dict) {
var arr = [];

for (var key in dict) {
    if (dict.hasOwnProperty(key)) {
        arr.push( [ key, dict[key] ] );
    }
}
return arr;
}

function insertSectionHeader(sheet, startRow, startCol, headerNames) {
  // Get the spreadsheet and sheet
  // var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  // var sheet = spreadsheet.getSheetByName(sheetName);

  // Get the header range
  var headerRange = sheet.getRange(startRow, startCol, 1, headerNames.length);

  // Write the header names to the range
  headerRange.setValues([headerNames]);

  // Apply formatting
  // Set background color
  headerRange.setBackground("#cdcdcd");  // Light gray background

  // Set font weight (for bold text)
  headerRange.setFontWeight("bold")
}

function insertTableHeader(sheet, startRow, startCol, headerNames) {
  // Get the spreadsheet and sheet
  // var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  // var sheet = spreadsheet.getSheetByName(sheetName);

  // Get the header range
  var headerRange = sheet.getRange(startRow, startCol, 1, headerNames.length);

  // Write the header names to the range
  headerRange.setValues([headerNames]);

  // Apply formatting
  // Set background color
  headerRange.setBackground("#f0f0f0");  // Light gray background

  // Set font weight (for bold text)
  headerRange.setFontWeight("bold");
}


function insertKeyValues(sheet, startRow, startCol, keyValues) {
  // Get the range to write the key-value pairs
  var dataRange = sheet.getRange(startRow, startCol, keyValues.length, keyValues[0].length);

  // Write the key-value pairs to the range
  dataRange.setValues(keyValues);

  // Bold the first column
  var firstColumnRange = dataRange.offset(0, 0, keyValues.length, 1);
  firstColumnRange.setFontWeight("bold");
}

function alert(data) {
  SpreadsheetApp.getUi().alert(JSON.stringify(data));
}
