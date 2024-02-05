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

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Monte Carlo')
    .addItem('Run', 'runmc')
    .addToUi();
}


function getnormrand(min, max, mean = (min + max) / 2, stdDev = (max - min) / 6) {
  while (true) {
    // Use Box-Muller transform to generate normal distribution
    const u = Math.random();
    const v = Math.random();
    const z = Math.sqrt(-2.0 * Math.log(u)) * Math.cos(2.0 * Math.PI * v);

    const value = mean + z * stdDev;

    // Ensure value falls within the specified range
    if (value >= min && value <= max) {
      return value.toPrecision(3);
    }
  }
}


function getsheetname_from_a1notation(a1notation) {
  return a1notation.split("!")[0];
}

function get_input_values_and_ranges(spreadsheet, named_ranges) {
  let input_values_and_ranges = {};

  for (var key in named_ranges) {
    if (key.endsWith("_mci")) {
      nrkey = key;
      // alert(named_ranges[nrkey]);
      input_cell = spreadsheet.getRange(named_ranges[nrkey]);
      input_sheet = spreadsheet.getSheetByName(getsheetname_from_a1notation(named_ranges[nrkey]));
      min_range = input_sheet.getRange(input_cell.getRow(), input_cell.getColumn() + 1).getValue();
      max_range = input_sheet.getRange(input_cell.getRow(), input_cell.getColumn() + 2).getValue();
      input_values_and_ranges[nrkey] = [input_cell, min_range, max_range];
    }
  }
  return input_values_and_ranges;
}

function get_output_cells(spreadsheet, named_ranges) {
  let output_cells = {};

  for (var key in named_ranges) {
    if (key.endsWith("_mco")) {
      nrkey = key;
      // alert(named_ranges[nrkey]);
      output_cell = spreadsheet.getRange(named_ranges[nrkey]);
      output_cells[nrkey] = output_cell;
    }
  }
  return output_cells;
}

function runmc() {
  // SpreadsheetApp.getUi().alert(get_run_values()["salary_increase"]);
  // Get the active spreadsheet
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  // Check if the "MC" sheet already exists
  var mcsheet = spreadsheet.getSheetByName("MC");
  start_fresh = false;
  // If the sheet doesn't exist, create it
  if (!mcsheet) {
    mcsheet = spreadsheet.insertSheet("MC");
    start_fresh = true;
  } 

  named_ranges = getNamedRanges();
  input_values_and_ranges = get_input_values_and_ranges(spreadsheet, named_ranges);
  output_cells = get_output_cells(spreadsheet, named_ranges);

  // Initialize data structures for recording all the outputs
  mc_outputs = {}
  for (o in output_cells) {
    mc_outputs[o] = []
  }


  // RUN THE MODEL X TIMES
  let NUM_RUNS = 100;
  for (let i = 0; i < NUM_RUNS; i++) {
    // Change all the inputs randomly (within normal distribution)
    for (key in input_values_and_ranges) {
      inr = input_values_and_ranges[key];
      inr_run_value = getnormrand(inr[1], inr[2]);
      // Set value
      inr[0].setValue(inr_run_value);
    }

    // Save all the outputs
    for (o in output_cells) {
      orun_value = output_cells[o].getValue();
      mc_outputs[o].push(orun_value);
    }
  }

  oo = spreadsheet.getRange(named_ranges["output_order"]);
  oo_values = oo.getValues();
  // OUTPUT RESULTS
  if (start_fresh == true) {
    // Clear the sheet content (optional)
    mcsheet.clear();
    // SpreadsheetApp.getUi().alert(JSON.stringify(getNamedRanges()));
    insertSectionHeader(mcsheet, 1, 1, ["Named ranges used"]);
    insertTableHeader(mcsheet, 2, 1, ["Named reference", "A1notation"]);
    insertKeyValues(mcsheet, 3, 1, dictToList(named_ranges));

    next_row = 3 + dictToList(named_ranges).length;
    next_row++;
    insertSectionHeader(mcsheet, next_row++, 1, ["Outputs"]);

    output_header = [];
    for(oi = 0; oi < oo_values.length; oi++) {
      output_header.push(oo.getValues()[oi][1]);
    }
    insertTableHeader(mcsheet, next_row++, 1, ["Run number"].concat(output_header));

  } else {
    var Avals = mcsheet.getRange("A1:A").getValues();
    next_row = Avals.filter(String).length + 2;
    // alert("didn't start fresh, starting on row: " + next_row)
  }

  for (let run = 0; run < NUM_RUNS; run++) {
    next_col = 1;
    var cell = mcsheet.getRange(next_row, next_col++);
    cell.setValue(run);
    for (j = 0; j < oo_values.length; j++) {
      okey = oo_values[j][0];
      cell = mcsheet.getRange(next_row, next_col++);
      cell.setValue(mc_outputs[okey][run]);
      cell.setNumberFormat("£#,##0.00;$(#,##0.00)");
    }
    next_row++;
  }
}

