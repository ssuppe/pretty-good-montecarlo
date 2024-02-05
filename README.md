# pretty-good-montecarlo
Monte Carlo for Google Sheets using Google Apps Script (no remote data sharing)

This is a bare bones but pretty useful Monte Carlo-type engine for Google Sheets. It uses javascript and everything executes within the spreadsheet, so nothing leaves the browser. (Well, something might go to Google to run in sheets, I'm not sure, but it doesn't send anything to anything I own).

# Quickstart

Make a copy of the [demo spreadsheet](https://docs.google.com/spreadsheets/d/1tk-_XRlynu7NkyomjNsSYGbdB1qTGoSdjVB1ckgyl30/edit#gid=0) and follow the instructions in the "Variables" tab.

# How it works - high level

Pretty Good Monte Carlo allows you to run randomized trials of as many input variables as you need, recording the results of as many output variables as you need. It then writes the results of all those randomized trials to a sheet called "MC." To stay within Google Apps Script time limit, it only runs 100 trials at a time. By running multiple times, you can continue to append more and more trials.

A few caveates:
*  This is currently for power users who are familiar with things like named ranges
*  The only distribution that is supported is the Normal distribution.
*  There is no data or formatting validation. The script expects everything to be written correctly and well formed, or will die ungracefully. That being said, it's not so difficult.

# How to import the code into your own existing sheet
1) Go to Extensions > Apps Script
1) Create two files (Code.gs and Util.gs). Copy/paste the code from the files in this repository into those two files respectively.
1) Reload the spreadsheet. You should now have a menu item at the top called "Monte Carlo"
1) Create your named ranges for inputs/outs as specified in the [demo spreadsheet](https://docs.google.com/spreadsheets/d/1tk-_XRlynu7NkyomjNsSYGbdB1qTGoSdjVB1ckgyl30/edit#gid=0), or as described below
1) Make sure you don't have a sheet called "MC" (PCG expects a blank one when it runs a new model)
1) Press Monte Carlo > Run
1) Wait
1) See your results on the "MC" sheet

# How variables work
Pretty Good Monte Carlo does not have a custom UI - instead it uses [named ranges](https://support.google.com/docs/answer/63175?hl=en-GB&co=GENIE.Platform%3DDesktop) with standardized names. These variables can be anywhere on any sheet, but must follow the naming convention.

1) **Input variables** are the variables that will be randomly selected. The named range must end with "_mci" ("monte carlo input"). It should reference the cell that PCG will insert the randomly selected value. The next two cells in the same row should be the minimum and maximum of the range. PCG will randomly pick a number from that range, over a normal distribution, using the Box-Muller method. You can have as many as you'd like.
2) **Output variables** are the variables that have the final answer of your model. The named range must end with "_mco" ("monte carlo output"). PCG will record the result of your model on each of its randomized trials. You can have as many as you'd like.
3) **Output order** - you must have a 2 column table, with a named range of "output_order." The first column is the name of each output variable. The 2nd column is the human-friendly/pretty printed name you'd like PCG to output to the "MC" results tab as a header.
