# sort-columns-by-row-script

# Google Sheets Script: Sort Columns by Row 9

This repository contains a Google Sheets script created to organize a spreadsheet that tracks the need for budget increases. The script identifies the highest sum of costs from the 9th row (which displays the totals in each column) and then shifts the columns to the left to ensure everything is sorted in descending order from highest cost to lowest cost.

## Overview

The script consists of two main functions:

- `sortColumnsByRow9`: This function sorts the columns based on the values in row 9 and reorders them accordingly.
- `onOpen`: This function adds a custom menu to the Google Sheets UI to trigger the `sortColumnsByRow9` function.

## Purpose

The purpose of this script is to help organize a Google Sheets spreadsheet that tracks budget increases by identifying and sorting columns based on the total costs in row 9. The columns are shifted to ensure that the highest costs are on the left, making it easier to prioritize budget increases.

## Features

- **Preserves Existing Formulae**: Unlike previous versions, this script does not remove existing formulae, ensuring that any conditional formatting remains intact and functional.
- **Conditional Formatting Compatibility**: The script works well with conditional formatting, such as a traffic light system (red, amber, green) to highlight costs within certain ranges and identify the urgency and importance of budget raises needed.

## Usage

1. Open your Google Sheets document.
2. Go to `Extensions > Apps Script` and paste the code into the script editor.
3. Save the script project.
4. Reload the Google Sheets document.
5. A new menu item `Custom Scripts` will appear in the menu bar.
6. Click on `Custom Scripts` and select `Sort Columns by Row 9` to sort the columns based on the values in row 9.

## Example
![Demo of My Project](https://github.com/isaac-nabe/sort-columns-by-row-script/raw/main/sort-columns-by-row9-example.gif)



## License

This project is licensed under the MIT License.
