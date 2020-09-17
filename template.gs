var settingsArray = getSettings()

// createTemplate(): Creates templates for each sport type

function createTemplate(name, templateNum) {
  
  // Define global variables and creates a spreadsheet
  var spread = SpreadsheetApp.openById(settingsArray[6][0])
  var sheet = spread.insertSheet(name, spread.getSheets().length)
  
  // Define text
  var titles = [
    [[]],
    [["Sport", "Team", "Away Team", "Score", "Home Team", "Score", "Date", "Start Time", "End Time", "Location", "Notes", "Make Up", "Make Sc", "Skip Up", "Skip Sc"]],
    [["Sport", "Type", "Event Name", "Team Name", "Individual Name", "Category", "Place", "Time", "Date", "Event Start Time", "Event End Time", "Event Location", "Post Type", "Make", "Skip"]],
    [["Sport", "Team", "Title", "Message", "Date", "Make", "Skip"]]
  ]
  
  // Set spacing and text
  sheet.getRange(1, 1, 1, titles[templateNum][0].length).setValues(titles[templateNum]).setFontWeight("bold")
  sheet.getRange(1, 1, 1000, 26).setFontFamily("Barlow")
  sheet.setFrozenRows(1)
  
  // Define widths
  var widths = [
    [],
    [0, 90, 135, 150, 50, 150, 50, 150, 80, 80, 300, 200, 50, 50, 50, 50],
    [0, 90, 150, 150, 150, 150, 150, 150, 80, 150, 100, 100, 300, 100, 50, 50],
    [0, 90, 150, 150, 150, 150, 50, 50]
  ]
  
  // Set widths
  for (var i = 1; i < widths[templateNum].length; i++) {
    sheet.setColumnWidth(i, widths[templateNum][i]) 
  }
  
  // Define formatting
  var settings = [
    [],
    [12, 4, 7], // Column of checkbox, Number of columns of checkbox, Column of date
    [14, 2, 9], 
    [6, 2, 5]
  ]
  
  // Set formatting
  sheet.getRange(2, settings[templateNum][0], 999, settings[templateNum][1]).setDataValidation(SpreadsheetApp.newDataValidation().requireCheckbox())
  sheet.getRange(2, settings[templateNum][2], 999, 1).setNumberFormat("ddd mmm d, yyyy")
  sheet.getRange(1, 26).setValue(templateNum)
  if (templateNum == 2) {
    sheet.getRange(2, 13, 999, 1).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(["Upcoming", "Results"], true))
  }
  
  return sheet
  
}
