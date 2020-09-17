var settingsArray = getSettings()

// onOpen(): Creates menu for GUI use.

function onOpen() {
  SpreadsheetApp.getUi()
  .createMenu('Athletics')
  .addItem('Update Scores', 'saveData')
  .addItem('Create Posts', 'editSlide')
  .addItem('Export Posts', 'savePhoto')
  
  .addSeparator()
  
  .addItem("Update Status", "updateStatus")
  
  .addSeparator()
  
  .addItem("Clear Marked", "clearMarked")
  .addItem("Clear Posts", "clearPosts")
  .addItem("Clear Old News", "autoClearNews")
  
  .addSeparator()
  
  .addItem("Update Logo", "changeLogo")
  .addItem("Update Colours", "changeColor")
  .addItem("Reset Auto Sheets", "resetAutoSheets")
  .addItem("Reset All Sheets", "resetSheets")
   
  .addToUi();
}

// updateStatus(): Updates status indicator on overview page.

function updateStatus() {
  
  var sheet = SpreadsheetApp.openById(settingsArray[6][0])
  var slide = SlidesApp.openById(settingsArray[6][2]).getSlides()
  var leng = slide.length
  var sheetArray = sheet.getSheets()
  var array
  var row 
  var mark = 0
  
  for (var i = 2; i < sheet.getSheets().length; i++) {
    row = 2
    while (sheetArray[i].getRange(row, 3).isBlank() == false) {
      row++
    }
    
    array = sheetArray[i].getRange(2, 1, row, 15).getDisplayValues()
    
    for (var x = 0; x < array.length; x++) {
      
      if (sheetArray[i].getRange(1, 26).getDisplayValue() == "1") {
        
        if (array[x][11] == "TRUE" || array[x][12] == "TRUE") {
          mark++
        }
      } else if (sheetArray[i].getRange(1, 26).getDisplayValue() == "2") {
        if (array[x][13] == "TRUE") {
          mark++
        }
      } else if (sheetArray[i].getRange(1, 26).getDisplayValue() == "3") {
        if (array[x][5] == "TRUE") {
          mark++
        }
      }
    }
  }
    sheet.getSheets()[0].getRange(5, 6).setValue(parseInt(mark)) // LINKED TO SHEET
    sheet.getSheets()[0].getRange(9, 6).setValue(leng) // LINKED TO SHEET
  
}


function addToUpdate(num, add) {
  
  var array = [5, 9] // LINKED TO SHEET
  var range = SpreadsheetApp.openById(settingsArray[6][0]).getSheets()[0].getRange(array[num], 6)
  range.setValue(add + parseInt(range.getDisplayValue()))

}

function resetUpdate(num) {
  var array = [5, 9] // LINKED TO SHEET
  var sheet = SpreadsheetApp.openById(settingsArray[6][0]).getSheets()[0].getRange(array[num], 6).setValue(0)
}
