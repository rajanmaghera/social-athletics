var settingsArray = getSettings()

// checkSettings(): Verifies settings on the main page and that they are correct and that they aren't empty.
// Return true = good to go; return false = missing

function checkSettings() {
  var ui = SpreadsheetApp.getUi() // Grab UI for alerts
  
  if (settingsArray[2][0] == "TRUE" || settingsArray[8][0] == "TRUE") { // If email is checked
    if (settingsArray[2][1] == "" || settingsArray[2][2] == "" || settingsArray[2][3] == "" || settingsArray[6][2] == "" || settingsArray[6][3] == "") { // If slots are empty
      var result = ui.alert( // Alert
        'Error',
        'Email settings are incomplete. Please verify email settings and links.',
        ui.ButtonSet.OK
      )  
      return false
    }
  }
  
  if (settingsArray[3][0] == "TRUE") { // If Twitter is checked
    if (settingsArray[3][1] == "" ||  settingsArray[3][2] == "" ||  settingsArray[3][3] == "" ||  settingsArray[3][4] == "") {
      var result = ui.alert(
        'Error',
        'Twitter settings are incomplete. Please verify Twitter settings.',
        ui.ButtonSet.OK
      )  
      return false      
    }
  }
  
  if (settingsArray[4][0] == "TRUE") { // If Slides is checked
    if (settingsArray[4][1] == "" || settingsArray[4][2] == "") {
      var result = ui.alert(
        'Error',
        'Slides settings are incomplete. Please verify Slides settings and links.',
        ui.ButtonSet.OK
      )  
      return false    
    }
  }
  
  if (settingsArray[5][0] == "TRUE") { // If image is checked
    if (settingsArray[5][1] == "" || settingsArray[6][2] == "" || settingsArray[6][3] == "") {
      var result = ui.alert(
        'Error',
        'Image settings are incomplete. Please verify image settings and links.',
        ui.ButtonSet.OK
      )  
      return false    
    }
  }
  
  
  return true
}

// changeLogo(): Replaces logo on template with "Logo" file in "Icons" folder.

function changeLogo() {
  
  var slide = SlidesApp.openById(settingsArray[6][1]).getSlides()
  var folder = DriveApp.getFolderById(settingsArray[6][3])
  var image
  
  if (folder.getFilesByName("Logo").hasNext()) {
    image = folder.getFilesByName("Logo").next() 
  } else {
    image = folder.getFilesByName("Blank").next() 
  }
  
  for (var i = 0; i < slide.length; i++) {
    slide[i].getImages()[2].replace(image)
  }
  
}

// changeColor(): Edits colors of template with values in spreadsheet.

function changeColor() {
  
  
  var hex1 = settingsArray[0][0]
  var hex2 = settingsArray[0][1]
  
  var textArray1 = [
    [5, 7],
    [7, 10],
    [5],
    [7],
    []
  ]
  
  var textArray2 = [
    [6],
    [5, 8],
    [],
    [5],
    [5, 6]
  ]
  
  var slide
  var slideNum = [6,4]
  for (var x = 0; x < slideNum.length; x++) {
    
    slide = SlidesApp.openById(settingsArray[slideNum[x]][1]).getSlides()
    for (var s = 0; s < slide.length; s++) {   
      try {
      slide[s].getShapes()[2].getFill().setSolidFill(hex1)
      } catch(error) {
        
      }
      for (var i = 0; i < textArray1[s].length; i++) {
        slide[s].getShapes()[textArray1[s][i]].getText().getTextStyle().setForegroundColor(hex1)
      }
      
      for (var i = 0; i < textArray2[s].length; i++) {
        slide[s].getShapes()[textArray2[s][i]].getText().getTextStyle().setForegroundColor(hex2)
      }
    }
  }
}

// resetAutoSheets(): Deletes all automatic sheets

function resetAutoSheets() {
  
  var ui = SpreadsheetApp.getUi()
  
  var spread = SpreadsheetApp.openById(settingsArray[6][0])
  var delCur = 2
  var final = 2
  
  var result = ui.alert(
    'Reset Auto Sheets',
    'The following will clear all automated sheets. Are you sure you would like continue?',
    ui.ButtonSet.YES_NO
  )
  
  if (result == ui.Button.YES){

    for (var i = spread.getSheets().length - 1; i > 1; i--){ 
      if (spread.getSheets()[i].getRange(1, 26).getDisplayValue() == "1" && spread.getSheets()[i].getRange(1, 25).getDisplayValue() != "1") {
        spread.deleteSheet(spread.getSheets()[i])
      }
    }
  }
}

// resetSheets(): Deletes all sheets

function resetSheets() { 
  var ui = SpreadsheetApp.getUi()
  var spread = SpreadsheetApp.openById(settingsArray[6][0])
  
  var result = ui.alert(
    'Reset Sheets',
    'The following will clear ALL sheets including custom scores. Are you sure you would like continue?',
    ui.ButtonSet.YES_NO
  )
  
  if (result == ui.Button.YES) {
    
    for (var i = spread.getSheets().length - 1; i > 1; i--){ 
      spread.deleteSheet(spread.getSheets()[i])
    }
  }
}

// clearMarked(): Clear all marked checkmarks

function clearMarked() {
  var ui = SpreadsheetApp.getUi()
  var spread = SpreadsheetApp.openById(settingsArray[6][0])
  var templateNum
  var array = [
    [],
    [12, 2],
    [14, 1],
    [6, 1]
  ]
  
  var result = ui.alert(
    'Clear Marked',
    'The following will clear all marked posts in this spreadsheet. Are you sure you would like to continue?',
    ui.ButtonSet.YES_NO
  )
  
  if (result == ui.Button.YES && spread.getSheets().length > 2) {
    for (var i = 2; i < spread.getSheets().length; i++) {
      
      templateNum = parseInt(spread.getSheets()[i].getRange(1, 26).getDisplayValue())
      spread.getSheets()[i].getRange(2, array[templateNum][0], 999, array[templateNum][1]).setValue("FALSE")
      
    }
  }
  
  resetUpdate(0)
  
}

// clearPosts(): Deletes all posts made for exporting 

function clearPosts() {
  
  var ui = SpreadsheetApp.getUi()
  var slide = SlidesApp.openById(settingsArray[6][2])
  
  var result = ui.alert(
    'Clear Posts',
    'The following will clear all created posts ready for export. Are you sure you would like to continue?',
    ui.ButtonSet.YES_NO
  )
  
  if (result == ui.Button.YES && slide.getSlides().length > 0) {
    for (var i = slide.getSlides().length - 1; i > -1; i--) {
      
      slide.getSlides()[i].remove()
      
    }
  }
  
  resetUpdate(1)
}

// autoClearNews(): Deletes all old and expired news

function autoClearNews() {
  const MILLS = 1000 * 60 * 60 * 24
  var slide = SlidesApp.openById(settingsArray[4][2]).getSlides()
  var info
  for (var i = slide.length-1; i > -1; i--) {
    info = slide[i].getShapes()[0].getTitle()
    if (

      info.substring(0, 1) == "U" && (new Date(info.substring(1))).getTime() < (new Date()).getTime() - MILLS + MILLS/24*6 ||
      info.substring(0, 1) == "R" && (new Date(info.substring(1))).getTime() < (new Date()).getTime() - MILLS*3 ||
        info.substring(0, 1) == "A" && (new Date(info.substring(1))).getTime() < (new Date()).getTime() - MILLS*5
        )
        {
          slide[i].remove()
        }
    
    
    
  }
  
}

// getSettings(): Returns values from overview sheet into 2D array

function getSettings() {
  

  var sheetArray = SpreadsheetApp.getActive().getSheets()[0].getRange(1, 1, 42, 10).getDisplayValues()
  
  var emailMarked, twitterMarked, slidesMarked, imageMarked
  
  var color1 = sheetArray[14][2]
  var color2 = sheetArray[15][2]
  
  
  var autoMarkScore = sheetArray[18][2]
  var autoMarkUpcoming = sheetArray[19][2]
  var autoMarkThreshold = sheetArray[20][2]
  var autoMarkTime = sheetArray[21][2]
  
  var emailSwitch = sheetArray[25][2]
  var emailRecipient = sheetArray[27][2]
  var emailTitle = sheetArray[28][2]
  var emailBody = sheetArray[29][2]

  
  if (emailSwitch == "TRUE"&& emailRecipient != "" && emailTitle != "" && emailBody != "") {
    emailMarked = true
  } else {
    emailMarked = false
  }
  
  var twitterSwitch = sheetArray[25][5]
  var twitterConsumer = sheetArray[27][5]
  var twitterConsumerSecret = sheetArray[28][5]
  var twitterAccess = sheetArray[30][5]
  var twitterAccessSecret = sheetArray[31][5]
  
  if (twitterSwitch == "TRUE"&& twitterConsumer != "" && twitterConsumerSecret != "" && twitterAccess != "" && twitterAccessSecret != "") {
    twitterMarked = true
  } else {
    twitterMarked = false
  }
  
  var slidesSwitch = sheetArray[25][8]
  var slidesTemplateId = sheetArray[40][2].substring(sheetArray[40][2].indexOf("/d/") + 3, sheetArray[40][2].indexOf("/edit"))
  var slidesId = sheetArray[39][2].substring(sheetArray[39][2].indexOf("/d/") + 3, sheetArray[39][2].indexOf("/edit"))
  var slidesAutoClear = sheetArray[27][8]
  
  if (slidesSwitch == "TRUE"&& slidesTemplateId != "" && slidesId != "") {
    slidesMarked = true
  } else {
    slidesMarked = false
  }
  
  var imageSwitch = sheetArray[29][8]
  var imageFolderId = sheetArray[41][2].substring(sheetArray[41][2].indexOf("/folders/") + 9, sheetArray[41][2].length)
  
  if (imageSwitch == "TRUE"&& imageFolderId != "") {
    imageMarked = true
  } else {
    imageMarked = false
  }

  var spreadsheetId = SpreadsheetApp.getActive().getId()
  var templateId = sheetArray[35][2].substring(sheetArray[35][2].indexOf("/d/") + 3, sheetArray[35][2].indexOf("/edit"))
  var postId = sheetArray[36][2].substring(sheetArray[36][2].indexOf("/d/") + 3, sheetArray[36][2].indexOf("/edit"))
  var iconId = sheetArray[37][2].substring(sheetArray[37][2].indexOf("/folders/") + 9, sheetArray[37][2].length)
  var backgroundId = sheetArray[38][2].substring(sheetArray[38][2].indexOf("/folders/") + 9, sheetArray[38][2].length)
  
  var masterSwitch = sheetArray[12][9]
  
  var emailTextSwitch = sheetArray[31][8]
  
  var settingsArray = [
    [color1, color2],
    [autoMarkScore, autoMarkUpcoming, autoMarkThreshold, autoMarkTime],
    [emailSwitch, emailRecipient, emailTitle, emailBody],
    [twitterSwitch, twitterConsumer, twitterConsumerSecret, twitterAccess, twitterAccessSecret],
    [slidesSwitch, slidesTemplateId, slidesId, slidesAutoClear],
    [imageSwitch, imageFolderId],
    [spreadsheetId, templateId, postId, iconId, backgroundId],
    [masterSwitch],
    [emailTextSwitch]
  ]
  
  return settingsArray
}
