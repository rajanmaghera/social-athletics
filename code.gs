var settingsArray = getSettings()

//tweetEmail(post): Sends contents of tweet to email

function tweetEmail(post) {
  
  MailApp.sendEmail(settingsArray[2][1], 
                    settingsArray[2][2] + " (" + new Date().toString().substring(0, 24) + ")", 
    post,
      {
        name: 'Automated',
      }) 
}

// savePhoto(): Turns spreadsheets from post slideshow into images saved in the output Google Drive folder and sends an email (if specified).

function savePhoto() {
  
  if (checkSettings()) {
    
    // Setup global variables
    var sheet = SpreadsheetApp.openById(settingsArray[6][0]).getSheets()[0]
    var slide = SlidesApp.openById(settingsArray[6][2]).getSlides()
    var folder = DriveApp.getFolderById(settingsArray[5][1])
    var attachArray = []
    var url, response, options
    
    if (slide.length > 0 && (settingsArray[2][0] == "TRUE" || settingsArray[5][0] == "TRUE")) { // Only run code if slides exist in the post slideshow
      
      if (settingsArray[5][0] == "TRUE") {
        folder = folder.createFolder(Utilities.formatDate(new Date(), "America/Edmonton", "yyyy-MM-dd HH:mm:ss")) // Create dated folder
      }
      
      for (var i = slide.length - 1; i > -1 ; i--) { // For each slide:
        
        // Save an image for each slide
        url = 'https://docs.google.com/presentation/d/' + settingsArray[6][2] + '/export/png?id=' + settingsArray[6][2] + '&pageid=' + slide[i].getObjectId() // Determine the URL to the export link
        options = { // Setup OAuth
          headers: {
            Authorization: 'Bearer ' + ScriptApp.getOAuthToken()
          }
        }
        response = UrlFetchApp.fetch(url, options).getAs(MimeType.PNG).setName(" Image " + (i+1)) // Grab image from URL
        
        if (settingsArray[5][0] == "TRUE") {
          
          folder.createFile(response) // Create image in folder
        }
        
        attachArray.push(response) // Add image as an attachment
        
        slide[i].remove() // Remove slide 
        
      }
      
      if (settingsArray[2][0] == "TRUE") { //If email is wanted: 
        
        // Send an email with all the images
        MailApp.sendEmail(settingsArray[2][1], 
                          settingsArray[2][2] + " (" + new Date().toString().substring(0, 24) + ")", 
          settingsArray[2][3],
            {
              name: 'Automated',
                attachments: attachArray
            })   
      }
    }
    
    resetUpdate(1)
    
  }
}

// setDay(date, slide, templateNum): Edits date widget using "date" inupt on "slide" based on "templateNum."
// If date is impossible, the widget will disappear.

function setDay(date, slide, templateNum) {
  
  // Setup global variables
  var objects = slide.getShapes()
  var calNum = [13, 13, 11, 11, 9]
  const MILLS = 1000 * 60 * 60 * 24 // The number of millisecconds in a day; used for calculations
  
  // Edit numbers
  if ((new Date(date)).getTime() > 100) { // If date exists:
    for (var i = 0; i < 7; i++) { // For each number of calendar dates on template (1-7):
      slide.getShapes()[calNum[templateNum] + i].getText().setText((new Date((new Date((new Date(date)).getTime() - (new Date(date).getDay()*MILLS))).getTime() + i*MILLS)).getDate()) // Set text to day in accordance to weekday
    }
    
    slide.getShapes()[calNum[templateNum] + (new Date(date)).getDay()].getFill().setSolidFill(settingsArray[0][1]) // Fill date day with background color
    slide.getShapes()[calNum[templateNum] + (new Date(date)).getDay()].getText().getTextStyle().setForegroundColor(255, 255, 255) // Set date day text to white
    
    // Delete widget if date does not work
  } else {
    for (var i = 0; i < 8; i++) { // For each date widget number plus day labels:
      slide.getShapes()[calNum[templateNum] - 1 + i].getText().setText("") // Set text to blank
    }
  }
}

// setImages(slide, icon, background): Edit images on "slide" in accordance to "icon" name (eg. Rugby) and "background" name (eg. Sr Boys Rugby)

function setImages(slide, icon, background) {
  
  
  // Setup global variables
  var iconFolder = DriveApp.getFolderById(settingsArray[6][3]).getFilesByName(icon)
  var backgroundFolder = DriveApp.getFolderById(settingsArray[6][4])
  var backgroundImages = []
  var ui = SpreadsheetApp.getUi() // Grab UI for alerts
  
  // Edit sport icon (bottom image)
  if (iconFolder.hasNext()) { // If icon image exists:
    slide.getImages()[1].replace(iconFolder.next()) // Replace placeholder with icon
  } else { // Else:
    slide.getImages()[1].remove() // Remove image
  }
  
  
  // Edit background image
  if (backgroundFolder.getFoldersByName(background).hasNext()) { // If a folder for the team exists:
    backgroundFolder = backgroundFolder.getFoldersByName(background).next().getFiles() // Get files from folder
    if (backgroundFolder.hasNext()) { // If images in folder exist:
      while (backgroundFolder.hasNext()) { // While there are images:
        backgroundImages.push(backgroundFolder.next()) // Add image to an intermediary array
      }
      slide.getImages()[0].replace(backgroundImages[Math.floor(Math.random() * Math.floor(backgroundImages.length))], true) // Set background to random image from array
    }
  } else { // If folder for the team doesn't exist
    backgroundFolder = backgroundFolder.createFolder(background).getFiles() // Create a folder
  }
  
}

// editSlide(): Turn each of the marked events in the spreadsheet into a slide for exporting.

function editSlide() {
  
  if (checkSettings()){
    
    // Setup global variables
    var template
    var sheetForLength = [7, 7]
    var sheetObjects = [
      [0, 1, 6, 2, 4, 7, 9],
      [0, 1, 6, 2, 4, 3, 5]
    ]
    var templateObjects = [ // Object Index
      [11, 4, 9, 5, 7, 8, 10], // Sport, Team Name, Date, Away Team, Home Team, Start Time, Location
      [11, 4, 5, 7, 10, 6, 9], // Sport, Team Name, Date, Away Team, Home Team, Away Score, Home Score
      [9, 4, 5, 7, 6, 8], // Sport, Sport + Type , Event Name, Date, Start Time, Location
      [9, 3, 4, 5, 6, 7, 8], // Sport, Type, Event Name, Date, Headline, Mainline, Subline
      [7, 4, 5, 6] // Sport, Title, Date, Message
    ]
    
    // NEW OBJECT INDEX (To be implemented)
    // This is for use to simplify the type 2 sport edit module
    
    var ss = SpreadsheetApp.openById(settingsArray[6][0])
    var row, col, sheet, templateNum, slide, sheetArray
    var updater = 0
    var post, postA
    var objectA, objectB, objectC
    var templatePicker = []
    
    for (var sheetNum = 2; sheetNum < ss.getSheets().length; sheetNum++) { // For each sheet (sport) in document:
      
      sheet = ss.getSheets()[sheetNum]
      row = 2
      
      // Run code for type 1 sport
      if (sheet.getRange(1, 26).getDisplayValue() == "1") {
        while (sheet.getRange(row, 3).isBlank() == false) { // While an event exists:
          row++
        }
        
        sheetArray = sheet.getRange(2, 1, row-2, 15).getDisplayValues()
        
        for (var arrayRow = 0; arrayRow < sheetArray.length; arrayRow++) {
          
          if (sheetArray[arrayRow][11] == "TRUE" || sheetArray[arrayRow][12] == "TRUE") { // If it is marked:
            
            
            // Determine type of template
            if (sheetArray[arrayRow][12] == "TRUE") {
              templateNum =  1
              var temp = "R"
              } else {
                templateNum = 0
                var temp = "U"
                }
            
            
            if (templateNum == 1 && settingsArray[3][0] == "TRUE" || templateNum == 1 && settingsArray[8][0] == "TRUE") {
              post = ( 
                "SCORE :: " +
                sheetArray[arrayRow][sheetObjects[templateNum][1]] + ": " +
                sheetArray[arrayRow][sheetObjects[templateNum][3]] + " [" +
                sheetArray[arrayRow][sheetObjects[templateNum][5]] + "] @ " +
                sheetArray[arrayRow][sheetObjects[templateNum][4]] + " [" +
                sheetArray[arrayRow][sheetObjects[templateNum][6]] + "] :: " +
                sheetArray[arrayRow][sheetObjects[templateNum][2]]
                
              )
              if (settingsArray[3][0] == "TRUE") {
                postTweet(post)
              } 
              
              if (settingsArray[8][0] == "TRUE") {
                tweetEmail(post)              
              }
              
              Utilities.sleep(5000)
            }
            
            if (templateNum == 0 && settingsArray[3][0] == "TRUE" || templateNum == 0 && settingsArray[8][0] == "TRUE") {
              post = (  
                "UPCOMING :: " +
                sheetArray[arrayRow][sheetObjects[templateNum][1]] + ": " +
                sheetArray[arrayRow][sheetObjects[templateNum][3]] + " @ " +
                sheetArray[arrayRow][sheetObjects[templateNum][4]] + " :: " +
                sheetArray[arrayRow][sheetObjects[templateNum][2]] + " " +
                sheetArray[arrayRow][sheetObjects[templateNum][5]] + " @ " +
                sheetArray[arrayRow][sheetObjects[templateNum][6]]
                
              )
              
              if (settingsArray[3][0] == "TRUE") {
                postTweet(post)
              } 
              
              if (settingsArray[8][0] == "TRUE") {
                tweetEmail(post)              
              }
              
              Utilities.sleep(5000)
            }
            
            if (settingsArray[2][0] == "TRUE" || settingsArray[5][0] == "TRUE" || settingsArray[4][0] == "TRUE") {
              
              
              if (settingsArray[2][0] == "TRUE" || settingsArray[5][0] == "TRUE") {
                templatePicker.push(6)
                updater++
              }
              
              if (settingsArray[4][0] == "TRUE") {
                templatePicker.push(4) 
              }
              
              
              
              for (var x = 0; x < templatePicker.length; x++){
                
                template = SlidesApp.openById(settingsArray[templatePicker[x]][1]).getSlides()
                slide = SlidesApp.openById(settingsArray[templatePicker[x]][2]).appendSlide(template[templateNum]) // Create an empty template slide
                slide.getShapes()[0].setTitle(temp + sheetArray[arrayRow][sheetObjects[templateNum][2]]) // Set Date in title for automatic removal of tiger news
                
                // Edit slides in accodance to templateNum
                for (var i = 0; i < sheetForLength[templateNum]; i++) {
                  slide.getShapes()[templateObjects[templateNum][i]].getText().setText(sheetArray[arrayRow][sheetObjects[templateNum][i]])
                  if (i < 3) {
                    slide.getShapes()[templateObjects[templateNum][i]].getText().setText(slide.getShapes()[templateObjects[templateNum][i]].getText().asString().toUpperCase())
                  }
                }
                
                setDay(slide.getShapes()[templateObjects[templateNum][2]].getText().asString().substring(4), slide, templateNum) // Setup date widget
                setImages(slide, sheetArray[arrayRow][sheetObjects[templateNum][0]], sheetArray[arrayRow][sheetObjects[templateNum][1]])
                
              }
            }
            
            sheetArray[arrayRow][templateNum+13] = "TRUE" // Mark skipped field
            sheetArray[arrayRow][11] = "FALSE" // Reset mark field
            sheetArray[arrayRow][12] = "FALSE" // Reset mark field
            
            templatePicker = []
            
          }
        }
        
        sheet.getRange(2, 1, row-2, 15).setValues(sheetArray)
        
        // Run code for type 2 sport
      } else if (sheet.getRange(1, 26).getDisplayValue() == "2") {
        
        while (sheet.getRange(row, 3).isBlank() == false) { // While an event exists:
          row++
        }
        
        if (row > 2) {
          sheetArray = sheet.getRange(2, 1, row-2, 15).getDisplayValues()
          
          for (var arrayRow = 0; arrayRow < sheetArray.length; arrayRow++) {
            
            if (sheetArray[arrayRow][13] == "TRUE") { // If it is marked:
              
              
              
              // Determine templateNum based on dropdown
              if (sheetArray[arrayRow][12] == "Upcoming") { 
                var templateNum = 2
                } else if (sheetArray[arrayRow][12] == "Results") {
                  var templateNum = 3
                  }
              
              
              if (settingsArray[2][0] == "TRUE" || settingsArray[5][0] == "TRUE" || settingsArray[4][0] == "TRUE") {
                
                if (settingsArray[2][0] == "TRUE" || settingsArray[5][0] == "TRUE") {
                  templatePicker.push(6) 
                  updater++
                } 
                
                if (settingsArray[4][0] == "TRUE") {
                  templatePicker.push(4) 
                }
                
                
              }
              
              // Edit slides in accodance to templateNum
              
              if (templateNum == 2) {
                
                
                if (settingsArray[2][0] == "TRUE" || settingsArray[5][0] == "TRUE" || settingsArray[4][0] == "TRUE") { // If an image is required
                  
                  for (var x = 0; x < templatePicker.length; x++) {
                    template = SlidesApp.openById(settingsArray[templatePicker[x]][1]).getSlides()
                    slide = SlidesApp.openById(settingsArray[templatePicker[x]][2]).appendSlide(template[templateNum]) // Create an empty template slide
                    slide.getShapes()[0].setTitle("U" + sheetArray[arrayRow][9])
                    slide.getShapes()[templateObjects[templateNum][1]].getText().setText((sheetArray[arrayRow][0] + " " + sheetArray[arrayRow][1]).toUpperCase()) // Sport + Type
                    slide.getShapes()[templateObjects[templateNum][4]].getText().setText(sheetArray[arrayRow][9].toUpperCase()) // Start Time
                    slide.getShapes()[templateObjects[templateNum][5]].getText().setText(sheetArray[arrayRow][11]) // Location 
                    
                    slide.getShapes()[templateObjects[templateNum][0]].getText().setText(sheetArray[arrayRow][0].toUpperCase()) // Sport
                    slide.getShapes()[templateObjects[templateNum][2]].getText().setText(sheetArray[arrayRow][2].toUpperCase()) // Event Name
                    slide.getShapes()[templateObjects[templateNum][3]].getText().setText(sheetArray[arrayRow][8].toUpperCase()) // Date
                    
                    setDay(slide.getShapes()[templateObjects[templateNum][3]].getText().asString(), slide, templateNum)
                    setImages(slide, sheetArray[arrayRow][0], sheetArray[arrayRow][0])
                  }
                }
                
                if (settingsArray[3][0] == "TRUE" || settingsArray[8][0] == "TRUE") { // If Twitter is enabled
                  post = (  
                    "UPCOMING :: " +
                    sheetArray[arrayRow][0] + " " +
                    sheetArray[arrayRow][1] + ": " +
                    sheetArray[arrayRow][2] + " :: " +
                    sheetArray[arrayRow][8] + " " +
                    sheetArray[arrayRow][9] + " @ " +
                    sheetArray[arrayRow][11] 
                  )
                  if (settingsArray[3][0] == "TRUE") {
                    postTweet(post)
                  } 
                  
                  if (settingsArray[8][0] == "TRUE") {
                    tweetEmail(post)              
                  }
                  Utilities.sleep(5000)
                }
                
              } else if (templateNum == 3)  {
                
                objectC = sheetArray[arrayRow][5] // Bottom: Category
                postA = sheetArray[arrayRow][5]
                
                if (sheetArray[arrayRow][4] == "") { // Team Results
                  objectA = sheetArray[arrayRow][3] // Top: Team Name
                  postA = postA + " - " + sheetArray[arrayRow][3]
                  
                  if (sheetArray[arrayRow][6] != "") { // PLACE (1) exists
                    objectB = sheetArray[arrayRow][6] // Middle: Place
                    postA = postA + " [" + sheetArray[arrayRow][6]
                    
                    if (sheetArray[arrayRow][7] != "") { // TIME (2) exists
                      objectB = objectB + " - " + sheetArray[arrayRow][7] // MIDDLE: Time
                      postA = postA + " - " + sheetArray[arrayRow][7]
                    }
                    postA = postA + "]"
                    
                  } else if (sheetArray[arrayRow][7] != "") { // TIME (1) exists
                    objectB = sheetArray[arrayRow][7] // Middle: Time 
                    postA = postA + " [" + sheetArray[arrayRow][7] + "]"
                    
                  } else { // Only CATEGORY exists
                    
                    objectB = sheetArray[arrayRow][5] // Middle: Category
                    objectC = "" // (Remove Category from bottom)
                  }
                  
                } else { // Individual Results
                  objectB = sheetArray[arrayRow][4] // Middle: Individual Name
                  postA = postA + " - " + sheetArray[arrayRow][4]
                  
                  if (sheetArray[arrayRow][7] != "") { // PLACE (1) exists
                    objectA = sheetArray[arrayRow][6] // Top: Place
                    postA = postA + " [" + sheetArray[arrayRow][6]
                    
                    if (sheetArray[arrayRow][7] != "") { // TIME (2) exists
                      objectA = objectA + " - " + sheetArray[arrayRow][7] // Top: Time
                      postA = postA + " - " + sheetArray[arrayRow][7]
                    }
                    
                    postA = postA + "]"
                    
                  } else if (sheetArray[arrayRow][7] != "") { // TIME (1) exists
                    objectA = sheetArray[arrayRow][7] // Top: Time
                    postA = postA + " [" + sheetArray[arrayRow][7] + "]"
                    
                  } else { // CATEGORY exists
                    
                    objectA = sheetArray[arrayRow][5] // Top: Category
                    objectC = "" // (Remove Category from bottom)
                  }
                }
                
                if (settingsArray[3][0] == "TRUE" || settingsArray[8][0] == "TRUE") {
                  post = (  
                    sheetArray[arrayRow][1].toUpperCase() + " :: " +
                    sheetArray[arrayRow][0] + ": " +
                    sheetArray[arrayRow][2] + " :: " +
                    postA + " :: " +
                    sheetArray[arrayRow][8]
                  )
                  if (settingsArray[3][0] == "TRUE") {
                    postTweet(post)
                  } 
                  
                  if (settingsArray[8][0] == "TRUE") {
                    tweetEmail(post)              
                  }
                  Utilities.sleep(5000)
                }
                
                if (settingsArray[2][0] == "TRUE" || settingsArray[5][0] == "TRUE" || settingsArray[4][0] == "TRUE") {
                  for (var x = 0; x < templatePicker.length; x++) {
                    template = SlidesApp.openById(settingsArray[templatePicker[x]][1]).getSlides()
                    slide = SlidesApp.openById(settingsArray[templatePicker[x]][2]).appendSlide(template[templateNum]) // Create an empty template slide
                    slide.getShapes()[0].setTitle("R" + sheetArray[arrayRow][8])
                    
                    slide.getShapes()[templateObjects[templateNum][1]].getText().setText(sheetArray[arrayRow][1].toUpperCase()) // Type
                    slide.getShapes()[templateObjects[templateNum][4]].getText().setText(objectA) // Top
                    slide.getShapes()[templateObjects[templateNum][5]].getText().setText(objectB) // Middle
                    slide.getShapes()[templateObjects[templateNum][6]].getText().setText(objectC) // Bottom
                    
                    slide.getShapes()[templateObjects[templateNum][0]].getText().setText(sheetArray[arrayRow][0].toUpperCase()) // Sport
                    slide.getShapes()[templateObjects[templateNum][2]].getText().setText(sheetArray[arrayRow][2].toUpperCase()) // Event Name
                    slide.getShapes()[templateObjects[templateNum][3]].getText().setText(sheetArray[arrayRow][8].toUpperCase()) // Date
                    
                    setDay(slide.getShapes()[templateObjects[templateNum][3]].getText().asString(), slide, templateNum)
                    setImages(slide, sheetArray[arrayRow][0], sheetArray[arrayRow][0])
                  }
                }
              }
              
              
              
              
              sheetArray[arrayRow][14] = "TRUE" // Mark skipped field
              sheetArray[arrayRow][13] = "FALSE" // Reset mark field
              
              templatePicker = []
            }
          }
          sheet.getRange(2, 1, row-2, 15).setValues(sheetArray);
        }
        // Runcode for type 3 sport (Announcements)
      } else if (sheet.getRange(1, 26).getDisplayValue() == "3") {
        
        while (sheet.getRange(row, 4).isBlank() == false) { // While an event exists:
          row++
        }
        if (row > 2) {
          sheetArray = sheet.getRange(2, 1, row-2, 7).getDisplayValues()
          for (var arrayRow = 0; arrayRow < sheetArray.length; arrayRow++) {
            if (sheetArray[arrayRow][5] == "TRUE") {
              
              templateNum = 4 
              if (settingsArray[2][0] == "TRUE" || settingsArray[5][0] == "TRUE" || settingsArray[4][0] == "TRUE") {
                
                if (settingsArray[2][0] == "TRUE" || settingsArray[5][0] == "TRUE") {
                  templatePicker.push(6) 
                  updater++
                } 
                
                if (settingsArray[4][0] == "TRUE") {
                  templatePicker.push(4) 
                }
                
                
                
                for (var x = 0; x < templatePicker.length; x++) {
                  template = SlidesApp.openById(settingsArray[templatePicker[x]][1]).getSlides()
                  slide = SlidesApp.openById(settingsArray[templatePicker[x]][2]).appendSlide(template[templateNum])
                  slide.getShapes()[0].setTitle("A" + sheetArray[arrayRow][4])
                  
                  slide.getShapes()[templateObjects[templateNum][0]].getText().setText(sheetArray[arrayRow][0].toUpperCase())
                  slide.getShapes()[templateObjects[templateNum][1]].getText().setText(sheetArray[arrayRow][2].toUpperCase())
                  slide.getShapes()[templateObjects[templateNum][2]].getText().setText(sheetArray[arrayRow][4].toUpperCase())
                  slide.getShapes()[templateObjects[templateNum][3]].getText().setText(sheetArray[arrayRow][3])
                  
                  setImages(slide, sheetArray[arrayRow][0], sheetArray[arrayRow][1])
                  setDay(sheetArray[arrayRow][4], slide, templateNum)
                }
                
                templatePicker = []
                
              }
              
              if (settingsArray[3][0] == "TRUE" || settingsArray[8][0] == "TRUE") {
                post = (  
                  "ANNOUNCEMENT :: " +
                  sheetArray[arrayRow][2] + ": " +
                  sheetArray[arrayRow][3] + " :: " +
                  sheetArray[arrayRow][0]
                )
                
                if (sheetArray[arrayRow][4] != "") {
                  post = post + " - " + sheetArray[arrayRow][4]
                }
                
                if (settingsArray[3][0] == "TRUE") {
                  postTweet(post)
                } 
                
                if (settingsArray[8][0] == "TRUE") {
                  tweetEmail(post)              
                }
                Utilities.sleep(5000)
              }
              
              sheetArray[arrayRow][5] = "FALSE"
              sheetArray[arrayRow][6] = "TRUE"
              
              
            }
          }
          sheet.getRange(2, 1, row-2, 7).setValues(sheetArray)
          
        }
      }
    }
    
    addToUpdate(1, updater)
    resetUpdate(0)
    
    
  }
}

// saveData(): This function loops through all sports on the 'sports' sheet in Google Sheets and, using the ID, will create the necessary sheet and folder structure for images, and will scan the websites for score information.

function saveData() {
  
  if (checkSettings()) {
    
    var spread = SpreadsheetApp.openById(settingsArray[6][0]) // Open spreadsheet
    var sports = SpreadsheetApp.openById(settingsArray[6][0]).getSheetByName("Sports") // Open list of sports with ID
    var made = false // To check if the sheet for each individual sheet is made
    var cur = 2 // Current row/sport in sports sheet
    var sheetSport, page, ind, clearRow, text, numLenA, numLenB, row, newIndex, checkRange, endRow, endRowScore, added // Initalize empty variables
    const MILLS = 1000 * 60 * 60 * 24 // The number of millisecconds in a day; used for calculations
    
    var updater = 0 
    var sportArray, sheetArray, sheetArrayRow
    var updateArray = [] // 1D Array for each 
    var updateArrayLocal = [] // Each individual row for sports 
    var dateArrayID = [1, 1, 1, 1, 1, 1, 1, 0, 0, 0, 0, 0] // When getting sports from page, determine whether to add or remove a year to ensure site PHP query doesn't fail
    
    while (sports.getRange(cur, 5).getDisplayValue() != "" ) { // Determine the rows of sports on sports sheet
      cur++  
    }
    
    if (cur > 2) { // If the sports sheet isn't empty
      
      sportArray = sports.getRange(2, 1, cur-2, 5).getDisplayValues() // Get all sports information into a 2D array; this is much faster than getting each individual display value.
      
      for (var sportArrayRow = 0; sportArrayRow < sportArray.length; sportArrayRow++) { // For sport in the list of sports
        
        updateArray = [] // Clear the array 
        
        if (sportArray[sportArrayRow][0] == "TRUE") { // If the sport is enabled; using checkmarks and comparing with boolean value was unstable, henceforth the string was used in its placed
          
          
          // If a sport has no name or ID, skip 
          if (sportArray[sportArrayRow][1] == "" && sportArray[sportArrayRow][2] == "") {
            Logger.log("WARNING: Empty sport is enabled.") 
            
            // If sport is type 3 (Announcements)
          } else if (sportArray[sportArrayRow][4] == "3") {
            
            sportArray[sportArrayRow][3] = "TRUE" // Auto enable manual; type 3 sport cannot be automated
            
            // If sheet is non-existant by name of sport, create a sheet
            if (spread.getSheetByName(sportArray[sportArrayRow][1]) == null) {
              createTemplate(sportArray[sportArrayRow][1], 3) // Create a sheet based on sport type and templated numbers in array
            } 
            
            // If sport is type 2 (Non-Scored)
          } else if (sportArray[sportArrayRow][4] == "2") {
            
            sportArray[sportArrayRow][3] = "TRUE" // Auto enable manual; type 2 sport cannot be automated
            
            // If background image folder is non-existant, create a folder
            if (DriveApp.getFolderById(settingsArray[6][4]).getFoldersByName(sportArray[sportArrayRow][1]).hasNext() == false) {
              DriveApp.getFolderById(settingsArray[6][4]).createFolder(sportArray[sportArrayRow][1]).getFiles()
            }
            
            // If sheet is non-existant by name of sport, create a sheet
            if (spread.getSheetByName(sportArray[sportArrayRow][1]) == null) {
              createTemplate(sportArray[sportArrayRow][1], 2) // Create a sheet based on sport type and templated numbers in array
            } 
            
            // If sport is type 1 (Scored) or an ID exists
          } else if (sportArray[sportArrayRow][4] == "1" || sportArray[sportArrayRow][2] != "") {
            
            sportArray[sportArrayRow][4] = "1" // Set the sport type to 1 (if ID exists but type is blank)
            
            // If sport is automated and doesn't have a name
            if (sportArray[sportArrayRow][1] == "" && sportArray[sportArrayRow][2] != "") { 
              page = UrlFetchApp.fetch("http://metroathletics.ca/team_schedule.php?team_id=" + sportArray[sportArrayRow][2] + "&from_date=" + ((new Date()).getFullYear() - dateArrayID[(new Date()).getMonth()]) + "-08-01&to_date=" + ((new Date()).getFullYear() + 1) + "-01-01").getContentText() // Fetch contents of score page
              ind = page.indexOf('schedule-list') + 1 // Search for name of sport by finding key terms in HTML
              ind = page.indexOf('id="game_', ind) + 1
              ind = page.indexOf('<strong>', ind) + 8
              text = (page.substring(ind, page.indexOf('</strong>', ind))).replace("Senior", "Sr").replace("Junior", "Jr").replace(" Game", "") // Set temp variable to a cleaned up name, removing unnecessary information and HTML tags
              
              // Fix for football names; they use divisions instead of explicitly 'Football'
              if (text == "Miles Division" || text == "Carr Division" || text == "Gilfillan Division" || text == "Bright Division") { 
                sportArray[sportArrayRow][1] = "Sr Football"
              } else {
                sportArray[sportArrayRow][1] = text
              }
            }
            
            // If background image folder is non-existant, create a folder
            if (DriveApp.getFolderById(settingsArray[6][4]).getFoldersByName(sportArray[sportArrayRow][1]).hasNext() == false) {
              DriveApp.getFolderById(settingsArray[6][4]).createFolder(sportArray[sportArrayRow][1]).getFiles()
            }
            
            
            // If sheet is non-existant by name of sport, create a sheet
            if (spread.getSheetByName(sportArray[sportArrayRow][1]) == null) {
              sheetSport = createTemplate(sportArray[sportArrayRow][1], 1) // Create a sheet based on sport type and templated numbers in array file.
              
              // If the sport is manual, store value in sheet corner; this is the simplest way to store information about each sheet between runs
              if (sportArray[sportArrayRow][3] == "TRUE") {
                sheetSport.getRange(1, 25).setValue("1") 
              }
              
              made = true // Set the made value to true for future reference
              
              // If sheet exists, grab sheet contents
            } else {
              
              sheetSport = spread.getSheetByName(sportArray[sportArrayRow][1]) // Grab sheet
              sheetArrayRow = 0 // Array value as to how many events in sheet
              
              // Determine the number of filled rows/events
              while (sheetSport.getRange(sheetArrayRow+1, 3).isBlank() == false) {
                sheetArrayRow++
              }
              
              sheetArray = sheetSport.getRange(2, 1, sheetArrayRow, 15).getDisplayValues() // Grab display values of sheet into a 2D Array
            }
            
            // If sport is automated, get sport details from online and save to spreadsheet
            if (sportArray[sportArrayRow][3] != "TRUE") {
              row = 0 // Setup row
              page = UrlFetchApp.fetch("http://metroathletics.ca/team_schedule.php?team_id=" + sportArray[sportArrayRow][2] + "&from_date=" + ((new Date()).getFullYear() - dateArrayID[(new Date()).getMonth()] - 1) + "-08-01&to_date=" + ((new Date()).getFullYear() + 1) + "-01-01").getContentText() // Fetch text of page
              newIndex = page.indexOf('schedule-list') + 1; // Determine index of general location of code
              
              // While a game on the website exists
              while (page.indexOf('id="game_', newIndex) > 1) {
                
                updateArrayLocal = ["", "", "", "", "", "", "", "", "", "", "", "FALSE", "FALSE", "FALSE", "FALSE"] // Setup blank variables for row array
                
                newIndex = page.indexOf('id="game_', newIndex) + 1 // Determine index of game details
                newIndex = page.indexOf('<strong>', newIndex) + 8
                
                text = (page.substring(newIndex, page.indexOf('</strong>', newIndex))).replace("Senior", "Sr").replace("Junior", "Jr").replace(" Game", "") // Set temp variable to a cleaned up name, removing unnecessary information and HTML tags
                Logger.log(text)
                if (text == "Miles Division" || text == "Carr Division" || text == "Gilfillan Division" || text == "Bright Division") { // Fix for football names; they use divisions instead of explicitly 'Football'
                  updateArrayLocal[0] = "Football"
                  updateArrayLocal[1] = "Sr Football"
                } else { 
                  updateArrayLocal[0] = (page.substring(newIndex, page.indexOf('</strong>', newIndex))).replace("Senior ", "").replace("Junior ", "").replace(" Game", "").replace("Boys ", "").replace("Girls ", "") // Save Sport
                  updateArrayLocal[1] = (page.substring(newIndex, page.indexOf('</strong>', newIndex))).replace("Senior", "Sr").replace("Junior", "Jr").replace(" Game", "") // Save Type
                }
                newIndex = page.indexOf('<td>', newIndex) + 4 // Determine index of score information
                text = page.substring(newIndex, page.indexOf('</td>', newIndex)) // Original Game
                
                // Replace the values of HTML spaces and uncessary charectors
                while (text.indexOf('&nbsp;') > -1) { 
                  text = text.replace('&nbsp;', ' ') }
                while (text.indexOf('<') > -1) {
                  text = text.replace(text.substring(text.indexOf('<'), text.indexOf('>') + 1), '') }
                while (text.indexOf('Summary') > -1) {
                  text = text.replace('Summary', '') }
                
                // Determine the length of the score number, then save into array
                numLenB = 0
                numLenA = 0
                
                while (text.substring(text.indexOf('@') - 2 - numLenA, text.indexOf('@') - numLenA - 1) != ' ') { // Determine length of number A
                  numLenA++; 
                }
                
                while (text.substring(text.length - 1 - numLenB, text.length-numLenB) != ' ') { // Determine length of number B
                  numLenB++; 
                }
                
                updateArrayLocal[2] = (text.substring(0, text.indexOf('@') - numLenA - 2)).replace("\'", "'") // Save Away Team
                updateArrayLocal[4] = (text.substring(text.indexOf('@') + 2, text.length-numLenB-1)).replace("\'", "'") // Save Home Team
                
                if (numLenB != 0 && numLenA != 0) { // If a game has a score
                  updateArrayLocal[3] = text.substring(text.indexOf('@') - numLenA - 1, text.indexOf('@') - 1) // Save Away Score
                  updateArrayLocal[5] = text.substring(text.length-numLenB, text.length) // Save Home Score
                } else { // Else put a message
                  updateArrayLocal[10] = "No Score Available" 
                }
                
                newIndex = page.indexOf('<td>', newIndex) + 4 // Increase Index
                text = (page.substring(newIndex, page.indexOf('</td>', newIndex)))
                
                if (text.substring(0, 1) == "<") { 
                  updateArrayLocal[9] = text.substring(text.indexOf('">') + 2, text.indexOf('</', 3)) // Save Location if it's a link
                } else {
                  updateArrayLocal[9] = text // Save Location
                }
                
                newIndex = page.indexOf('<td>', newIndex) + 4
                text = page.substring(newIndex, page.indexOf('</td>', newIndex))
                updateArrayLocal[6] = text // Save Date
                
                // If the date has passed, mark file as created/skipped
                if ((new Date(text)).getTime() < (new Date()).getTime() - MILLS*2) {
                  updateArrayLocal[13] = "TRUE" // Mark as skipped
                  updateArrayLocal[14] = "TRUE" // Mark as skipped
                }
                
                newIndex = page.indexOf('<td>', newIndex) + 4 // Increase Index
                
                // If the game is rescheduled, add helper text
                if (text == '<span class="game_status">To Be Rescheduled</span>') { // Added fix for Rescheduled games
                  updateArrayLocal[7] = "To Be Rescheduled"
                  updateArrayLocal[8] = "To Be Rescheduled"
                  
                  // Else if the game is TBA, add helper text
                } else if (text == "TBA") {
                  updateArrayLocal[7] = "TBA" 
                  updateArrayLocal[8] = "TBA"
                  
                  // Else save time
                } else {
                  text = page.substring(newIndex, page.indexOf('</td>', newIndex)) // Get text of time
                  updateArrayLocal[7] = text.substring(0, text.indexOf(' -')) // Save start time text
                  updateArrayLocal[8] = text.substring(text.indexOf('- ') + 2, text.length) // Save end time text
                }
                
                
                newIndex = page.indexOf('</tr>', newIndex) + 4  // Increase Index
                
                // If the sheet had been existing and this current event already exists in the rows, update and compare with new values
                if (made == false && sheetArray[row] != undefined) {
                  
                  // Keep checkmarks from sheet
                  updateArrayLocal[11] = sheetArray[row][11]
                  updateArrayLocal[12] = sheetArray[row][12]
                  updateArrayLocal[13] = sheetArray[row][13]
                  updateArrayLocal[14] = sheetArray[row][14]
                  
                  // If the event is upcoming, mark an upcoming post to be created
                  if (settingsArray[1][1] == "TRUE" && new Date(updateArrayLocal[6].substring(4)).getTime() - MILLS*settingsArray[1][2] + MILLS/24*settingsArray[1][3] < (new Date()).getTime() && sheetArray[row][13] != "TRUE") {
                    updateArrayLocal[11] = "TRUE"
                    updater++
                  }
                  
                  // If the events do not match up, save that a new event was added
                  if (updateArrayLocal[6] != sheetArray[row][6]) {
                    added = true
                    
                    // If a new score was updated, mark a score post to be created 
                  } else if (settingsArray[1][0] == "TRUE" && updateArrayLocal[10] == "" && sheetArray[row][10] == "No Score Available" && sheetArray[row][14] != "TRUE") {
                    updateArrayLocal[12] = "TRUE" 
                    updater++
                  }
                  
                  // If the events match up and scores are empty, add scores.
                  if (added != true) {
                    if (sheetArray[row][3] != "") {
                      updateArrayLocal[3] = sheetArray[row][3]
                    }
                    if (sheetArray[row][5] != "") {
                      updateArrayLocal[5] = sheetArray[row][5]
                    }
                  }
                  
                }
                
                added = false // Reset variables
                row ++ // Next iteration
                  
                  updateArray.push(updateArrayLocal) // Push temp 1D array to 2D array to be put into the place
              }
              
              // If events exist for sport, save 2D array into spreadsheet
              if (row > 0) {
                sheetSport.getRange(2, 1, row, 15).setValues(updateArray)
              }
            }
          }
        }
      }
    }
    
    sports.getRange(2, 1, cur-2, 5).setValues(sportArray) // Save sports list page with updated info
    addToUpdate(0, updater) // Update status number
    
  }
}

// autoRun(): Automatically run functions every 5 minutes.

function autoRun() {
  
  if (settingsArray[7][0] == "TRUE") {
    
    saveData()
    editSlide()
    
    if (settingsArray[4][3] == "TRUE") {
      autoClearNews() 
    }
    
    
    
  }
}

// autoRun2(): Automatically run functions once a day.

function autoRun2() {
  if (settingsArray[7][0] == "TRUE") {
    savePhoto()
  }
}
