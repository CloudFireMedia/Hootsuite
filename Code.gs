var SCRIPT_NAME    = 'HootSuite'
var SCRIPT_VERSION = 'v1.0'

// Public Functions
// ----------------

function deleteRows()                  {return deleteRows_()}
function saveAsCSV()                   {return saveAsCSV_()}
function importRecurringEvents(config) {return importRecurringEvents_(config)}

// Private Functions
// -----------------

/**
 * Delete all the rows from the sheet
 */

function deleteRows_() {
  
  var sheet = SpreadsheetApp.getActive().getSheetByName("Current Month");
  sheet.deleteRows(2, sheet.getLastRow() - 2);
  
} // deleteRows_()

/**
 * Save the "Current Month" sheet as a CSV file
 */

function saveAsCSV_() {///yeah, this doesn't work

  //the developer got this from: https://gist.github.com/mderazon/9655893

  ////Hey Chad, do you need this in drive or can it download to your local system when you need it?
  //why not link to the download url instead? show link in a madal window
  //uses the doc name - sheet name
  //Utilities.formatString('https://docs.google.com/spreadsheets/d/%s/export?format=%s&gid=%s', sheet.getParent().getId(), 'csv', sheet.getSheetId())//csv|pdf|odt|doc
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Current Month');
  var fileName = 'Upload to Hootsuite.csv'; //don't forget the .csv extension
  var csvFile = convertRangeToCsvFile(sheet);
  DriveApp.createFile(fileName, csvFile); // create a file in the Docs List with the given name and the csv data
  Browser.msgBox("File has been saved to Google Drive as '" + fileName + "'");
  return 
  
  // Private Functions
  // -----------------

  function convertRangeToCsvFile(sheet) {
  
    var range = sheet.getDataRange();
    
    try {
    
      var data = range.getValues();
      
      // loop through the data in the range and build a string with the csv data
      if (data.length > 1) {
      
        var csv = "";
        for (var row = 1; row < data.length; row++) {
          var str1 = data[row][1].replace(/[\u2018\u2019]/g, "'");//left and right single quotes to regular single quotes
          csv += '"' + data[row][0] + '","' + str1 + '","' + data[row][2] + '"\r\n';
        }
        var csvFile = csv;
      }
      
      return csvFile;
      
    } catch (err) {
      Logger.log(err);
      Browser.msgBox(err);
    }
    
  } // saveAsCSV_.convertRangeToCsvFile()

} // saveAsCSV_()

/**
 * Import recurring events into the "Current Month" sheet from the master sunday announcement GDoc
 */

function importRecurringEvents_(config) {

  insertRecurringContent(31, config); //number of days to search for announcement date 
  
  return 
  
  // Private Functions
  // -----------------
  
  function insertRecurringContent(daysToSearch, config) {
  
    var announcementsFile = DocumentApp.openById(config.files.announcements.master); //source file
    var spreadsheet = SpreadsheetApp.getActive();
    var sheet = spreadsheet.getSheetByName("Current Month");
    
    if (sheet === null) {
      throw new Error('Error opening sheet "Current Month" in ' + spreadsheet.getName())
    }
    
    var BASE_DATE = new Date(config.basedate);
    var DATE_FOR_THIS_EXECUTION = new Date(BASE_DATE.getYear(), BASE_DATE.getMonth(), BASE_DATE.getDate());
    
    var body = announcementsFile.getBody();
    var numChildren = body.getNumChildren();
    var pageBreakCounter = 0;
    var insertIndex = 0;
    var recurringContentParagraphs = [];
    
    for (var i = 0; i < numChildren; ++i) {
    
      var child = body.getChild(i);
      var childType = child.getType();
      
      if (childType == DocumentApp.ElementType.PAGE_BREAK) {
        ++pageBreakCounter;
        continue;
      }
      
      if (childType == DocumentApp.ElementType.PARAGRAPH) {
      
        var paragraph = child.asParagraph();
        var paragraphNumChildren = paragraph.getNumChildren();
        
        if (pageBreakCounter == 1) {
        
          var text = paragraph.getText().trim();
          
          if ((text !== "[ RECURRING CONTENT ]") && (text !== "")) {
          
            Logger.log('Found non recurring content')
            
            var paraText = child.asParagraph().getText();
            
            Logger.log('paraText: ' + paraText)
            
            if (paraText.indexOf("??>>??") === -1) {
              recurringContentParagraphs.push(child.asParagraph());
            }
          }
        }
        
        if (pageBreakCounter > 1) {
          insertIndex = i;
          break;
        }
        
        var foundPageBreak = false;
        
        for (var j = 0; j < paragraphNumChildren; ++j) {
        
          var paragraphChild = paragraph.getChild(j);
          var paragraphChildType = paragraphChild.getType();
          
          if (paragraphChildType == DocumentApp.ElementType.PAGE_BREAK) {
            ++pageBreakCounter;
            foundPageBreak = true;
            break;
          }
        }
        
        if (foundPageBreak) {
          continue;
        }
      }
      
    } // for each child in the doc
        
    recurringContentParagraphs.forEach(function(recurringContentParagraph, index) {
      Logger.log('recurringContentParagraphs: [' + index + '] ' + recurringContentParagraph.getText())    
    })
        
    var content = []
    
    for (var i = 0; i < recurringContentParagraphs.length; ++i) {
    
      var paragraph = recurringContentParagraphs[i];
      var text = paragraph.editAsText();
      var textAsString = text.getText();
      var criteriaString = "";
      var foundStartChar = 0;
      var foundEndChar = 0;
      
      for (var j = 0; j < textAsString.length; ++j) {
      
        if (textAsString[j] == "<") {
          foundStartChar += 1;
        }
        if (textAsString[j] == ">") {
          foundEndChar += 1;
        }
        if ((foundStartChar == 2) && (foundEndChar <= 2)) {
          criteriaString += textAsString[j];
        }
        if (foundEndChar == 2) {
          break;
        }
      }
      
      criteriaString = criteriaString.toLowerCase().replace(/ /g, ""); //flexibility with capitalization and spaces   
      Logger.log('criteriaString: ' + criteriaString);
      var fullDate = new Date();
      var todaysDate = new Date(fullDate.getYear(), fullDate.getMonth(), fullDate.getDate());
      var currentMonth = new Date(fullDate.getYear(), fullDate.getMonth());
      
      if (criteriaString.indexOf("first sunday of the month".replace(/ /g, "")) != -1) {
        
        for (var d = 0; d < daysToSearch; ++d) {
          var date = new Date(todaysDate.getYear(), todaysDate.getMonth(), todaysDate.getDate() + d);
          if (date.getDay() == 0) {
            if (date.getDate() / 7 <= 1) {
              var dateFormated = date.getMonth() + 1 + "/" + date.getDate() + "/" + date.getYear() + " 6:00:00";
              if (dateFormated != undefined) {
                var arr5 = paragraph.getText().split(";");
                var extractedPara = arr5[1];
                extractedPara = extractedPara.trim();
                sheet.appendRow([dateFormated, extractedPara]);
              }
            }
          }
        }
        
      } else if (criteriaString.indexOf("second sunday of the month".replace(/ /g, "")) != -1) {
        
        for (var d = 0; d < daysToSearch; ++d) {
          var date = new Date(todaysDate.getYear(), todaysDate.getMonth(), todaysDate.getDate() + d);
          if (date.getDay() == 0) {
            if ((date.getDate() / 7 > 1) && (date.getDate() / 7 <= 2)) {
              var dateFormated = date.getMonth() + 1 + "/" + date.getDate() + "/" + date.getYear() + " 6:00:00";
              if (dateFormated != undefined) {
                var arr5 = paragraph.getText().split(";");
                var extractedPara = arr5[1];
                extractedPara = extractedPara.trim();
                sheet.appendRow([dateFormated, extractedPara]);
              }
            }
          }
        }
        
      } else if (criteriaString.indexOf("third sunday of the month".replace(/ /g, "")) != -1) {
        
        for (var d = 0; d < daysToSearch; ++d) {
          var date = new Date(todaysDate.getYear(), todaysDate.getMonth(), todaysDate.getDate() + d);
          if (date.getDay() == 0) {
            if ((date.getDate() / 7 > 2) && (date.getDate() / 7 <= 3)) {
              var dateFormated = date.getMonth() + 1 + "/" + date.getDate() + "/" + date.getYear() + " 6:00:00";
              if (dateFormated != undefined) {
                var arr5 = paragraph.getText().split(";");
                var extractedPara = arr5[1];
                extractedPara = extractedPara.trim();
                sheet.appendRow([dateFormated, extractedPara]);
              }
            }
          }
        }
        
      } else if ((criteriaString.indexOf("fourth sunday of the month".replace(/ /g, "")) != -1) && 
                 (criteriaString.indexOf("fifth sunday exists in the same month".replace(/ /g, "")) != -1)) {
        
        for (var d = 0; d < daysToSearch; ++d) {
          var date = new Date(todaysDate.getYear(), todaysDate.getMonth(), todaysDate.getDate() + d);
          if (date.getDay() == 0) {
            if ((date.getDate() / 7 > 3) && (date.getDate() / 7 <= 4)) {
              var hasFifthSunday = false;
              for (var dd = date.getDate(); dd <= daysInMonth(date.getMonth(), date.getYear()); ++dd) {
                var testForFifthSundayDate = new Date(date.getYear(), date.getMonth(), dd);
                Logger.log('testForFifthSundayDate: ' + testForFifthSundayDate);
                if ((testForFifthSundayDate.getDay() == 0) && (testForFifthSundayDate.getDate() / 7 > 4)) {
                  hasFifthSunday = true;
                  break;
                }
              }
              if (hasFifthSunday == false) {
                continue;
              }
              var dateFormated = date.getMonth() + 1 + "/" + date.getDate() + "/" + date.getYear() + " 6:00:00";
              if (dateFormated != undefined) {
                var arr5 = paragraph.getText().split(";");
                if (arr5[2] != undefined) {
                  var extractedPara = arr5[2];
                  extractedPara = extractedPara.trim();
                  sheet.appendRow([dateFormated, extractedPara]);
                }
              }
            }
          }
        }
        
      } else if (criteriaString.indexOf("fourth sunday of the month".replace(/ /g, "")) != -1) {
        
        for (var d = 0; d < daysToSearch; ++d) {
          var date = new Date(todaysDate.getYear(), todaysDate.getMonth(), todaysDate.getDate() + d);
          if (date.getDay() == 0) {
            if ((date.getDate() / 7 > 3) && (date.getDate() / 7 <= 4)) {
              var dateFormated = date.getMonth() + 1 + "/" + date.getDate() + "/" + date.getYear() + " 6:00:00";
              if (dateFormated != undefined) {
                var arr5 = paragraph.getText().split(";");
                var extractedPara = arr5[1];
                extractedPara = extractedPara.trim();
                sheet.appendRow([dateFormated, extractedPara]);
              }
            }
          }
        }
        
      } else if (criteriaString.indexOf("last sunday of the month".replace(/ /g, "")) != -1) {
        for (var d = 0; d < daysToSearch; ++d) {
          var date = new Date(todaysDate.getYear(), todaysDate.getMonth(), todaysDate.getDate() + d);
          if ((date.getDay() == 0) && (daysInMonth(date.getMonth()) - date.getDate() < 7)) {
            var dateFormated = date.getMonth() + 1 + "/" + date.getDate() + "/" + date.getYear() + " 6:00:00";
            if (typeof(dateFormated) == "undefined") {
              
            } else {
              var arr5 = paragraph.getText().split(";");
              var extractedPara = arr5[1];
              extractedPara = extractedPara.trim();
              sheet.appendRow([dateFormated, extractedPara]);
            }
          }
        }
        
      } else if (criteriaString.indexOf("should be appended on ".replace(/ /g, "")) != -1) {
        
        var specificDates = criteriaString.split("should be appended on ".replace(/ /g, ""))[1].split(",");
        for (var d = 0; d < specificDates.length; ++d) {
          var month = parseInt(specificDates[d].split(".")[0].replace(/\D/g, '')) - 1;
          var day = parseInt(specificDates[d].split(".")[1].replace(/\D/g, ''));
          var date = new Date(todaysDate.getYear(), month, day);
          if (date < todaysDate) {
            date = new Date(todaysDate.getYear() + 1, month, day);
          }
          var isWithinSearchRange = false;
          for (var dd = 0; dd < daysToSearch; ++dd) {
            if ((new Date(todaysDate.getYear(), todaysDate.getMonth(), todaysDate.getDate() + dd)).toDateString() == date.toDateString()) {
              isWithinSearchRange = true;
              break;
            }
          }
          
          if (isWithinSearchRange == false) {
            continue;
          }
          
          var sunday = "";
          var day = date.getDay();
          if (day == 0) {
            if (date.getDate() / 7 <= 1) {
              sunday = "First Sunday of the month";
            } else if (date.getDate() / 7 <= 2) {
              sunday = "Second Sunday of the month";
            } else if (date.getDate() / 7 <= 3) {
              sunday = "Third Sunday of the month";
            } else if (date.getDate() / 7 <= 4) {
              sunday = "Fourth Sunday of the month";
            } else {
              sunday = "Fifth Sunday of the month";
            }
          }
          var dateFormated = date.getMonth() + 1 + "/" + date.getDate() + "/" + date.getYear() + " 6:00:00";
          
          if (dateFormated != undefined) {
            var arr5 = paragraph.getText().split(";");
            if (arr5[1] == undefined) {
              var arr6 = paragraph.getText().split(">>");
              var extractedPara = arr6[1];
              extractedPara = extractedPara.trim();
              sheet.appendRow([dateFormated, extractedPara]);
            } else {
              if (arr5[1] != undefined) {
                var extractedPara = arr5[1];
                extractedPara = extractedPara.trim();
                sheet.appendRow([dateFormated, extractedPara]);
              }
            }
          }
        }
      } 
    }
    
    sortSpreadSheet();
     
    return 
    
    // Private Functions
    // -----------------

    function sortSpreadSheet() {
      
      var sheetArray = [];
      var dataRange = sheet.getDataRange();
      var values = dataRange.getValues();
      
      for (var i = 1; i < values.length; i++) {
      
        var row = "";
        var sheetDate = values[i][0];
        var typeIs = typeof sheetDate;
        
        if (typeIs != "object") {
        
          var arr6 = sheetDate.split(" ");
          var arr7 = arr6[0].split(".");
          
          if (arr7[0].length == 1) {
            var month = "0" + arr7[0];
          } else {
            var month = arr7[0];
          }
          
          if (arr7[1].length == 1) {
            var day = "0" + arr7[1];
          } else {
            var day = arr7[1];
          }
          
          var newFormatedDate = arr7[2] + "-" + month + "-" + day;
          var newdate = new Date(newFormatedDate);
          
        } else {
        
          sheetDate = new Date(sheetDate).toISOString();
          //2017-11-05T
          var arr6 = sheetDate.split("T");
          newFormatedDate = arr6[0];
          var newdate = new Date(newFormatedDate);
        }
        
        sheetArray.push({
          date: newFormatedDate,
          dateForSheet: values[i][0],
          paragraph: values[i][1],
          registrationUrl: values[i][2]
        });
      }
      
      sheetArray.sort(sortByDate);
      
      var tempI = 1;
      for (data in sheetArray) {
        tempI++;
        
        sheet.getRange(tempI, 1).setValue(sheetArray[data].dateForSheet);
        sheet.getRange(tempI, 2).setValue(sheetArray[data].paragraph);
        sheet.getRange(tempI, 3).setValue(sheetArray[data].registrationUrl);
      }
      
      return
      
      // Private Functions
      // -----------------
      
      function sortByDate(a, b) {
        if (new Date(a.date) < new Date(b.date)) return -1;
        else if (new Date(a.date) > new Date(b.date)) return 1;
        else return 0;
      }
      
    } // importRecurringEvents_.insertRecurringContent.sortSpreadSheet()

    function daysInMonth(month, year) {
    
      switch (month) {
        case 0:
          return 31;
        case 1:
          if (year % 4 == 0) return 29;
          else return 28;
        case 2:
          return 31;
        case 3:
          return 30;
        case 4:
          return 31;
        case 5:
          return 30;
        case 6:
          return 31;
        case 7:
          return 31;
        case 8:
          return 30;
        case 9:
          return 31;
        case 10:
          return 30;
        case 11:
          return 31;
        default:
          throw "Invalid month number: " + month;
      }
      
    } // importRecurringEvents_.insertRecurringContent.daysInMonth()
    
  } // importRecurringEvents_.insertRecurringContent()
   
} // importRecurringEvents_()