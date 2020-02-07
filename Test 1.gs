// THIS SCRIPT WILL READ DATA FROM A FROM ONE SHEET AND THEN SCAN A 
// SECOND SHEET TO FIND THE Licence Plate AND EMAIL the student the parking violation

// GLOBAL VARIABLES
var target; // Licence Plate of offender
var code; // Access code
var rows = 35000; // Rows in Parent Access Code sheet from Schoology
var fName; // Students first name
var lName; // Students last name
var violation; // Parking offense
var studentEmail; // email to send the access code to
var completion; // true = Access code sent false = access code not found
var date = "11/24/19"  // This is the last updated date sent to parents in the code not found email

function main() // RUNS all the functions
{
  
  // Grab sheet 
  var sheet = SpreadsheetApp.getActiveSheet();
  var ss = SpreadsheetApp.openById("1oElQkT_XinHa5BWzmz7-GDBfObtssFNqqgP2fqN8Xbo"); //Test - OSH Project (Responses)
  var activeSheet = ss.getSheets()[1]; // "access data on different tabs"
  ss.setActiveSheet(sheet); // First Sheet
  
  // Grab info to look up
  GrabStudentID();
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheets()[1];
  
  // Returns the active range
  var range = sheet.getActiveRange();
  
  var data1 = sheet.getRange(2,4).getValue(); // Grabs Parent Access code of row 2
  Logger.log(data1);
  
  
  // Grabs the 35000 Parent Access codes in 2d array
  var numRows = 35000;
  // startRow, startColumn, numRows, NumColumns
  var values = sheet.getSheetValues(1, 1, numRows, 4);
   
 // Loop Throught the Data
 completion = false;
 for(var i=0; i<numRows; i++)
 {
      // Do this if matching StudentID if found
      if(values[i][0]==target)
      {    
        Logger.log("TARGET FOUND");
        Logger.log(values[i][3]); // Access Code
        code = values[i][3];
        fName = values[i][2];
        lName = values[i][1];
        emailStudent();
        completion = true; // Code was found, end loop
      }
 }

  
 markEmailed(); // Mark on the spreadsheet 1 that the parent was emailed
  
}

// This function grabs the last form entries Student ID & Email 
function GrabStudentID()
{
 var ss = SpreadsheetApp.getActiveSpreadsheet();
 var sheet = ss.getSheets()[0]; // Changes the Sheet Tab to look at
                                // [0] reads Form Response Sheet
 var sheetDB = ss.getSheets()[1]

 // This logs the value in the very last cell of this sheet
 var lastRow = sheet.getLastRow();
 var lastID = sheet.getRange(lastRow, 3);// LIcence plate column
 var lastViolation = sheet.getRange(lastRow, 2); // Get Violation
 var lastTime = sheet.getRange(lastRow, 1); // Get time
 var lastEmail = sheetDB.getRange(lastRow, 5); // Email Address column


 // Assign values to global variables 
 target = lastID.getValue();
 studentEmail = lastEmail.getValue();
 violation = lastViolation.getValue();
 date = lastTime.getValue();

 
 // DEBUGGING
 Logger.log("Target: "+target);
 Logger.log("Student Email: "+ studentEmail);
 Logger.log("Parking Violation: "+ violation);
 Logger.log("Date: "+date);
}

// Craft a message and email the student
function emailStudent()
{
    // format email message
    var message = "*This is an automatic message*\n\nToday ("+date+ "\n" +
    "Your vehicle with the licence plate: "+target+ " was cited for the following parking violation: "+violation +"\n\n"+"DO NOT REPLY TO THIS MESSAGE"; 

    // format subject of the email
    var subject = "OSH Parking Violation - " + lName + ", " + fName + " - " + "LP:"+target;
    
    // send the email
    MailApp.sendEmail(studentEmail, subject, message);
}


// This function will right in the response sheet that an email was sent
function markEmailed()
{
  // Set EMAIL
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheets()[0]; // Changes the Sheet Tab to look at
                                 // [0] reads Form Response Sheet
  var lastRow = sheet.getLastRow();
  var lastEmail = sheet.getRange(lastRow,4); // Column 'F'
  
  // Mark Sheet with status
  if(completion==true)
    lastEmail.setValue("EMAIL_SENT_COMPLETE");
  else 
    lastEmail.setValue("LP_NOT_FOUND"); 
}












