// This Script Admin,Owner and Developer is Haris Nasir.
// Any and all changes must be approved by Haris Nasir due to the complex build of this script.
// Friday 06/5/2020 18:00 ET


function createSheet() {
  // Initilizing Function to Create New Sheet for Daily Peacock VOD from Template 
  
  var sheet = SpreadsheetApp.getActiveSpreadsheet();
  
  // Get the Date in 6/6/20 format to use as Name for new sheet
  var date = getDate();
  //Logger.log("Sheet Name: "+date);
  
  // Create the sheet with the date name
  sheet.insertSheet(date);
  
  // Paste to sheet we just created
  maker(date);
  
}

function getDate(){
 
  // Objective is to get Date in format m/d/yy
  // ex; 6/6/20 & 6/10/20
  
  // Create a new Date object
     var date = new Date();
  
  // Get the Various parts as simplified digit values, no decimal points or leading zeros
     var month = date.getMonth()+1;month = month.toString().slice(-1);
     var day = date.getDate(); if(day.toString().length==1){var day = day.toString().slice(-1);}
     var year = date.getFullYear().toString().slice(-2);

        // Logger.log("\nMonth: "+month+"\nDay: "+day+"\nYear: "+year);

     // Return the Date that will be used as the name for the sheet
     var d = month+"/"+day+"/"+year;
        //Logger.log(d);
   return d
}

function maker(name){
  // Copy Template from Template Sheet and Paste into New Sheet
 var sheetName = name;
 var targetSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
 var template = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Template");
  
  // Select the Range we want to copy from Template
  template.getRange("A1:I12").activate();
  
  // Select new Sheet and paste the Template values
  targetSheet.getRange("A1").activate();
  targetSheet.getRange('Template!A1:I12').copyTo(targetSheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
  
  // Change the Date in Header
  targetSheet.getRange("A4").activate();
  targetSheet.getRange("A4").activate().setValue("Daily Status Report, "+sheetName);
}

