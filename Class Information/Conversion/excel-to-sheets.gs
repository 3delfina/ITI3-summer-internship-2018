///////////////////////////////////////
// Excel to Google Sheet conversion  //
///////////////////////////////////////

function fixAllRegisters() {
  
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet()
  var conversion_sheet = SpreadsheetApp.getActiveSheet();
  var REGISTERS_FOLDER_ID = conversion_sheet.getRange("C19").getValue();
  
  spreadsheet.toast("Starting...", "Progress", 5)
  
  // Get all files from registers folder
  // Check if folder exists
  try {
    var register_folder = DriveApp.getFolderById(REGISTERS_FOLDER_ID);
  } catch(err) {
    Browser.msgBox("Folder containing registers not found. ");
    return;
  }
  
  var this_id = SpreadsheetApp.getActive().getId();  
  var excel_files = register_folder.getFiles();
  
  if(!(excel_files.hasNext())){
    Browser.msgBox("Folder " + register_folder.getName() + " is empty. ");
    return;
  }
  
  // Iterate over all files in folder
  while (excel_files.hasNext()) {
    
    var current = excel_files.next(); 
    
    // Proccess only spreadsheet type files
    if ((current.getMimeType() == MimeType.GOOGLE_SHEETS) || (current.getMimeType() == MimeType.MICROSOFT_EXCEL)){
      
      // Ignore spreadsheet holding this script
      if (current.getId() != this_id) {
        
        // If file type is excel, convert it to google sheets type
        if (current.getMimeType() == MimeType.MICROSOFT_EXCEL) {
          
          spreadsheet.toast("Converting Excel register: " + current.getName(), "Progress", 5)
          
          var xlsxBlob = current.getBlob();
          var old_id = current.getId();
          
          var filetype = ".xlsx"
          var expected_filename = current.getName().slice(0,-(filetype.length))
          
          var expected_files = register_folder.getFilesByName(expected_filename)
          
          // If already existing, update file
          if (expected_files.hasNext()) {
            
            gsheet_file = expected_files.next()
            
            try {
              current = Drive.Files.update({title: gsheet_file.getName(), mimeType: gsheet_file.getMimeType()}, gsheet_file.getId(), xlsxBlob, {convert: true});
            }  catch(err) {
              Browser.msgBox("Cannot overwrite existing Google Sheet file with Excel equivalent. Aborting conversion.\nDetails: " + err.message);
              return;
            }
            
            // else create new file  
          } else {
            
            var file_data = {title: current.getName(), 
                             mimeType: MimeType.GOOGLE_SHEETS,
                             key: old_id,
                             parents: [{id: Drive.Files.get(old_id).parents[0].id}]
                            };
            
            try {
              current = Drive.Files.insert(file_data, xlsxBlob, {convert: true}); // Needs permission to Drive API
            } catch(err) {
              Browser.msgBox("Cannot create new Google Sheet file from Excel file. Aborting conversion.\nDetails: " + err.message);
              return;
            } 
          }
          
          
          // Delete the old file
          try {
            spreadsheet.toast("Deleting Excel version...", "Progress", 5)
            Drive.Files.remove(old_id);
          } catch(err) {
            Browser.msgBox("Excel file delelation failed. Cannot convert the file from Excel to Google Sheets. Aborting conversion.\nDetails: " + err.message);
            return;
          } 
          
          // Work on the converted   
          try {
            current = DriveApp.getFileById(current.id);  
          } catch(err) {
            Browser.msgBox("Cannot open newly converted Google Sheets file. Aborting conversion.\nDetails: " + err.message);
            return;
          } 
          
          
        // Fix checkboxes in current spreadsheet
        spreadsheet.toast("Fixing formatting in newly converted Google Sheets register: " + current.getName(), "Progress", 5)
        fixRegister(current);
        
          
        }
        
        
      }
    }   
  }
  
  // Save changes
  spreadsheet.toast("Saving changes...", "Progress", 5)
  SpreadsheetApp.flush();
  
  Browser.msgBox("Finished.");
  
}

// Change true/false values into checkboxes
function fixRegister(spreadsheet_file) {
  
  Logger.log("Starting " + spreadsheet_file.getName() + "...");
  
  // Get proper sheet
  var spreadsheet = SpreadsheetApp.open(spreadsheet_file);
  
  try {
    offline_sheet = spreadsheet.getSheetByName("Offline"); 
    // Fix the offline checkbox
    offline_sheet.getRange(9, 3).setDataValidation(SpreadsheetApp.newDataValidation().requireCheckbox());
  } catch(err) {
    Browser.msgBox("Cannot open Offline sheet on " + spreadsheet_file.getName() + ".\nProceeding with remaining registers.");
    return;
  }
  

  
  try {
    sheet = spreadsheet.getSheetByName("Class"); 
  } catch(err) {
    Browser.msgBox("Cannot open Class sheet on " + spreadsheet_file.getName() + ".\nProceeding with remaining registers.");
    return;
  }
  
  // Set constants
  var MEMBERSHIP_COLUMN = 5;                // Constant in template
  var START_COLUMN = MEMBERSHIP_COLUMN + 1; // Start of repeatable lessons columns
  var START_ROW = 11;                       // Constant in template
  
  // Get boundaries
  try {
    var last_row_used = sheet.getLastRow();
    var last_col_used = sheet.getLastColumn();
  } catch(err) {
    Browser.msgBox("Class sheet in " + spreadsheet_file.getName() + " is empty or corrupted.\nProceeding with remaining registers.");
    return;
  }
  
  // Check if there are any members in register
  if (START_ROW <= last_row_used) {
    try {
      // Change whole range into checkboxes
      target_range = sheet.getRange(START_ROW, MEMBERSHIP_COLUMN, last_row_used - START_ROW + 1, last_col_used - MEMBERSHIP_COLUMN + 1);
      
      // Change data validation rules to checkboxes    
      target_range.setDataValidation(SpreadsheetApp.newDataValidation().requireCheckbox());
      
      // Change font to make checkboxes bigger
      target_range.setFontSize(25);
      
      
      // Clear validation rules and font size for comment column and non boolean values in payment column
      var cell;
      var target_range;       
      for (var col = START_COLUMN; col <= last_col_used; col++) {
        
        // For comment column, clear whole column
        if (col % 3 == 2){     
          
          // Change data validation rules to checkboxes    
          sheet.getRange(START_ROW, col, last_row_used - START_ROW + 1, 1).setDataValidation(null).setFontSize(11);
          
          // For payment column, clear only non-booleans  
        } else if (col % 3 == 1){
          
          var values = sheet.getRange(START_ROW, col, last_row_used - START_ROW + 1, 1).getValues();        
          var to_clear = [];
          
          for (var row = 0; row < (values.length); row++) {          
            if (!(values[row][0] == true || values[row][0] == false)) {
              to_clear.push(columnToLetter(col) + (row + START_ROW));
            }
          }
          
          if (to_clear.length > 0) {
            var ranges_to_set = sheet.getRangeList(to_clear).getRanges();
            for (var index = 0; index < to_clear.length; index++){
              // Clear
              ranges_to_set[index].setDataValidation(null).setFontSize(11);
            }
          }         
        }
 
      }
    } catch(err) {
      Browser.msgBox("Checkbox fixing on spreadsheet " + spreadsheet_file.getName() + " falied. Proceeding with remaining registers. Details: " + err.message);
      return;
    }
  } else {
    Browser.msgBox("Register for class: " + spreadsheet_file.getName() + " is empty. Proceeding with remaining registers to convert.");
    return;
  }
  
  Logger.log("Finished " + spreadsheet_file.getName() + ".");
  
};

// Courtesy of rheajt; https://gist.github.com/rheajt/d48a54be3aad01a1931a1a433dc99e5c
function columnToLetter(column) {
  
  var temp;
  var letter = '';
  
  while (column > 0) {
    temp = (column - 1) % 26;
    letter = String.fromCharCode(temp + 65) + letter;
    column = (column - temp - 1) / 26;
  }
  
  return letter;
}
