///////////////////////////////////////
// Google Sheet to Excel conversion  //
///////////////////////////////////////
function gsheetToExcel() {


    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var conversion_sheet = SpreadsheetApp.getActiveSheet();
    var REGISTERS_FOLDER_ID = conversion_sheet.getRange("C19").getValue();

    registers_folder = DriveApp.getFolderById(REGISTERS_FOLDER_ID);

    spreadsheet.toast("Starting...", "Progress", 5);

    var files_iterator = registers_folder.getFiles();
    var current;

    if (!(files_iterator.hasNext())) {
        Browser.msgBox("Folder " + registers_folder.getName() + " has no Google Sheet files. ");
        return;
    }

    // Iterate over files
    while (files_iterator.hasNext()) {

        current = files_iterator.next();
        spreadsheet.toast("Converting register: " + current.getName(), "Progress", 5);

        // If gsheet file, convert to xlsx
        if (current.getMimeType() == MimeType.GOOGLE_SHEETS) {
          
          // Check if Excel version already exists
            if (!(registers_folder.getFilesByName(current.getName() + ".xlsx").hasNext())) {

                // If file not offline on any device, convert to xlsx and notify user
                current_file = DriveApp.getFileById(current.getId())
                current_spreadsheet = SpreadsheetApp.open(current_file)
                current_sheet = current_spreadsheet.getSheetByName("Offline")

                if (current_sheet != null && !(current_sheet.getRange(9, 3).getValue())) {

                    ////////////////////////////
                    // Make formatting changes
                    ////////////////////////////
                    spreadsheet.toast("Making formatting changes in register: " + current.getName(), "Progress", 5)

                    current_sheet = current_spreadsheet.getSheetByName("Class")

                    // Set constants
                    var MEMBERSHIP_COLUMN = 5; // Constant in template
                    var START_COLUMN = MEMBERSHIP_COLUMN + 1; // Start of repeatable lessons columns
                    var START_ROW = 11; // Constant in template

                    // Get boundaries
                    var last_row_used = current_sheet.getLastRow();
                    var last_col_used = current_sheet.getLastColumn();

                    Logger.log(last_row_used + ":" + last_col_used + " - " + current_spreadsheet.getName())
                    // Check if there are any members in register
                    try {
                        // Change whole range into checkboxes
                        target_range = current_sheet.getRange(START_ROW, MEMBERSHIP_COLUMN, last_row_used - START_ROW + 1, last_col_used - MEMBERSHIP_COLUMN + 1);

                        // Change font to make checkboxes bigger
                        target_range.setFontSize(11);

                        Logger.log(target_range.getA1Notation())

                    } catch (err) {
                        Browser.msgBox("Checkbox fixing on spreadsheet " + current_spreadsheet.getName() + " falied. Proceeding with remaining registers. Details: " + err.message);
                    }


                    ////////////////////////////
                    // Convert
                    ////////////////////////////
                    spreadsheet.toast("Converting register: " + current.getName(), "Progress", 5)

                    var file_id = current.getId();

                    // Export to Excel file and get blob
                    try {

                        var url = "https://docs.google.com/spreadsheets/d/" + file_id + "/export?format=xlsx&access_token=" + ScriptApp.getOAuthToken();
                        var blob = UrlFetchApp.fetch(url).getBlob().setName(current.getName() + ".xlsx");

                    } catch (err) {
                        Browser.msgBox("Conversion failed for  " + current.getName() + ". Cannot export to Excel file.\nProceeding with remaining registers.\nDetails: " + err.message);
                    }

                    // Put newly converted file to registers folder
                    try {
                        registers_folder.createFile(blob);
                    } catch (err) {
                        Browser.msgBox("Conversion failed for  " + current.getName() + ". Cannot save Excel file.\nProceeding with remaining registers.\nDetails: " + err.message);

                    }

                } else {
                    if (current_sheet == null) {
                        Browser.msgBox("The file " + current.getName() + " has no Offline sheet. \nProceeding with remaining registers.");

                    } else {
                        Browser.msgBox("The file " + current.getName() + " is currently used offline. Please try again later.\nProceeding with remaining registers.");
                    }

                }

            } else {
                spreadsheet.toast("The file " + current.getName() + " has its Excel version already in the registers folder. \nProceeding with remaining registers.", "Progress", 5);
            }
        }
    }
    spreadsheet.toast("Finished.", "Progress", 5);
    Browser.msgBox("Finished.");

}