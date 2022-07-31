function main() {
  let fileName = "excelData.xlsx";
  let sheetName = "info";

  toast(`Importing ${sheetName} from ${fileName} ...`);
  let spreadsheetId = convertExcelToGoogleSheets(fileName);
  let importedSheetName = importDataFromSpreadsheet(spreadsheetId, sheetName);
  toast(`Successfully imported data from ${sheetName} in ${fileName} to ${importedSheetName}`);
}

function convertExcelToGoogleSheets(fileName) {
  let files = DriveApp.getFilesByName(fileName);
  let excelFile = null;
  if(files.hasNext())
    excelFile = files.next();
  else
    return null;
  let blob = excelFile.getBlob();
  let config = {
    title: "[Google Sheets] " + excelFile.getName(),
    parents: [{id: excelFile.getParents().next().getId()}],
    mimeType: MimeType.GOOGLE_SHEETS
  };
  //delete the old [Google Sheets] file
  let googledlt = DriveApp.getFilesByName("[Google Sheets] excelData").next().getId();
  Drive.Files.remove(googledlt);
  let spreadsheet = Drive.Files.insert(config, blob);
  return spreadsheet.id;
}

function importDataFromSpreadsheet(spreadsheetId, sheetName) {
  let spreadsheet = SpreadsheetApp.openById(spreadsheetId);
  let currentSpreadsheet = SpreadsheetApp.getActive();
  let todltsheet = currentSpreadsheet.getActiveSheet();
  let newSheet = currentSpreadsheet.insertSheet();
  currentSpreadsheet.setActiveSheet(todltsheet);
  currentSpreadsheet.deleteActiveSheet();
  currentSpreadsheet.setActiveSheet(newSheet);
  currentSpreadsheet.renameActiveSheet("info")
  let dataToImport = spreadsheet.getSheetByName(sheetName).getDataRange();
  let range = newSheet.getRange(1,1,dataToImport.getNumRows(), dataToImport.getNumColumns());
  range.setValues(dataToImport.getValues());
  return newSheet.getName();
}

function toast(message) {
  SpreadsheetApp.getActive().toast(message);
}
