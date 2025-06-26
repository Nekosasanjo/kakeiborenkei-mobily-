function CSVtoSS() {

 const folder = DriveApp.getFolderById('csvを格納したフォルダのID');//csvを格納したフォルダのID
 const files  = folder.getFiles();
 const file   = files.next();
 const fileId = file.getId();

 const blob   = DriveApp.getFileById(fileId).getBlob();
 const csv    = blob.getDataAsString();
 const values = Utilities.parseCsv(csv);

 const sheet = SpreadsheetApp.getActiveSheet();
 sheet.getRange(1, 1, values.length, values[0].length).setValues(values);

}
