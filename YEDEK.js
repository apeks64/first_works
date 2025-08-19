function yedekAl() {
  var spreadsheetId = '1J5EMoXVpyk2WkhUrBJ4xoW810Gqs-9y1Z64KZQMNMpQ'; // Yedek alınacak Google Sheet dosyasının ID'si
  var folderId = '1ulmkp1A8Bn7AsowgrYt1Z4DCnEYHx1BQ'; // Yedeklerin kaydedileceği Google Drive klasörünün ID'si
  
  // Google Sheets belgesini aç
  var spreadsheet = SpreadsheetApp.openById(spreadsheetId);
  
  // Drive üzerinden bu belgeyi kopyala
  var file = DriveApp.getFileById(spreadsheetId);
  var copy = file.makeCopy(spreadsheet.getName() + ' - Yedek - ' + new Date().toISOString(), DriveApp.getFolderById(folderId));
  
  Logger.log('Yedek başarıyla alındı: ' + copy.getName());
}

function setupTrigger() {
  // Zamanlayıcıyı ayarla: Her gün belirli bir saatte yedek alma işlemi yapacak
  ScriptApp.newTrigger('yedekAl')
    .timeBased()
    .everyWeeks(1)
    .atHour(1) // Burada istediğiniz saati belirleyebilirsiniz (0 - 23 arası)
    .create();
}
