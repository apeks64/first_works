// Tetikleyiciyi her gün saat 10'da çalışacak şekilde oluşturur
function createTimeDrivenTriggers() {
  ScriptApp.newTrigger('checkForNewExcelFile')
    .timeBased()
    .atHour(10)  // Saat 10'da çalışacak
    .everyDays(1) // Her gün çalışacak
    .create();
}

// checkForNewExcelFile fonksiyonu, her gün saat 10'da çalışacak şekilde tetiklenir
function checkForNewExcelFile() {
  var folderId = "1tEfzI99GOQj7kSniQxz3KiPzufCTxUsV"; // Klasör ID'si
  var latestFile = getLatestExcelFile(folderId);

  if (!latestFile) {
    Logger.log("Yeni dosya bulunamadı.");
    return;
  }

  var scriptProperties = PropertiesService.getScriptProperties();
  var lastProcessedFileId = scriptProperties.getProperty("lastProcessedFileId");

  // Eğer en son işlenen dosya ID'si, en yeni dosya ID'sinden farklıysa, yeni dosya işlenmeli
  if (lastProcessedFileId !== latestFile.getId()) {
    Logger.log("Yeni dosya bulundu: " + latestFile.getName());
    importLatestExcelToAtamalar();  // Excel'i ATAMALAR sayfasına aktar
    scriptProperties.setProperty("lastProcessedFileId", latestFile.getId()); // İşlenen dosya ID'sini güncelle
  } else {
    Logger.log("Yeni bir dosya yok. En son işlenen dosya: " + lastProcessedFileId);
  }
}

// Klasördeki en son yüklenen Excel dosyasını bul
function getLatestExcelFile(folderId) {
  var folder = DriveApp.getFolderById(folderId);
  var files = folder.getFiles();
  var latestFile = null;
  var latestDate = 0;

  while (files.hasNext()) {
    var file = files.next();
    if (file.getName().endsWith(".xlsx") && file.getLastUpdated().getTime() > latestDate) {
      latestDate = file.getLastUpdated().getTime();
      latestFile = file;
    }
  }

  if (latestFile) {
    Logger.log("Bulunan en yeni dosya: " + latestFile.getName());
    return latestFile;
  } else {
    Logger.log("Klasörde Excel dosyası bulunamadı.");
    return null;
  }
}

// Excel dosyasını ATAMALAR sayfasına aktarır
function importLatestExcelToAtamalar() {
  var folderId = "1tEfzI99GOQj7kSniQxz3KiPzufCTxUsV"; // Drive klasör ID'si
  var sheetId = "1J5EMoXVpyk2WkhUrBJ4xoW810Gqs-9y1Z64KZQMNMpQ"; // Google Sheet ID'si
  var sheetName = "ATAMALAR"; // Hedef sayfa adı

  var latestFile = getLatestExcelFile(folderId);

  if (!latestFile) {
    SpreadsheetApp.getUi().alert("Klasörde Excel dosyası bulunamadı.");
    return;
  }

  var blob = latestFile.getBlob();

  try {
    // Dosyayı Google Sheets olarak Drive'a yükle
    var resource = {
      title: latestFile.getName(),
      mimeType: MimeType.GOOGLE_SHEETS,
      parents: [{ id: folderId }]
    };

    var convertedFile = Drive.Files.create(resource, blob);
    Logger.log("Dönüştürülmüş dosya ID'si: " + convertedFile.id);

    Utilities.sleep(10000); // 10 saniye bekleme

    var tempSpreadsheet = SpreadsheetApp.openById(convertedFile.id);
    var tempSheet = tempSpreadsheet.getSheets()[0];
    Logger.log("Geçici Google Sheet açıldı: " + tempSpreadsheet.getId());

    // ATAMALAR sayfasına verileri aktar
    var targetSpreadsheet = SpreadsheetApp.openById(sheetId);
    var targetSheet = targetSpreadsheet.getSheetByName(sheetName);

    if (!targetSheet) {
      targetSheet = targetSpreadsheet.insertSheet(sheetName);
    }

    var range = tempSheet.getDataRange();
    var values = range.getValues();

    targetSheet.clear();
    targetSheet.getRange(1, 1, values.length, values[0].length).setValues(values);

    Logger.log("Veriler başarıyla aktarıldı.");

    // Geçici dosyayı sil
    DriveApp.getFileById(convertedFile.id).setTrashed(true);

    // Görevlendirilen Kişi TC Kimlik No'ya göre Ad Soyad değiştirme işlemi
    updateNamesFromBelgeno(sheetId, sheetName);

  } catch (e) {
    Logger.log("Hata oluştu: " + e.message);
    throw e;
  }
}

// BELGE_NO sayfasındaki TC Kimlik No'ya göre Ad Soyad bilgisini döndüren fonksiyon
function updateNamesFromBelgeno(sheetId, sheetName) {
  var targetSpreadsheet = SpreadsheetApp.openById(sheetId);
  var targetSheet = targetSpreadsheet.getSheetByName(sheetName);
  var data = targetSheet.getDataRange().getValues();

  var belgeNoSheet = targetSpreadsheet.getSheetByName("BELGE_NO");
  var belgeNoData = belgeNoSheet.getDataRange().getValues();

  // ATAMALAR sayfasındaki TC Kimlik No ve Ad Soyad sütunlarının indeksleri
  var tcColumnIndex = 4; // ATAMALAR sayfasındaki 5. sütun (Görevlendirilen Kişi TC Kimlik No)
  var nameColumnIndex = 5; // ATAMALAR sayfasındaki 6. sütun (Görevlendirilen Kişi Ad Soyad)

  var belgeNoTCColumnIndex = 0; // BELGE_NO sayfasındaki 1. sütun (TC Kimlik No)
  var belgeNoNameColumnIndex = 2; // BELGE_NO sayfasındaki 3. sütun (Görevlendirilen Kişi Ad Soyad)

  for (var i = 0; i < data.length; i++) {
    var tcValue = data[i][tcColumnIndex]; // ATAMALAR sayfasındaki TC Kimlik No
    var newName = getNewNameFromBelgeno(belgeNoData, tcValue, belgeNoTCColumnIndex, belgeNoNameColumnIndex);
    
    if (newName) {
      data[i][nameColumnIndex] = newName; // ATAMALAR sayfasındaki Ad Soyad'ı güncelle
    }
  }

  targetSheet.getRange(1, 1, data.length, data[0].length).setValues(data);
  Logger.log("ATAMALAR sayfasındaki Ad Soyad başarıyla güncellendi.");
}

function getNewNameFromBelgeno(belgeNoData, tcValue, tcColumnIndex, nameColumnIndex) {
  for (var i = 0; i < belgeNoData.length; i++) {
    if (belgeNoData[i][tcColumnIndex] == tcValue) {
      return belgeNoData[i][nameColumnIndex]; // Ad Soyad döndür
    }
  }
  return null; // Eşleşme yoksa null döndür
}
