function showFilterDialog() {
  var html = HtmlService.createHtmlOutputFromFile("FilterAndDeleteUI")
      .setWidth(400)
      .setHeight(500);
  SpreadsheetApp.getUi().showModalDialog(html, "Veri Filtreleme ve Silme");
}

function getUniqueFirmalar() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("EK-2 BİLGİLER");
  var firmaSutunIndex = 2; // B sütunu
  var data = sheet.getRange(2, firmaSutunIndex, sheet.getLastRow() - 1, 1).getValues();
  var uniqueFirms = [...new Set(data.flat().filter(String))]; // Boş hücreleri at ve benzersizleri al
  return uniqueFirms;
}

function deleteSelectedFirms(selectedFirms) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("EK-2 BİLGİLER");
  var firmaSutunIndex = 2; // B sütunu
  var lastRow = sheet.getLastRow();
  var data = sheet.getRange(2, firmaSutunIndex, lastRow - 1, 1).getValues();
  var rowsToDelete = [];

  for (var i = 0; i < data.length; i++) {
    if (selectedFirms.includes(data[i][0])) {
      rowsToDelete.push(i + 2);
    }
  }

  if (rowsToDelete.length === 0) {
    SpreadsheetApp.getUi().alert("Seçili firmalar bulunamadı veya zaten silindi.");
    return;
  }

  // Satırları en sondan başlayarak sil
  for (var j = rowsToDelete.length - 1; j >= 0; j--) {
    sheet.deleteRow(rowsToDelete[j]);
  }
}
