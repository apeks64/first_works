//Bu kod, bir Google E-Tablolar'da AQ (43.) sütunundaki tüm değerleri "Aktif" olarak günceller, ancak zaten "Pasif" yazan hücrelere dokunmaz.
function aqSutununuGuncelle() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const lastRow = sheet.getLastRow();
  const aqValues = sheet.getRange(2, 0, lastRow - 1).getValues(); // AQ sütunu = 43. sütun

  for (let i = 0; i < aqValues.length; i++) {
    const currentValue = aqValues[i][0];
    if (currentValue !== "PASİF") {
      sheet.getRange(i + 2, 0).setValue("AKTİF");
    }
  }
}

// Bu işlev, A sütunundaki (1. sütun) tüm satırları "AKTİF" olarak günceller.
// "PASİF" veya boş olmasına bakmaksızın, tüm değerleri "AKTİF" yapar.
function tumSatirlariAktifYap() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('FLIST');
  
  if (!sheet) {
    Logger.log("Hata: 'FLIST' adında bir sayfa bulunamadı.");
    return;
  }
  
  const lastRow = sheet.getLastRow();
  // A sütunundaki tüm verileri (A2'den son satıra kadar) alır.
  const range = sheet.getRange(2, 1, lastRow - 1, 1);
  const values = range.getValues();
  
  // Her satırdaki değeri "AKTİF" olarak değiştirir.
  const newValues = values.map(() => ['AKTİF']);
  
  // Güncellenen değerleri tek bir işlemle sayfaya yazar.
  range.setValues(newValues);
  
  SpreadsheetApp.getUi().alert("Tüm satırlar başarıyla 'AKTİF' yapıldı.");
}