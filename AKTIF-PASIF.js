//Bu kod, bir Google E-Tablolar'da AQ (43.) sütunundaki tüm değerleri "Aktif" olarak günceller, ancak zaten "Pasif" yazan hücrelere dokunmaz.
function aqSutununuGuncelle() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const lastRow = sheet.getLastRow();
  const aqValues = sheet.getRange(2, 43, lastRow - 1).getValues(); // AQ sütunu = 43. sütun

  for (let i = 0; i < aqValues.length; i++) {
    const currentValue = aqValues[i][0];
    if (currentValue !== "PASİF") {
      sheet.getRange(i + 2, 43).setValue("AKTİF");
    }
  }
}
//Google E-Tablolar'da bir satırı pasif olarak işaretlediğinizde otomatik olarak gizleyecek bir Google Apps Script kodu
// Google E-Tablolar'da bir satırı pasif olarak işaretlediğinizde otomatik olarak gizleyecek bir Google Apps Script kodu
function onEdit(e) {
  // Düzenlemenin yapıldığı sayfa, satır ve sütunu al
  const sheet = e.range.getSheet();
  const row = e.range.getRow();
  const col = e.range.getColumn();

  // İşlemi sadece 'FLIST' sayfasında ve AV sütununda (48. sütun) yapmak için kontrol et
  if (sheet.getName() === 'FLIST' && col === 48) {
    // Düzenlenen hücrenin değerini al
    const cellValue = e.range.getValue();

    // Eğer hücredeki değer 'PASİF' ise satırı gizle
    if (cellValue === 'PASİF') {
      sheet.hideRows(row);
    }
    // Eğer hücre boşsa veya "PASİF" dışında bir şeyse, otomatik olarak "AKTİF" yap
    else if (cellValue !== 'AKTİF') {
      sheet.getRange(row, col).setValue('AKTİF');
    }
  }
}

// Bu işlev, AV sütunundaki tüm boş veya "PASİF" olmayan hücreleri "AKTİF" yapar.
// Manuel olarak çalıştırılabilir.
function tumSatirlariAktifYap() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('FLIST');
  const lastRow = sheet.getLastRow();
  const avValues = sheet.getRange(1, 48, lastRow, 1).getValues(); // AV sütunu (48)

  for (let i = 0; i < avValues.length; i++) {
    const currentValue = avValues[i][0];
    // Eğer hücre boşsa veya "PASİF" değilse "AKTİF" yap
    if (currentValue === "" || currentValue !== 'PASİF') {
      sheet.getRange(i + 1, 48).setValue('AKTİF');
    }
  }
}