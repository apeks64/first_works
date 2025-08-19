function aqSutununuGuncelle() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const lastRow = sheet.getLastRow();
  const aqValues = sheet.getRange(2, 43, lastRow - 1).getValues(); // AQ sütunu = 43. sütun

  for (let i = 0; i < aqValues.length; i++) {
    const currentValue = aqValues[i][0];
    if (currentValue !== "Pasif") {
      sheet.getRange(i + 2, 43).setValue("Aktif");
    }
  }
}
