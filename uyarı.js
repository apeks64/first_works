function onSelectionChange(e) {
  // The sheet where you want the alert to appear
  const alertSheetName = 'EK-2 BİLGİLER';

  // Check if the active sheet is the one you want to monitor
  const activeSheet = e.range.getSheet();
  if (activeSheet.getName() === alertSheetName) {
    // Display the alert message
    SpreadsheetApp.getUi().alert('Lütfen Ek-2 İşlemleri altında bulunan eski kayıtları sil menüsünden işi biten firmaları siliniz.');
  }
}