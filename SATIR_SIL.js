/**
 * FilterAndDeleteUI HTML dosyasını içeren bir modal diyalog gösterir.
 */
function showFilterDialog() {
  const ui = SpreadsheetApp.getUi();
  const activeSheetName = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getName();
  const targetSheetName = "EK-2 BİLGİLER";

  if (activeSheetName !== targetSheetName) {
    return;
  }
  
  const uniqueOldFirms = getUniqueOldFirms();

  if (uniqueOldFirms.length === 0) {
    ui.alert("Eski Kayıt Uyarısı", "7 günden eski silinebilecek kayıt bulunmamaktadır.", ui.ButtonSet.OK);
    return;
  }
  
  const htmlOutput = HtmlService.createHtmlOutputFromFile("FilterAndDeleteUI")
      .setWidth(500)
      .setHeight(900);
  
  ui.showModalDialog(htmlOutput, "Eski Kayıtları Sil");
}

/**
 * EK-2 BİLGİLER sayfasından 7 günden eski kayıtları olan benzersiz firma adlarını alır.
 */
function getUniqueOldFirms() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("EK-2 BİLGİLER");
  
  if (!sheet) {
    return [];
  }
  
  const timeStampColumn = 1; 
  const firmNameColumn = 2; 
  const lastRow = sheet.getLastRow();
  
  if (lastRow <= 1) {
    return [];
  }
  
  const data = sheet.getRange(2, timeStampColumn, lastRow - 1, firmNameColumn).getValues();
  const uniqueOldFirms = new Set();
  const today = new Date();
  
  data.forEach(row => {
    const timeStamp = row[0];
    const firmName = row[1];
    
    if (timeStamp instanceof Date) {
      const diffInDays = (today.getTime() - timeStamp.getTime()) / (1000 * 60 * 60 * 24);
      if (diffInDays >= 7) {
        uniqueOldFirms.add(firmName);
      }
    }
  });
  
  return Array.from(uniqueOldFirms);
}

/**
 * Seçilen firmaların sadece 7 günden eski kayıtlarını siler.
 */
function deleteOldRecords(selectedFirms) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("EK-2 BİLGİLER");
  if (!sheet) {
    throw new Error("Hata: 'EK-2 BİLGİLER' adında bir sayfa bulunamadı.");
  }
  
  const timeStampColumn = 1; 
  const firmNameColumn = 2; 
  const lastRow = sheet.getLastRow();
  const data = sheet.getRange(2, timeStampColumn, lastRow - 1, firmNameColumn).getValues();
  const rowsToDelete = [];
  
  for (let i = data.length - 1; i >= 0; i--) {
    const firmName = data[i][1];
    const timeStamp = data[i][0];
    
    const today = new Date();
    const diffInDays = (today.getTime() - timeStamp.getTime()) / (1000 * 60 * 60 * 24);

    if (selectedFirms.includes(firmName) && diffInDays >= 7) {
      rowsToDelete.push(i + 2);
    }
  }

  rowsToDelete.forEach(row => {
    sheet.deleteRow(row);
  });
  
  return getUniqueOldFirms();
}

/**
 * Seçilen firmalara ait TÜM kayıtları (tarih fark etmeksizin) siler.
 */
function deleteAllRecords(selectedFirms) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("EK-2 BİLGİLER");
  if (!sheet) {
    throw new Error("Hata: 'EK-2 BİLGİLER' adında bir sayfa bulunamadı.");
  }
  
  const firmNameColumn = 2; 
  const lastRow = sheet.getLastRow();
  const data = sheet.getRange(2, firmNameColumn, lastRow - 1, 1).getValues();
  const rowsToDelete = [];
  
  for (let i = data.length - 1; i >= 0; i--) {
    const firmName = data[i][0];
    
    if (selectedFirms.includes(firmName)) {
      rowsToDelete.push(i + 2);
    }
  }

  rowsToDelete.forEach(row => {
    sheet.deleteRow(row);
  });
  
  return getUniqueOldFirms();
}

/**
 * Belirtilen klasör ID'sine sahip Google Drive klasörünü açar.
 */
function klasoruAc() {
  var folderId = '1TVjW_R8SSZsQ3qNcyA6X6ghlJNyN0niJ'; 
  var url = 'https://drive.google.com/drive/folders/' + folderId;

  var html = HtmlService.createHtmlOutput('<script>window.open("' + url + '", "_blank");google.script.host.close();</script>')
      .setWidth(100)
      .setHeight(10);
  
  SpreadsheetApp.getUi().showModalDialog(html, 'Klasör Açılıyor...');
}