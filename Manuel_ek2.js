// --- Belge oluşturma ana fonksiyonu ---
function generateDocument(rowNumber, silentMode = false) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('EK-2 BİLGİLER');
    if (!sheet) throw new Error('Tablo bulunamadı.');

    if (!rowNumber) rowNumber = sheet.getLastRow();

    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const data = sheet.getRange(rowNumber, 1, 1, sheet.getLastColumn()).getValues()[0];
    const responseData = headers.reduce((obj, header, i) => ({ ...obj, [header]: data[i] }), {});
    Logger.log('responseData:', responseData);

    const workplaceName = responseData['İş Yerinin Adı'] ? responseData['İş Yerinin Adı'].substring(0, 13) : 'Bilinmiyor';
    const firstName = responseData['Adı'] || 'Bilinmiyor';
    const lastName = responseData['Soyadı'] || 'Bilinmiyor';

    const fileName = `${workplaceName} ${firstName} ${lastName} - Muayene Formu`;

    const destinationFolder = DriveApp.getFolderById('1TVjW_R8SSZsQ3qNcyA6X6ghlJNyN0niJ');
    const template = DriveApp.getFileById('1jnytmZfHhHi9bkgBCX5uMRLNFTffmGKqTurMfwKzShE');
    const newDoc = template.makeCopy(fileName, destinationFolder);
    const doc = DocumentApp.openById(newDoc.getId());
    const body = doc.getBody();

    fillPlaceholders(body, responseData);
    addFLISTData(responseData, body);
    addSignatureImage(responseData, body);

    doc.saveAndClose();
    Logger.log('Belge başarıyla oluşturuldu!');

  } catch (e) {
    Logger.log('Hata: ' + e.message);
    if (!silentMode) {
      SpreadsheetApp.getUi().alert('Hata oluştu: ' + e.message);
    }
  }
}

// --- Tek bir belge oluştur (seçili satır) ---
function createSingleDocument() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('EK-2 BİLGİLER');
  const activeCell = sheet.getActiveCell();
  const rowNumber = activeCell.getRow();
  generateDocument(rowNumber, true); // Sessiz modda çalıştır
  SpreadsheetApp.getUi().alert('✅ Belge başarıyla oluşturuldu!');
}

// --- Birden fazla belge oluştur (seçili satırlar) ---
function generateDocumentsFromSelection() {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('EK-2 BİLGİLER');
    if (!sheet) throw new Error('Tablo bulunamadı.');

    const selection = SpreadsheetApp.getActiveSpreadsheet().getActiveRange();
    if (!selection) throw new Error('Herhangi bir seçim yapılmadı.');
    const startRow = selection.getRow();
    const numRows = selection.getNumRows();

    const ui = SpreadsheetApp.getUi();

    let createdCount = 0;
    const progressText = ui.showModalDialog(HtmlService.createHtmlOutput('<div id="progress">Belge oluşturuluyor...</div><script>function updateProgress(text) { document.getElementById("progress").innerText = text; }</script>').setWidth(300).setHeight(50), 'Belge Oluşturuluyor');

    for (let i = 0; i < numRows; i++) {
      const currentRow = startRow + i;
      const rowData = sheet.getRange(currentRow, 1, 1, sheet.getLastColumn()).getValues()[0];
      if (rowData.some(cell => cell !== "")) {
        try {
          // İlerleme durumunu güncelle
          SpreadsheetApp.getUi().showModalDialog(HtmlService.createHtmlOutput(`<div id="progress">${i + 1}/${numRows} belge oluşturuluyor...</div>`).setWidth(300).setHeight(50), 'Belge Oluşturuluyor');

          generateDocument(currentRow, true); // Sessiz çalıştırıyoruz
          createdCount++;
        } catch (e) {
          Logger.log(`Satır ${currentRow} için belge oluşturulurken hata: ` + e.message);
        }
      }
    }

    SpreadsheetApp.getUi().alert(`✅ İşlem tamamlandı! ${createdCount} adet belge oluşturuldu.`);
  } catch (e) {
    Logger.log('Hata: ' + e.message);
    SpreadsheetApp.getUi().alert('Hata oluştu: ' + e.message);
  }
}