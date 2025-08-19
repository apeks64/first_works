function showSearchModal() {
  const html = HtmlService.createHtmlOutputFromFile('search_modal')
      .setWidth(600)
      .setHeight(1000)
      .setTitle('Personel Arama ve Yönlendirme');
  SpreadsheetApp.getUi().showModalDialog(html, html.getTitle());
}

function searchEmployees(isYeriAdi, adi, soyadi) {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadsheet.getSheetByName('EK-2_DB');
  const dataRange = sheet.getDataRange();
  const values = dataRange.getValues();
  
  const results = [];
  
  const cleanedIsYeriAdi = isYeriAdi.toLowerCase().replace(/[.,\/#!$%\^&\*;:{}=\-_`~()]/g,"");
  const cleanedAdi = adi.toLowerCase().replace(/[.,\/#!$%\^&\*;:{}=\-_`~()]/g,"");
  const cleanedSoyadi = soyadi.toLowerCase().replace(/[.,\/#!$%\^&\*;:{}=\-_`~()]/g,"");
  
  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    
    const cellIsYeriAdi = row[1] ? row[1].toString().toLowerCase().replace(/[.,\/#!$%\^&\*;:{}=\-_`~()]/g,"") : '';
    const cellAdi = row[2] ? row[2].toString().toLowerCase().replace(/[.,\/#!$%\^&\*;:{}=\-_`~()]/g,"") : '';
    const cellSoyadi = row[3] ? row[3].toString().toLowerCase().replace(/[.,\/#!$%\^&\*;:{}=\-_`~()]/g,"") : '';

    const matchIsYeriAdi = cleanedIsYeriAdi ? cellIsYeriAdi.includes(cleanedIsYeriAdi) : true;
    const matchAdi = cleanedAdi ? cellAdi.includes(cleanedAdi) : true;
    const matchSoyadi = cleanedSoyadi ? cellSoyadi.includes(cleanedSoyadi) : true;
    
    if (matchIsYeriAdi && matchAdi && matchSoyadi) {
      // Satır numarasını (i + 1) ve sadece B, C, D sütunlarını gönder
      const displayData = [i + 1, row[1], row[2], row[3]];
      results.push(displayData);
    }
  }
  
  return results;
}

// Bu fonksiyon, kullanıcıyı belirli bir satıra yönlendirir ve o satırı seçer.
function goToRow(rowIndex) {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadsheet.getSheetByName('EK-2_DB');
  
  // Belirtilen sayfaya geçiş yap
  sheet.activate();
  
  // Belirtilen satırdaki hücreyi seçerek görünür hale getir
  sheet.getRange(rowIndex, 1).activate();
  
  // Modal pencereyi kapat
  SpreadsheetApp.getUi().showSidebar(null);
}