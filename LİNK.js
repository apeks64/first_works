//BU KOD ÇALIŞTIRILDIĞINDA; MUAYENE FORMU İÇİN GEREKLİ BİLGİLER SHEETS DOSYASINDAKİ TÜM ÇALIŞMA SAYFALARINI
//LİNK SAYFASININ E SÜTUNUNA KISAYOL ŞEKLİNDE GETİRİR.
function sayfaKısayollarıOluştur() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = spreadsheet.getSheets();
  let linkSheet = spreadsheet.getSheetByName("LİNK");

  // "Linkler" sayfası yoksa oluştur
  if (!linkSheet) {
    linkSheet = spreadsheet.insertSheet("LİNK");
  } else {
    // BURADAKİ SATIRI YORUM SATIRI YAPTIK veya SİLDİK.
    // Artık Linkler sayfasındaki mevcut içeriği temizlemeyecek.
    // linkSheet.clearContents(); 
  }

  // Başlık satırını ekle
  linkSheet.getRange("E1").setValue("Sayfa Kısayolları");
  linkSheet.getRange("E1").setFontWeight("bold");

  let row = 2; // Kısayollar F2'den başlayacak

  sheets.forEach(sheet => {
    const sheetName = sheet.getName();
    // "Linkler" sayfasının kendisine kısayol oluşturmaktan kaçın
    if (sheetName !== "Linkler") {
      const sheetId = sheet.getSheetId();
      const spreadsheetUrl = spreadsheet.getUrl();
      // Belirli sayfaya doğrudan bağlantı oluşturma
      const sheetUrl = `${spreadsheetUrl}#gid=${sheetId}`;
      
      // HYPERLINK formülü ile kısayol oluştur
      const hyperlinkFormula = `=HYPERLINK("${sheetUrl}"; "${sheetName}")`; 
      linkSheet.getRange(`F${row}`).setFormula(hyperlinkFormula);
      row++;
    }
  });

  // F sütununu otomatik boyutlandır
  linkSheet.autoResizeColumn(6); 
}