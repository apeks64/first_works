function showSahaTablosu() {
  const html = HtmlService.createHtmlOutputFromFile("SahaTablosu")
    .setWidth(1000)
    .setHeight(700);
  SpreadsheetApp.getUi().showModalDialog(html, "FLIST Tablosu");
}

// Tüm benzersiz uzmanlar
function getUniqueEValues() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("FLIST");
  const values = sheet.getRange("E2:E" + sheet.getLastRow()).getValues().flat();
  return [...new Set(values.filter(v => v))];
}

// Tüm benzersiz bölgeler
function getUniqueFValues() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("FLIST");
  const values = sheet.getRange("F2:F" + sheet.getLastRow()).getValues().flat();
  return [...new Set(values.filter(v => v))];
}

// Filtreye göre tablo verisi
function getTableData(selectedEList, selectedFList) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("FLIST");
  const allData = sheet.getRange(2, 1, sheet.getLastRow() - 1, 21).getValues(); // A-U sütunları

  // Filtreleme
  const filtered = allData.filter(row =>
    (selectedEList.length === 0 || selectedEList.includes(row[4])) &&
    (selectedFList.length === 0 || selectedFList.includes(row[5]))
  );

  // Benzersiz listeler
  const uniqueOList = [...new Set(filtered.map(r => r[14]).filter(v => v))];
  const uniqueSList = [...new Set(filtered.map(r => r[18]).filter(v => v))];

  // Satırlar (Uzman ve Bölge sütunları çıkarıldı)
  return filtered.map((row, index) => {
    // Firma (A) sütununu ilk 21 karakter ile sınırla
    const firmaAdi = row[0] ? (row[0].length > 21 ? row[0].substring(0, 21) + "..." : row[0]) : "";

    return [
      firmaAdi,
      row[6],
      row[14],
      row[18],
      uniqueOList[index] || "",
      uniqueSList[index] || ""
    ];
  });
}

function exportToDocs(uzmanlar, bolgeler) {
  const data = getTableData(uzmanlar, bolgeler); // Güncel tablo verisini GAS tarafında al
  const doc = DocumentApp.create("FLIST Tablosu Çıktısı");
  const body = doc.getBody();

  // Başlık ekle
  body.appendParagraph("FLIST Tablosu Çıktısı").setHeading(DocumentApp.ParagraphHeading.HEADING1);

  // Filtre bilgisi ekle
  if (uzmanlar.length > 0) {
    body.appendParagraph("Seçilen Uzmanlar: " + uzmanlar.join(", "));
  }
  if (bolgeler.length > 0) {
    body.appendParagraph("Seçilen Bölgeler: " + bolgeler.join(", "));
  }

  // Tablo ekle
  const header = ["Firma (A)", "G", "O", "S", "O (Benzersiz)", "S (Benzersiz)"];
  data.unshift(header);
  const table = body.appendTable(data);
  
  // Belge URL'ini döndür
  return doc.getUrl();
}