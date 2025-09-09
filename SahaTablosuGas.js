function showSahaTablosu() {
  const html = HtmlService.createHtmlOutputFromFile("SahaTablosu")
    .setWidth(1000)
    .setHeight(700);
  SpreadsheetApp.getUi().showModalDialog(html, "FLIST Tablosu");
}

// Tüm benzersiz uzmanlar
function getUniqueEValues() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("FLIST");
  // F sütunu artık G sütunu oldu, bu nedenle indeks F'den G'ye kaydırıldı.
  const values = sheet.getRange("F2:F" + sheet.getLastRow()).getValues().flat();
  return [...new Set(values.filter(v => v))];
}

// Tüm benzersiz bölgeler
function getUniqueFValues() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("FLIST");
  // G sütunu artık H sütunu oldu, bu nedenle indeks G'den H'ye kaydırıldı.
  const values = sheet.getRange("G2:G" + sheet.getLastRow()).getValues().flat();
  return [...new Set(values.filter(v => v))];
}

// Filtreye göre tablo verisi
function getTableData(selectedEList, selectedFList) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("FLIST");
  // Yeni sütun eklenmesiyle, A-U sütunları artık A-V sütunlarıdır.
  // Bu yüzden aralığı 22 sütun (1'den 22'ye kadar) olarak güncelledik.
  const allData = sheet.getRange(2, 1, sheet.getLastRow() - 1, 22).getValues();

  // Filtreleme
  const filtered = allData.filter(row =>
    // 'Uzman' sütunu (önceden F) artık G sütunu, bu yüzden indeksi 4'ten 5'e güncelledik.
    // 'Bölge' sütunu (önceden G) artık H sütunu, bu yüzden indeksi 5'ten 6'ya güncelledik.
    (selectedEList.length === 0 || selectedEList.includes(row[5])) &&
    (selectedFList.length === 0 || selectedFList.includes(row[6]))
  );

  // Benzersiz listeler
  // O sütunu (önceden 15. indeks) artık P sütunu, indeksi 14'ten 15'e güncelledik.
  const uniqueOList = [...new Set(filtered.map(r => r[15]).filter(v => v))];
  // S sütunu (önceden 19. indeks) artık T sütunu, indeksi 18'den 19'a güncelledik.
  const uniqueSList = [...new Set(filtered.map(r => r[19]).filter(v => v))];

  // Satırlar (Uzman ve Bölge sütunları çıkarıldı)
  return filtered.map((row, index) => {
    // Burada, çıktı tablosunun ilk sütunu artık B sütunundaki veriler olacak.
    // B sütunu, 1. indeks (0'dan başladığı için).
    const ikinciSutun = row[1] ? row[1] : "";

    return [
      ikinciSutun, // Artık tablodaki ilk veri B sütunundan gelecek.
      // G sütunu (önceden 7. indeks) artık H sütunu, indeksi 6'dan 7'ye güncelledik.
      row[7],
      // O sütunu (önceden 15. indeks) artık P sütunu, indeksi 14'ten 15'e güncelledik.
      row[15],
      // S sütunu (önceden 19. indeks) artık T sütunu, indeksi 18'den 19'a güncelledik.
      row[19],
      uniqueOList[index] || "",
      uniqueSList[index] || ""
    ];
  });
}

function exportToDocs(uzmanlar, bolgeler) {
  const data = getTableData(uzmanlar, bolgeler);
  const doc = DocumentApp.create("FLIST Tablosu Çıktısı");
  const body = doc.getBody();

  body.appendParagraph("FLIST Tablosu Çıktısı").setHeading(DocumentApp.ParagraphHeading.HEADING1);

  if (uzmanlar.length > 0) {
    body.appendParagraph("Seçilen Uzmanlar: " + uzmanlar.join(", "));
  }
  if (bolgeler.length > 0) {
    body.appendParagraph("Seçilen Bölgeler: " + bolgeler.join(", "));
  }

  // Tablo başlıkları B sütununa göre güncellendi.
  const header = ["FIRMA", "TS", "IGU", "IH", "IGU (Benzersiz)", "IH (Benzersiz)"];
  data.unshift(header);
  const table = body.appendTable(data);

  return doc.getUrl();
}