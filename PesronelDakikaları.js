/**
 * ATAMALAR sayfasından kişi bazlı toplam süreleri döndürür
 */
function getPersonelSureler() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('ATAMALAR');
  if (!sheet) return [];

  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const fIndex = headers.indexOf('Görevlendirilen Kişi Ad Soyad'); // F sütunu
  const jIndex = headers.indexOf('Çalışma Süresi'); // J sütunu

  if (fIndex === -1 || jIndex === -1) {
    throw new Error("Gerekli sütunlar bulunamadı (F: Ad Soyad, J: Çalışma Süresi).");
  }

  const toplamlari = {};

  for (let i = 1; i < data.length; i++) {
    const adSoyad = String(data[i][fIndex]).trim();
    const sure = Number(data[i][jIndex]) || 0;
    if (!adSoyad) continue;
    if (!toplamlari[adSoyad]) {
      toplamlari[adSoyad] = 0;
    }
    toplamlari[adSoyad] += sure;
  }

  // Objeyi diziye çevir + büyükten küçüğe sırala
  const sonuc = Object.keys(toplamlari)
    .map(adSoyad => {
      return { adSoyad: adSoyad, toplamSure: toplamlari[adSoyad] };
    })
    .sort((a, b) => b.toplamSure - a.toplamSure);

  return sonuc;
}

/**
 * Modal pencereyi açar
 */
function showPersonelDurumDialog() {
  const html = HtmlService.createHtmlOutputFromFile('PersonelDurum')
    .setWidth(450)
    .setHeight(999);
  SpreadsheetApp.getUi().showModalDialog(html, 'Personel Dakika Durumu');
}
