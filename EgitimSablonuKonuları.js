function showEgitimPano() {
  const html = HtmlService.createHtmlOutputFromFile("konular")
    .setWidth(800)
    .setHeight(600);
  SpreadsheetApp.getUi().showModalDialog(html, "Eğitim Seçim ve Pano Görüntüle");
}

function getEgitimVerileri() {
  const sheet = SpreadsheetApp.getActive().getSheetByName("EGITIM");
  const data = sheet.getDataRange().getValues();
  
  const headers = data[0]; // İlk satır başlıklar
  const egitimVerileri = {};

  for (let col = 0; col < headers.length; col++) {
    const konu = headers[col];
    if (!konu) continue;

    const detaylar = [];
    for (let row = 1; row < data.length; row++) {
      if (data[row][col]) {
        detaylar.push(data[row][col]);
      }
    }

    egitimVerileri[konu] = detaylar;
  }

  // Konu sırasını korumak için ayrıca sıralı başlık listesi döndürelim
  return {
    siraliKonuListesi: headers.filter(h => h), // boş olmayan başlıklar
    detaylar: egitimVerileri
  };
}
function getEgitimVerileriSirali() {
  const sheet = SpreadsheetApp.getActive().getSheetByName("EGITIM");
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const allData = sheet.getRange(2, 1, sheet.getLastRow() - 1, headers.length).getValues();

  let veriler = {};
  headers.forEach((baslik, i) => {
    const detaylar = allData.map(row => row[i]).filter(cell => cell);
    veriler[baslik] = detaylar;
  });

  return { veriler, konuSirasi: headers };
}
