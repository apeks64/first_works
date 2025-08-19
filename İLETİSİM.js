function veriAktarFlistIletisim() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var flistSheet = ss.getSheetByName("FLIST");
  var iletisimSheet = ss.getSheetByName("İLETİŞİM");

  if (!flistSheet || !iletisimSheet) {
    Logger.log("FLIST veya İLETİŞİM sayfası bulunamadı!");
    return;
  }

  // FLIST ve İLETİŞİM sayfalarındaki verileri al
  var flistData = flistSheet.getDataRange().getValues();
  var iletisimData = iletisimSheet.getDataRange().getValues();

  // Sütun başlıklarını belirleme (İLETİŞİM sayfası)
  var headers = iletisimData[0];
  var adresIndex = headers.indexOf("Adres");
  var epostaIndex = headers.indexOf("E-Posta");
  var uzmanIndex = headers.indexOf("Uzman");
  var telefonIndex = headers.indexOf("Telefon");
  var yetkiliIndex = headers.indexOf("Yetkili");

  if (adresIndex === -1 || epostaIndex === -1 || uzmanIndex === -1 || telefonIndex === -1 || yetkiliIndex === -1) {
    Logger.log("İLETİŞİM sayfasında gerekli sütunlardan biri eksik!");
    return;
  }

  // FLIST sayfasındaki sütun indexleri (A, ADRESI, E_POSTA, UZMAN, TELEFON, YETKILI)
  var flistAdIndex = 0;
  var flistAdresIndex = 3;
  var flistEpostaIndex = 8;
  var flistUzmanIndex = 2;
  var flistTelefonIndex = 6;
  var flistYetkiliIndex = 7;

  // İLETİŞİM verilerini eşleşme için nesneye dönüştür
  var iletisimMap = {};
  for (var i = 1; i < iletisimData.length; i++) {
    var isim = iletisimData[i][0].trim().toUpperCase(); // A sütunu (isim/unvan)
    if (isim.length < 5) continue;

    iletisimMap[isim] = {
      adres: iletisimData[i][adresIndex],
      eposta: iletisimData[i][epostaIndex],
      uzman: iletisimData[i][uzmanIndex],
      telefon: iletisimData[i][telefonIndex],
      yetkili: iletisimData[i][yetkiliIndex]
    };
  }

  // FLIST sayfasında verileri güncelle
  var updates = [];
  for (var j = 1; j < flistData.length; j++) {
    var flistIsim = flistData[j][flistAdIndex].trim().toUpperCase();
    if (flistIsim.length < 5) continue;

    for (var key in iletisimMap) {
      if (flistIsim.startsWith(key.substring(0, 5))) { // En az 5 karakter eşleşmeli
        updates.push({
          row: j + 1,
          adres: iletisimMap[key].adres,
          eposta: iletisimMap[key].eposta,
          uzman: iletisimMap[key].uzman,
          telefon: iletisimMap[key].telefon,
          yetkili: iletisimMap[key].yetkili
        });
        break;
      }
    }
  }

  // Güncellemeleri FLIST sayfasına yaz
  updates.forEach(update => {
    flistSheet.getRange(update.row, flistAdresIndex + 1).setValue(update.adres);
    flistSheet.getRange(update.row, flistEpostaIndex + 1).setValue(update.eposta);
    flistSheet.getRange(update.row, flistUzmanIndex + 1).setValue(update.uzman);
    flistSheet.getRange(update.row, flistTelefonIndex + 1).setValue(update.telefon);
    flistSheet.getRange(update.row, flistYetkiliIndex + 1).setValue(update.yetkili);
  });

  Logger.log("Veriler başarıyla aktarıldı!");
}
