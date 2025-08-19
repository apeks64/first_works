function karsilastirSGK() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var atamalarSheet = ss.getSheetByName("ATAMALAR");
  var yedekSheet = ss.getSheetByName("FLIST");

  if (!atamalarSheet || !yedekSheet) {
    Logger.log("Sayfalardan biri eksik! Lütfen sayfa adlarını kontrol edin.");
    return;
  }

  // ATAMALAR sayfasındaki SGK No ve işyeri bilgileri
  var atamalarData = atamalarSheet.getRange("Q2:Q").getValues().flat(); // ATAMALAR SGK No
  var atamalarIsyeri = atamalarSheet.getRange("R2:R").getValues().flat(); // ATAMALAR İşyeri Bilgisi

  // YEDEK sayfasındaki SGK No ve işyeri bilgileri
  var yedekData = yedekSheet.getRange("D2:D").getValues().flat(); // YEDEK SGK No
  var yedekIsyeri = yedekSheet.getRange("A2:A").getValues().flat(); // YEDEK İşyeri Bilgisi

  // Her iki sayfada da olmayan SGK No'lar için farkları tutacağımız dizi
  var farklar = [];
  
  // Her iki sayfada da SGK No'ları karşılaştırmak için set kullanarak kontrol ediyoruz
  var atamalarMap = new Map();
  var yedekMap = new Map();

  // ATAMALAR sayfasındaki verileri map'e ekliyoruz
  for (var i = 0; i < atamalarData.length; i++) {
    var sgkAtama = String(atamalarData[i]).trim(); // SGK No'yu string'e çevirip trimliyoruz
    var isyeriAtama = atamalarIsyeri[i] ? atamalarIsyeri[i].trim() : "Bilinmiyor"; // İşyeri Bilgisi
    if (sgkAtama) {
      atamalarMap.set(sgkAtama, { satir: i + 2, isyeri: isyeriAtama }); // SGK No ve işyeri bilgisi
    }
  }

  // YEDEK sayfasındaki verileri map'e ekliyoruz
  for (var i = 0; i < yedekData.length; i++) {
    var sgkYedek = String(yedekData[i]).trim(); // SGK No'yu string'e çevirip trimliyoruz
    var isyeriYedek = yedekIsyeri[i] ? yedekIsyeri[i].trim() : "Bilinmiyor"; // İşyeri Bilgisi
    if (sgkYedek) {
      yedekMap.set(sgkYedek, { satir: i + 2, isyeri: isyeriYedek }); // SGK No ve işyeri bilgisi
    }
  }

  // ATAMALAR sayfasındaki her SGK No'yu kontrol et
  for (var [sgkAtama, bilgilerAtama] of atamalarMap) {
    if (!yedekMap.has(sgkAtama)) {
      // Eğer SGK No YEDEK sayfasında yoksa, farkı yaz
      farklar.push([sgkAtama, "Yedek Sayfasında Yok", bilgilerAtama.isyeri]);
    }
  }

  // YEDEK sayfasındaki her SGK No'yu kontrol et
  for (var [sgkYedek, bilgilerYedek] of yedekMap) {
    if (!atamalarMap.has(sgkYedek)) {
      // Eğer SGK No ATAMALAR sayfasında yoksa, farkı yaz
      farklar.push([sgkYedek, "Atamalar Sayfasında Yok", bilgilerYedek.isyeri]);
    }
  }

  if (farklar.length === 0) {
    Logger.log("Farklı SGK No bulunamadı.");
    return;
  }

  // Sonuçları yeni bir sayfaya yaz
  var farkSheet = ss.getSheetByName("SGK_FARKLAR") || ss.insertSheet("SGK_FARKLAR");
  farkSheet.clear(); // Önceki sonuçları temizle
  farkSheet.getRange(1, 1).setValue("SGK No");
  farkSheet.getRange(1, 2).setValue("Farklı Sayfa");
  farkSheet.getRange(1, 3).setValue("İşyeri Bilgisi");
  farkSheet.getRange(2, 1, farklar.length, 3).setValues(farklar);

  Logger.log("Farklı SGK No'lar SGK_FARKLAR sayfasına yazıldı.");
}
