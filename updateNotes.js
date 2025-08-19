function updateNotesAdvancedMatching() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ek2Sheet = ss.getSheetByName('EK-2 BİLGİLER'); // EK-2 BİLGİLER sayfanızın adını kontrol edin
  var flistSheet = ss.getSheetByName('FLIST');     // FLIST sayfanızın adını kontrol edin (Görselde FLIST_2 yazıyor, bu ismi kontrol edin)

  if (!ek2Sheet || !flistSheet) {
    Logger.log('Sayfalardan biri bulunamadı. Lütfen sayfa adlarını kontrol edin.');
    SpreadsheetApp.getUi().alert('Hata', 'Sayfalardan biri bulunamadı. Lütfen "EK-2 BİLGİLER" ve "FLIST_2" sayfa adlarını kontrol edin.', SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }

  var ek2Data = ek2Sheet.getRange('B:B').getValues(); // EK-2 BİLGİLER B sütunundaki tüm verileri al
  var flistData = flistSheet.getRange('A:E').getValues(); // FLIST sayfasının A'dan E'ye kadar olan verilerini al

  // Metni normalleştiren yardımcı fonksiyon
  function normalizeText(text) {
    if (!text) return '';
    return text.toString()
               .toLowerCase()
               .replace(/i/g, 'i') // İ => i
               .replace(/ı/g, 'i') // I => i
               .replace(/ö/g, 'o') // Ö => o
               .replace(/ü/g, 'u') // Ü => u
               .replace(/ç/g, 'c') // Ç => c
               .replace(/ş/g, 's') // Ş => s
               .replace(/ğ/g, 'g'); // Ğ => g
               // Not: Metinlerdeki boşlukları ve diğer özel karakterleri kaldırmak istemediğiniz varsayılmıştır.
               // Eğer boşluklar ve diğer özel karakterler de eşleşmeyi etkiliyorsa,
               // .replace(/\s/g, '') // Tüm boşlukları kaldır
               // .replace(/[^a-z0-9]/g, '') // Sadece harf ve rakamları bırak (Türkçe karakterleri dönüştürdükten sonra)
               // gibi eklemeler yapabilirsiniz. Şimdilik sadece harflere odaklanalım.
  }

  // FLIST sayfasındaki UNVANI (A) ve UZMAN (E) verilerini hazırla
  var flistEntries = []; // Normalize edilmiş FLIST ünvanlarını tutar
  var flistNormalizedUnvanToExpertsMap = {}; // Anahtar: Normalize edilmiş FLIST UNVAN, Değer: O unvanla eşleşen TEKİL uzman(lar) listesi
  
  for (var i = 0; i < flistData.length; i++) {
    var unvan = flistData[i][0];
    var uzman = flistData[i][4];
    
    if (unvan && unvan.toString().trim() !== '') {
      var trimmedUnvan = unvan.toString().trim();
      var normalizedUnvan = normalizeText(trimmedUnvan);
      
      if (!flistNormalizedUnvanToExpertsMap.hasOwnProperty(normalizedUnvan)) {
          flistNormalizedUnvanToExpertsMap[normalizedUnvan] = new Set(); // Set kullanarak benzersiz uzmanları tut
      }
      flistNormalizedUnvanToExpertsMap[normalizedUnvan].add(uzman); // Uzmanı Set'e ekle
      flistEntries.push(normalizedUnvan); // Eşleşme araması için normalize edilmiş listeyi kullan
    }
  }

  // EK-2 BİLGİLER sayfasındaki B sütununu kontrol et ve not ekle
  for (var i = 0; i < ek2Data.length; i++) {
    var isYeriAdi = ek2Data[i][0]; // B sütunundaki İş Yerinin Adı
    var targetCell = ek2Sheet.getRange(i + 1, 2); // B sütunundaki ilgili hücre

    if (isYeriAdi && isYeriAdi.toString().trim() !== '') {
      var searchKey = isYeriAdi.toString().trim();
      var normalizedSearchKey = normalizeText(searchKey);
      var foundExpert = null;
      
      var parts = normalizedSearchKey.split(' ');
      var currentSearchPrefix = '';

      for (var k = 0; k < parts.length; k++) {
        currentSearchPrefix += (k > 0 ? ' ' : '') + parts[k];

        var currentMatches = []; // Bu önekle eşleşen FLIST ünvanlarının normalize edilmiş halleri
        for (var j = 0; j < flistEntries.length; j++) {
          var flistNormalizedUnvan = flistEntries[j];
          if (flistNormalizedUnvan.startsWith(currentSearchPrefix)) {
            currentMatches.push(flistNormalizedUnvan);
          }
        }

        // Eğer bu önekle eşleşen FLIST ünvanları varsa
        if (currentMatches.length > 0) {
          var distinctExpertsForThisPrefix = new Set();
          currentMatches.forEach(matchedNormalizedUnvan => {
              // Eşleşen her normalize edilmiş ünvan için, ilgili uzmanları al
              var expertsSet = flistNormalizedUnvanToExpertsMap[matchedNormalizedUnvan];
              if (expertsSet) {
                  expertsSet.forEach(expert => distinctExpertsForThisPrefix.add(expert));
              }
          });

          // Eğer bu önekle eşleşen ve TEKİL bir uzman bulunduysa
          if (distinctExpertsForThisPrefix.size === 1) {
            foundExpert = distinctExpertsForThisPrefix.values().next().value; // İlk ve tek elemanı al
            break; // Tekil bir uzman bulduk, daha fazla aramaya gerek yok
          } else if (k === parts.length - 1) { // Son kelimeye ulaştık ve hala birden fazla uzman varsa
            foundExpert = null; // Kesin eşleşme yok
            break;
          }
          // currentMatches.length > 1 ve distinctExpertsForThisPrefix.size > 1 ise,
          // veya currentMatches.length > 1 ve distinctExpertsForThisPrefix.size = 0 ise
          // bir sonraki kelimeyi ekleyerek öneki uzatmaya devam edeceğiz.
        } else { // Hiç eşleşme bulunamadıysa, daha uzun öneklerle denemek anlamsız.
          foundExpert = null;
          break;
        }
      }
      
      // Notu ekle veya kaldır
      if (foundExpert) {
        targetCell.setNote(foundExpert);
      } else {
        targetCell.clearNote(); // Eşleşme bulunamazsa veya birden fazla uzman varsa notu kaldır
      }

    } else {
      targetCell.clearNote(); // İş Yerinin Adı boşsa notu kaldır
    }
  }
}