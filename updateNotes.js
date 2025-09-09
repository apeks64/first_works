/**
 * EK-2 BİLGİLER sayfasındaki firmaların adıyla FLIST sayfasındaki uzmanları eşleştirir
 * ve ilgili hücreye not olarak yazar. Bu versiyon, nota Eğitim Şekli (AT sütunu) verisini de ekler.
 */
function updateNotesAdvancedMatching() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ek2Sheet = ss.getSheetByName('EK-2 BİLGİLER');
  var flistSheet = ss.getSheetByName('FLIST'); // Lütfen dosyanızdaki sayfa adını kontrol edin

  if (!ek2Sheet || !flistSheet) {
    Logger.log('Sayfalardan biri bulunamadı. Lütfen sayfa adlarını kontrol edin.');
    SpreadsheetApp.getUi().alert('Hata', 'Sayfalardan biri bulunamadı. Lütfen "EK-2 BİLGİLER" ve "FLIST" sayfa adlarını kontrol edin.', SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }

  var ek2Data = ek2Sheet.getRange('B:B').getValues();
  
  // DEĞİŞİKLİK 1: FLIST sayfasından A'dan AT'ye kadar olan verileri al.
  // AT sütunu 20. sütundur ve indeks 45'a denk gelir.
  var flistData = flistSheet.getRange('A:AT').getValues();

  // Metni normalleştiren yardımcı fonksiyon
  function normalizeText(text) {
    if (!text) return '';
    return text.toString()
                .toLowerCase()
                .replace(/i/g, 'i') 
                .replace(/ı/g, 'i')
                .replace(/ö/g, 'o')
                .replace(/ü/g, 'u')
                .replace(/ç/g, 'c')
                .replace(/ş/g, 's')
                .replace(/ğ/g, 'g');
  }

  // FLIST sayfasındaki B, F ve AT sütunlarını işle
  var flistNormalizedUnvanToExpertsMap = {};
  
  for (var i = 0; i < flistData.length; i++) {
    var unvan = flistData[i][1]; // B sütunu
    var uzman = flistData[i][5]; // F sütunu
    var egitimSekli = flistData[i][45]; // AT sütunu (indeks 45)
    
    if (unvan && unvan.toString().trim() !== '') {
      var trimmedUnvan = unvan.toString().trim();
      var normalizedUnvan = normalizeText(trimmedUnvan);
      
      // Her uzman için hem adını hem de eğitim şeklini tutacak bir nesne oluştur
      var expertInfo = {
        name: uzman,
        education: egitimSekli
      };
      
      if (!flistNormalizedUnvanToExpertsMap.hasOwnProperty(normalizedUnvan)) {
        flistNormalizedUnvanToExpertsMap[normalizedUnvan] = new Set();
      }
      flistNormalizedUnvanToExpertsMap[normalizedUnvan].add(JSON.stringify(expertInfo));
    }
  }

  // EK-2 BİLGİLER sayfasındaki B sütununu kontrol et ve not ekle
  for (var i = 0; i < ek2Data.length; i++) {
    var isYeriAdi = ek2Data[i][0];
    var targetCell = ek2Sheet.getRange(i + 1, 2);

    if (isYeriAdi && isYeriAdi.toString().trim() !== '') {
      var normalizedSearchKey = normalizeText(isYeriAdi.toString().trim());
      var foundExpertInfo = null;
      var foundMatch = false;
      
      var parts = normalizedSearchKey.split(' ');
      var currentSearchPrefix = '';

      for (var k = 0; k < parts.length; k++) {
        currentSearchPrefix += (k > 0 ? ' ' : '') + parts[k];

        var currentMatches = [];
        for (var flistNormalizedUnvan in flistNormalizedUnvanToExpertsMap) {
          if (flistNormalizedUnvan.startsWith(currentSearchPrefix)) {
            currentMatches.push(flistNormalizedUnvan);
          }
        }
        
        if (currentMatches.length > 0) {
          var distinctExpertsForThisPrefix = new Set();
          currentMatches.forEach(matchedNormalizedUnvan => {
              var expertsSet = flistNormalizedUnvanToExpertsMap[matchedNormalizedUnvan];
              if (expertsSet) {
                  expertsSet.forEach(expertInfoStr => distinctExpertsForThisPrefix.add(expertInfoStr));
              }
          });

          if (distinctExpertsForThisPrefix.size === 1) {
            foundExpertInfo = JSON.parse(distinctExpertsForThisPrefix.values().next().value);
            foundMatch = true;
            break;
          } else if (k === parts.length - 1) {
            foundExpertInfo = null;
            foundMatch = false;
            break;
          }
        } else {
          foundExpertInfo = null;
          foundMatch = false;
          break;
        }
      }
      
      // Notu ekle veya kaldır
      if (foundMatch && foundExpertInfo) {
        var newNote = `Uzman: ${foundExpertInfo.name}\nEğitim Şekli: ${foundExpertInfo.education}`;
        targetCell.setNote(newNote);
      } else {
        targetCell.clearNote();
      }
    } else {
      targetCell.clearNote();
    }
  }
}