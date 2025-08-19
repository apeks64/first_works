/**
 * Ana arayüzü gösterir.
 */
function showKatilanListesi() {
  const htmlOutput = HtmlService.createTemplateFromFile('KatilanListesiHTML');
  // Başlangıçta boş bir liste gönderiyoruz, çünkü veriler filtreleme ile yüklenecek
  htmlOutput.initialPersonnelData = JSON.stringify([]);
  const ui = HtmlService.createHtmlOutput(htmlOutput.evaluate())
      .setWidth(900) // Genişliği koruduk
      .setHeight(650); // Yüksekliği koruduk
  SpreadsheetApp.getUi().showModalDialog(ui, 'Katılanlar Listesi');
}

/**
 * Filtre değerine göre personelleri getirir.
 * @param {string} filterValue - Kullanıcının girdiği filtre değeri.
 * @returns {Array<Object>} Filtrelenmiş personellerin listesi.
 */
function getFilteredPersonnel(filterValue) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('EK-2 Bilgiler');
  if (!sheet) {
    throw new Error("EK-2 Bilgiler sayfası bulunamadı. Lütfen sayfa adını kontrol edin.");
  }

  const range = sheet.getDataRange();
  const values = range.getValues();
  const dataRows = values.slice(1); // Başlık satırını atla

  const filteredPersonnel = [];
  const lowerCaseFilter = filterValue ? filterValue.toLowerCase() : '';

  // B sütunundaki verileri benzersiz olarak takip etmek için bir Set kullanalım
  // Bu, aynı iş yerine ait farklı satırlardaki aynı personellerin tekrar listelenmesini önler.
  // Ancak isteğiniz "B sütunundaki veriler benzersiz olarak yüklenecek" ise,
  // ve sonra "aynı karakterleri içeren verileri süzmek için filtre input alanı" varsa,
  // bu biraz çelişkili. B, C, D verilerinin gösterileceği personel listesi isteniyor.
  // Eğer aynı B değeri ve aynı kişi (C+D+E) farklı satırlarda ise ve sadece bir kere gelmesini istiyorsak,
  // benzersizliği daha geniş bir anahtar üzerinden yapmalıyız.
  // Şimdilik B, C, D'den herhangi birinde eşleşen tüm benzersiz *personelleri* getireceğiz.
  // Bir personel (TC bazında) farklı iş yerlerinde (B) olabilir.
  // Sizin "B sütunundaki veriler benzersiz olarak yüklenecek" ifadenizi,
  // "Personel listesinde, B sütunu verisi *de gösterilecek* ve benzersizlik genel olarak sağlanacak" olarak yorumluyorum.
  // Eğer sadece B sütunu değerini benzersiz yapıp, o B'ye ait tüm kişileri almak isteseydik, önceki yöntem gibi iki aşama gerekirdi.
  // Şimdi tek input olduğu için, personel bazında benzersizlik daha mantıklı.

  const seenPersonnel = new Set(); // TC Kimlik No. ile benzersizliği takip edelim.

  dataRows.forEach(row => {
    const bValue = row[1] ? row[1].toString().trim() : '';     // B sütunu (İş Yeri Adı)
    const ad = row[2] ? row[2].toString().trim() : '';        // C sütunu (Adı)
    const soyad = row[3] ? row[3].toString().trim() : '';     // D sütunu (Soyadı)
    const tc = row[4] ? row[4].toString().trim() : '';         // E sütunu (TC Kimlik No.)
    const isValue = row[16] ? row[16].toString().trim() : '';  // Q sütunu (İşi)

    // Arama yapmak istediğimiz değerler
    const searchableString = `${bValue} ${ad} ${soyad}`.toLowerCase();

    // Eğer filtre değeri boşsa veya ilgili sütunlardan birinde eşleşme varsa VE TC bazında bu kişi daha önce eklenmediyse
    if ((!lowerCaseFilter || searchableString.includes(lowerCaseFilter)) && tc && !seenPersonnel.has(tc)) {
      filteredPersonnel.push({
        bValue: bValue,
        ad: ad,
        soyad: soyad,
        tc: tc,
        is: isValue
      });
      seenPersonnel.add(tc);
    }
  });

  return filteredPersonnel.sort((a, b) => { // Sonuçları Ad Soyad'a göre sıralayalım
    const nameA = `${a.ad} ${a.soyad}`;
    const nameB = `${b.ad} ${b.soyad}`;
    return nameA.localeCompare(nameB, 'tr', { sensitivity: 'base' });
  });
}