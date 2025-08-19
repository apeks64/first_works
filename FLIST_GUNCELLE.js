/**
 * Herhangi bir Google Sheets sayfasının ilk satırındaki sütun başlıklarını okur
 * ve her başlığın 0-tabanlı indeksini içeren bir Map döndürür.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - İşlem yapılacak Google Sheets sayfası.
 * @returns {Map<string, number>} Sütun başlıklarını indeksleriyle eşleştiren bir Map.
 */
function getColumnHeaderMap(sheet) {
  const headerRow = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const headerMap = new Map();
  headerRow.forEach((header, index) => {
    headerMap.set(header.trim(), index); // Boşlukları temizle ve Map'e ekle
  });
  return headerMap;
}

/**
 * Ana işlem fonksiyonu. Tüm alt fonksiyonları sırayla çalıştırır.
 */
function tumIslemleriCalistir() {
  veriAktarimi();
  hesaplaVeYaz();
  logUnmatchedDistricts();
}

/**
 * ATAMALAR sayfasındaki verileri FLIST sayfasına aktarır ve günceller.
 * NACE kod tanımlarını ve İSG-Katip atamalarını işler.
 */
function veriAktarimi() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const atamalar = ss.getSheetByName("ATAMALAR");
  const flist = ss.getSheetByName("FLIST");
  const naceSheet = ss.getSheetByName("NACE");

  if (!flist || !atamalar || !naceSheet) {
    Browser.msgBox("Hata: Gerekli sayfalardan biri bulunamadı. Lütfen 'FLIST', 'ATAMALAR' ve 'NACE' sayfalarının doğru adlandırıldığından emin olun.");
    return;
  }

  // --- Sütun Başlık Map'lerini Oluştur ---
  const flistHeaders = getColumnHeaderMap(flist);
  const atamalarHeaders = getColumnHeaderMap(atamalar);
  const naceHeaders = getColumnHeaderMap(naceSheet);

  // --- Gerekli Sütunların Varlığını Kontrol Et ---
  // FLIST sayfasından beklenen başlıklar
  const requiredFlistHeaders = [
    "ID", "UNVANI", "SGKNO", "TEHLİKE_SINIFI", "İL", "NACE_KODU",
    "ISGUZMANI", "IGU_GOREVI", "IGU_BELGE_NO",
    "ISYERIHEKIMI", "IH_GOREVI", "IH_BELGE_NO",
    "ATAMA TARIHI", "ÇALIŞAN SAYISI", "NACE_TANIM"
  ];
  // ATAMALAR sayfasından beklenen başlıklar
  const requiredAtamalarHeaders = [
    "Görevlendirilen Kişi Sertifika Tipi", "Görevlendirilen Kişi Ad Soyad", "Görevlendirilen Kişi Sertifika No",
    "Hizmet Alan İşyeri Unvanı", "Hizmet Alan İşyeri SGK/DETSİS No", "Hizmet Alan İşyeri Tehlike Sınıfı",
    "Hizmet Alan İşyeri Nace Kodu", "Hizmet Alan İşyeri İli", "Hizmet Alan İşyeri Çalışan Sayısı",
    "Hizmet Alan İşyeri ID", "Sözleşme Başlangıç Tarihi"
  ];
  // NACE sayfasından beklenen başlıklar
  const requiredNaceHeaders = ["NACE Rev.2.1 Ver.3 Altılı Kod", "NACE Rev.2.1 Ver.3 Altılı Tanım"];

  // Başlık kontrol fonksiyonu
  function checkHeaders(sheetName, headersMap, requiredHeaders) {
    for (const header of requiredHeaders) {
      if (!headersMap.has(header)) {
        Browser.msgBox(`Hata: ${sheetName} sayfasında '${header}' sütunu bulunamadı. Lütfen sütun başlıklarını kontrol edin.`);
        return false;
      }
    }
    return true;
  }

  if (!checkHeaders("FLIST", flistHeaders, requiredFlistHeaders) ||
      !checkHeaders("ATAMALAR", atamalarHeaders, requiredAtamalarHeaders) ||
      !checkHeaders("NACE", naceHeaders, requiredNaceHeaders)) {
    return; // Eksik başlık varsa dur
  }

  // --- NACE Kodu Tanım Map'ini Oluştur ---
  const naceData = naceSheet.getDataRange().getValues();
  const naceMap = new Map();
  const naceKoduCol = naceHeaders.get("NACE Rev.2.1 Ver.3 Altılı Kod");
  const naceTanimiCol = naceHeaders.get("NACE Rev.2.1 Ver.3 Altılı Tanım");

  for (let i = 1; i < naceData.length; i++) {
    const naceKodu = naceData[i][naceKoduCol];
    const naceTanimi = naceData[i][naceTanimiCol];
    if (naceKodu) {
      naceMap.set(String(naceKodu).trim(), naceTanimi);
    }
  }

  // --- ATAMALAR Sayfası Verilerini Oku ve İşle ---
  const atamalarVeri = atamalar.getDataRange().getValues();
  const atamalarIDSet = new Set();
  const idMap = new Map();

  // ATAMALAR sütun indeksleri (Dinamik)
  const atamalarSertifikaTipiCol = atamalarHeaders.get("Görevlendirilen Kişi Sertifika Tipi");
  const atamalarAdSoyadCol = atamalarHeaders.get("Görevlendirilen Kişi Ad Soyad");
  const atamalarSertifikaNoCol = atamalarHeaders.get("Görevlendirilen Kişi Sertifika No");
  const atamalarIsverenUnvanCol = atamalarHeaders.get("Hizmet Alan İşyeri Unvanı");
  const atamalarSgkNoCol = atamalarHeaders.get("Hizmet Alan İşyeri SGK/DETSİS No");
  const atamalarTehlikeSinifiCol = atamalarHeaders.get("Hizmet Alan İşyeri Tehlike Sınıfı");
  const atamalarNaceKoduCol = atamalarHeaders.get("Hizmet Alan İşyeri Nace Kodu");
  const atamalarIlCol = atamalarHeaders.get("Hizmet Alan İşyeri İli");
  const atamalarCalisanSayisiCol = atamalarHeaders.get("Hizmet Alan İşyeri Çalışan Sayısı");
  const atamalarIsyeriIDCol = atamalarHeaders.get("Hizmet Alan İşyeri ID");
  const atamalarBaslangicTarihiCol = atamalarHeaders.get("Sözleşme Başlangıç Tarihi");


  for (let i = 1; i < atamalarVeri.length; i++) {
    const row = atamalarVeri[i];
    const id = row[atamalarIsyeriIDCol];
    if (!id) continue;

    atamalarIDSet.add(id);

    let data = idMap.get(id) || {
      unvan: row[atamalarIsverenUnvanCol],
      sgkNo: row[atamalarSgkNoCol],
      tehlikeSinifi: row[atamalarTehlikeSinifiCol],
      naceKodu: row[atamalarNaceKoduCol],
      il: row[atamalarIlCol],
      calisanSayisi: row[atamalarCalisanSayisiCol],
      atamaTarihi: "",
      iguAdSoyad: "",
      iguSertifikaTipi: "", // Bu değer aşağıda dönüştürülecek
      iguSertifikaNo: "",
      ihAdSoyad: "",
      ihSertifikaTipi: "", // Bu değer aşağıda dönüştürülecek
      ihSertifikaNo: ""
    };

    const sertifikaTipiHam = String(row[atamalarSertifikaTipiCol]).trim();

    // İŞ GÜVENLİĞİ UZMANLIĞI SERTİFİKALARI İÇİN DÜZENLEME BAŞLANGICI
    if (sertifikaTipiHam.includes("İş Güvenliği Uzmanlığı Sertifikası")) {
      data.iguAdSoyad = row[atamalarAdSoyadCol];
      
      // Sertifika tipini kontrol et ve yeni formata dönüştür
      if (sertifikaTipiHam.includes("A Sınıfı")) {
        data.iguSertifikaTipi = "A Sınıfı İş Güvenliği Uzmanı";
      } else if (sertifikaTipiHam.includes("B Sınıfı")) {
        data.iguSertifikaTipi = "B Sınıfı İş Güvenliği Uzmanı";
      } else if (sertifikaTipiHam.includes("C Sınıfı")) {
        data.iguSertifikaTipi = "C Sınıfı İş Güvenliği Uzmanı";
      } else {
        data.iguSertifikaTipi = "İş Güvenliği Uzmanı"; // Diğer İGU tipleri için varsayılan
      }
      
      data.iguSertifikaNo = row[atamalarSertifikaNoCol];
      
      if (row[atamalarBaslangicTarihiCol]) {
        data.atamaTarihi = row[atamalarBaslangicTarihiCol];
      }
    } 
    // İŞYERİ HEKİMLİĞİ SERTİFİKASI İÇİN DÜZENLEME BAŞLANGICI
    else if (sertifikaTipiHam.includes("İşyeri Hekimliği Sertifikası")) {
      data.ihAdSoyad = row[atamalarAdSoyadCol];
      data.ihSertifikaTipi = "İşyeri Hekimi"; // Bu kısım sabit kalacak
      data.ihSertifikaNo = row[atamalarSertifikaNoCol];
    }
    // ... (diğer olası sertifika tipleri veya else durumu burada eklenebilir) ...

    idMap.set(id, data);
  }

  // --- FLIST Sayfası Verilerini Oku ve İşle ---
  const flistVeri = flist.getDataRange().getValues();
  const flistIDListesi = new Map();

  // FLIST sütun indeksleri (Dinamik)
  const flistIDCol = flistHeaders.get("ID");
  const flistUnvanCol = flistHeaders.get("UNVANI");
  const flistSgkNoCol = flistHeaders.get("SGKNO");
  const flistIlCol = flistHeaders.get("İL");
  const flistTehlikeSinifiCol = flistHeaders.get("TEHLİKE_SINIFI");
  const flistNaceKoduCol = flistHeaders.get("NACE_KODU");
  const flistIguAdSoyadCol = flistHeaders.get("ISGUZMANI");
  const flistIguSertifikaTipiCol = flistHeaders.get("IGU_GOREVI");
  const flistIguSertifikaNoCol = flistHeaders.get("IGU_BELGE_NO");
  const flistIhAdSoyadCol = flistHeaders.get("ISYERIHEKIMI");
  const flistIhSertifikaTipiCol = flistHeaders.get("IH_GOREVI");
  const flistIhSertifikaNoCol = flistHeaders.get("IH_BELGE_NO");
  const flistAtamaTarihiCol = flistHeaders.get("ATAMA TARIHI");
  const flistCalisanSayisiCol = flistHeaders.get("ÇALIŞAN SAYISI");
  const flistNaceTanimiCol = flistHeaders.get("NACE_TANIM");


  for (let i = 1; i < flistVeri.length; i++) {
    const row = flistVeri[i];
    const id = row[flistIDCol];
    if (id) {
      flistIDListesi.set(id, i + 1); // ID ve Satır Numarası (1-tabanlı)
    }
  }

  // --- Verileri Aktar ve Güncelle ---
  let maxColIndex = 0;
  flistHeaders.forEach((index) => {
    if (index > maxColIndex) maxColIndex = index;
  });
  const totalFlistColumns = maxColIndex + 1; // 0-tabanlı olduğu için +1

  // Tüm FLIST sayfasını beyaz renge çevir (başlık satırı hariç)
  if (flist.getLastRow() > 1) { // Sadece başlık değil, başka satırlar da varsa temizle
    flist.getRange(2, 1, flist.getLastRow() - 1, flist.getLastColumn()).setBackground(null);
  }

  idMap.forEach((data, id) => {
    if (flistIDListesi.has(id)) {
      // FLIST'te bulunan ID'ler için güncelleme
      const rowIndex = flistIDListesi.get(id);
      const mevcutVeri = flist.getRange(rowIndex, 1, 1, totalFlistColumns).getValues()[0];
      const yeniVeri = [...mevcutVeri];
      let degisenHücreler = [];

      const updates = [
        { flistCol: flistUnvanCol, atamalarData: data.unvan },
        { flistCol: flistSgkNoCol, atamalarData: data.sgkNo },
        { flistCol: flistIlCol, atamalarData: data.il },
        { flistCol: flistTehlikeSinifiCol, atamalarData: data.tehlikeSinifi },
        { flistCol: flistNaceKoduCol, atamalarData: data.naceKodu },
        { flistCol: flistIguAdSoyadCol, atamalarData: data.iguAdSoyad },
        { flistCol: flistIguSertifikaTipiCol, atamalarData: data.iguSertifikaTipi }, // Düzeltilen değer burada kullanılacak
        { flistCol: flistIguSertifikaNoCol, atamalarData: data.iguSertifikaNo },
        { flistCol: flistIhAdSoyadCol, atamalarData: data.ihAdSoyad },
        { flistCol: flistIhSertifikaTipiCol, atamalarData: data.ihSertifikaTipi }, // Düzeltilen değer burada kullanılacak
        { flistCol: flistIhSertifikaNoCol, atamalarData: data.ihSertifikaNo },
        { flistCol: flistAtamaTarihiCol, atamalarData: data.atamaTarihi },
        { flistCol: flistCalisanSayisiCol, atamalarData: data.calisanSayisi }
      ];

      updates.forEach(item => {
        const flistColIndex = item.flistCol;
        const yeniDeger = item.atamalarData;
        // String() ile dönüştürerek null/undefined ve farklı tiplerdeki karşılaştırma hatalarını önle
        if (flistColIndex !== undefined && String(mevcutVeri[flistColIndex] || "").trim() !== String(yeniDeger || "").trim()) {
          yeniVeri[flistColIndex] = yeniDeger;
          degisenHücreler.push(flistColIndex + 1); // 1-tabanlı sütun indeksi
        }
      });

      // NACE tanımını FLIST'e yaz (eğer mevcut ve farklıysa)
      if (data.naceKodu && naceMap.has(data.naceKodu) && flistNaceTanimiCol !== undefined) {
        const naceTanimi = naceMap.get(data.naceKodu);
        if (String(mevcutVeri[flistNaceTanimiCol] || "").trim() !== String(naceTanimi || "").trim()) {
          yeniVeri[flistNaceTanimiCol] = naceTanimi;
          degisenHücreler.push(flistNaceTanimiCol + 1);
        }
      }

      if (degisenHücreler.length > 0) {
        flist.getRange(rowIndex, 1, 1, totalFlistColumns).setValues([yeniVeri]);
        degisenHücreler.forEach(col => {
          flist.getRange(rowIndex, col).setBackground("yellow");
        });
      }

    } else {
      // FLIST'te bulunmayan ID'ler için yeni satır oluştur
      const yeniSatir = Array(totalFlistColumns).fill("");

      // Dinamik olarak sütunlara değer ata
      if (flistUnvanCol !== undefined) yeniSatir[flistUnvanCol] = data.unvan;
      if (flistSgkNoCol !== undefined) yeniSatir[flistSgkNoCol] = data.sgkNo;
      if (flistIDCol !== undefined) yeniSatir[flistIDCol] = id;
      if (flistTehlikeSinifiCol !== undefined) yeniSatir[flistTehlikeSinifiCol] = data.tehlikeSinifi;
      if (flistIlCol !== undefined) yeniSatir[flistIlCol] = data.il;
      if (flistNaceKoduCol !== undefined) yeniSatir[flistNaceKoduCol] = data.naceKodu;
      if (flistIguAdSoyadCol !== undefined) yeniSatir[flistIguAdSoyadCol] = data.iguAdSoyad;
      if (flistIguSertifikaTipiCol !== undefined) yeniSatir[flistIguSertifikaTipiCol] = data.iguSertifikaTipi; // Düzeltilen değer burada kullanılacak
      if (flistIguSertifikaNoCol !== undefined) yeniSatir[flistIguSertifikaNoCol] = data.iguSertifikaNo;
      if (flistIhAdSoyadCol !== undefined) yeniSatir[flistIhAdSoyadCol] = data.ihAdSoyad;
      if (flistIhSertifikaTipiCol !== undefined) yeniSatir[flistIhSertifikaTipiCol] = data.ihSertifikaTipi; // Düzeltilen değer burada kullanılacak
      if (flistIhSertifikaNoCol !== undefined) yeniSatir[flistIhSertifikaNoCol] = data.ihSertifikaNo;
      if (flistAtamaTarihiCol !== undefined) yeniSatir[flistAtamaTarihiCol] = data.atamaTarihi;
      if (flistCalisanSayisiCol !== undefined) yeniSatir[flistCalisanSayisiCol] = data.calisanSayisi;

      // NACE tanımını FLIST'e yaz
      if (data.naceKodu && naceMap.has(data.naceKodu) && flistNaceTanimiCol !== undefined) {
        yeniSatir[flistNaceTanimiCol] = naceMap.get(data.naceKodu);
      }

      flist.appendRow(yeniSatir);
      flist.getRange(flist.getLastRow(), 1, 1, totalFlistColumns).setBackground("yellow");
    }
  });

  // --- ATAMALAR'da olmayan FLIST ID'lerini TURUNCU yap ---
  const flistData = flist.getDataRange().getValues(); // Güncel veriyi tekrar oku
  const flistLastCol = flist.getLastColumn(); // Dinamik son sütun

  // Manisa için il sütununu dinamik olarak al
  const flistIlIndex = flistHeaders.get("İL");
  const flistIDIndex = flistHeaders.get("ID");


  // Başlık satırını atlamak için i=1'den başla
  for (let i = 1; i < flistData.length; i++) {
    const row = flistData[i];
    const flistID = row[flistIDIndex];
    const il = row[flistIlIndex];

    // Sayfa boşsa veya sadece başlık varsa, getValues() boş bir dizi döndürebilir
    if (flistID === undefined) continue;

    if (flistID && !atamalarIDSet.has(flistID) && String(il || "").trim().toUpperCase() !== "MANİSA") {
      flist.getRange(i + 1, 1, 1, flistLastCol).setBackground("#f4c7c3"); // Turuncu tonu
    }
  }

  Browser.msgBox("Veri aktarımı ve güncellemeler tamamlandı.");
}

/**
 * SGK No'dan ilçe bilgisini çıkarıp FLIST'e yazar.
 */
function logUnmatchedDistricts() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("FLIST"); // FLIST sayfası
  if (!sheet) {
    Browser.msgBox("Hata: 'FLIST' sayfası bulunamadı.");
    return;
  }
  const flistHeaders = getColumnHeaderMap(sheet);

  const sgkNoCol = flistHeaders.get("SGKNO");
  const ilceCol = flistHeaders.get("ILCE"); // FLIST'te "ILCE" diye bir sütun olmalı

  if (sgkNoCol === undefined || ilceCol === undefined) {
    Browser.msgBox("Hata: FLIST sayfasında 'SGKNO' veya 'ILCE' sütun başlığı bulunamadı.");
    return;
  }

  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) { // Sadece başlık satırı varsa veya sayfa boşsa
      Logger.log("FLIST sayfasında ilçe hesaplanacak veri yok.");
      return;
  }
  
  var dataRange = sheet.getRange(2, sgkNoCol + 1, lastRow - 1, 1);
  var data = dataRange.getValues();
  var output = [];

  var districtMap = {
    "03502": "BAYINDIR", "03503": "BERGAMA", "03504": "BORNOVA", "03505": "ÇEŞME",
    "03506": "DİKİLİ", "03507": "FOÇA", "03509": "KARŞIYAKA", "03510": "KEMALPAŞA",
    "03512": "KİRAZ", "03513": "MENEMEN", "03514": "ÖDEMİŞ", "03515": "SEFERİHİSAR",
    "03516": "SELÇUK", "03518": "TORBALI", "03519": "NARLIDERE", "03520": "ALİAĞA",
    "03521": "BUCA", "03522": "MENDERES", "03523": "KONAK", "03524": "BEYDAĞ",
    "03525": "BALÇOVA", "03526": "ÇİĞLİ", "03527": "GAZİEMİR", "03528": "GÜZELBAHÇE",
    "03529": "URLA", "03530": "BAYRAKLI", "03531": "KARABAĞLAR"
  };

  var logMessages = [];

  for (var i = 0; i < data.length; i++) {
    var sgkNo = data[i][0];

    if (!sgkNo || String(sgkNo).trim() === "") {
      output.push([""]);
      continue;
    }

    if (typeof sgkNo === "number") {
      sgkNo = sgkNo.toFixed(0);
    }

    sgkNo = String(sgkNo).padStart(26, "0"); // SGKNO'nun başında '0' eksikse ekle
    var districtCode = sgkNo.substring(16, 21); // SGKNO'nun 17-21. karakterleri (0-tabanlı 16-20)
    var districtName = districtMap[districtCode];

    if (districtName) {
      output.push([districtName]);
    } else {
      output.push(["BİLİNMİYOR"]);
      logMessages.push(`Satır ${i + 2}: '${districtCode}' kodu districtMap'te yok.`);
    }
  }

  // İlçe isimlerini 'ILCE' sütununa yaz
  if (output.length > 0) {
    sheet.getRange(2, ilceCol + 1, output.length, 1).setValues(output);
  }

  // Eğer eşleşmeyen varsa LOG sayfasına topluca yaz
  if (logMessages.length > 0) {
    var logSheet = ss.getSheetByName("LOG");
    if (!logSheet) {
      logSheet = ss.insertSheet("LOG");
    }

    var lastLogRow = logSheet.getLastRow();
    var now = new Date();
    var rows = logMessages.map(function(msg) {
      return [now.toLocaleString(), msg];
    });

    logSheet.getRange(lastLogRow + 1, 1, rows.length, 2).setValues(rows);
  }
}

/**
 * Atama Tarihi ve Tehlike Sınıfına göre hesaplama yapıp FLIST'e yazar.
 * ("ADP-RD-TARIH" sütununa)
 */
function hesaplaVeYaz() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('FLIST');
  if (!sheet) {
    Browser.msgBox("Hata: 'FLIST' sayfası bulunamadı.");
    return;
  }

  const flistHeaders = getColumnHeaderMap(sheet);

  const atamaTarihiCol = flistHeaders.get("ATAMA TARIHI"); // FLIST'teki "ATAMA TARIHI"
  const tehlikeSinifiCol = flistHeaders.get("TEHLİKE_SINIFI"); // FLIST'teki "TEHLİKE_SINIFI"
  const yillikSureCol = flistHeaders.get("ADP-RD-TARIH"); // FLIST'teki "ADP-RD-TARIH"

  if (atamaTarihiCol === undefined || tehlikeSinifiCol === undefined || yillikSureCol === undefined) {
    Browser.msgBox("Hata: FLIST sayfasında gerekli sütun başlıkları (ATAMA TARIHI, TEHLİKE_SINIFI, ADP-RD-TARIH) bulunamadı.");
    return;
  }

  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) { // Sadece başlık satırı varsa veya sayfa boşsa
      Logger.log("FLIST sayfasında hesaplanacak veri yok.");
      return;
  }
  const data = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();
  const valuesToUpdate = [];

  for (let i = 0; i < data.length; i++) {
    const tarih = data[i][atamaTarihiCol];
    const sinif = String(data[i][tehlikeSinifiCol] || "").trim(); // Boş veya null ise boş string'e çevir

    let yeniTarihDegeri = "";

    // Tarih değerinin geçerli olup olmadığını kontrol et
    if (tarih && !isNaN(new Date(tarih))) {
      const isgKatipTarih = new Date(tarih);
      let eklenecekYil;

      switch (sinif) {
        case 'Az Tehlikeli':
          eklenecekYil = 6;
          break;
        case 'Tehlikeli':
          eklenecekYil = 4;
          break;
        case 'Çok Tehlikeli':
          eklenecekYil = 2;
          break;
        default:
          eklenecekYil = 0; // Tanımlı sınıf yoksa veya geçersizse
      }

      if (eklenecekYil > 0) {
        let yeniTarih = new Date(isgKatipTarih.getFullYear() + eklenecekYil, isgKatipTarih.getMonth(), isgKatipTarih.getDate());
        yeniTarihDegeri = Utilities.formatDate(yeniTarih, ss.getSpreadsheetTimeZone(), "dd/MM/yyyy");
      }
    }
    valuesToUpdate.push([yeniTarihDegeri]);
  }

  // Hesaplanan yeni tarihleri 'RD-ADP SON TARIH' sütununa yaz
  if (valuesToUpdate.length > 0) {
    sheet.getRange(2, yillikSureCol + 1, valuesToUpdate.length, 1).setValues(valuesToUpdate);
  }
}function updateNotesAdvancedMatching() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ek2Sheet = ss.getSheetByName('EK-2 BİLGİLER'); // EK-2 BİLGİLER sayfanızın adını kontrol edin
  var flistSheet = ss.getSheetByName('FLIST');     // FLIST sayfanızın adını kontrol edin

  if (!ek2Sheet || !flistSheet) {
    Logger.log('Sayfalardan biri bulunamadı. Lütfen sayfa adlarını kontrol edin.');
    SpreadsheetApp.getUi().alert('Hata', 'Sayfalardan biri bulunamadı. Lütfen "EK-2 BİLGİLER" ve "FLIST" sayfa adlarını kontrol edin.', SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }

  var ek2Data = ek2Sheet.getRange('B:B').getValues(); // EK-2 BİLGİLER B sütunundaki tüm verileri al
  var flistData = flistSheet.getRange('A:E').getValues(); // FLIST sayfasının A'dan E'ye kadar olan verilerini al

  // FLIST sayfasındaki UNVANI (A) ve UZMAN (E) verilerini bir harita olarak hazırla
  var flistEntries = []; // Her bir FLIST A sütunu değerini tutmak için
  var flistMap = {};     // Key: FLIST UNVAN (trimmed), Value: UZMAN
  for (var i = 0; i < flistData.length; i++) {
    var unvan = flistData[i][0];
    var uzman = flistData[i][4];
    if (unvan && unvan.toString().trim() !== '') {
      var trimmedUnvan = unvan.toString().trim();
      flistEntries.push(trimmedUnvan);
      flistMap[trimmedUnvan] = uzman;
    }
  }

  // EK-2 BİLGİLER sayfasındaki B sütununu kontrol et ve not ekle
  for (var i = 0; i < ek2Data.length; i++) {
    var isYeriAdi = ek2Data[i][0]; // B sütunundaki İş Yerinin Adı
    var targetCell = ek2Sheet.getRange(i + 1, 2); // B sütunundaki ilgili hücre

    // Geçerli bir değer olduğundan emin olalım
    if (isYeriAdi && isYeriAdi.toString().trim() !== '') {
      var searchKey = isYeriAdi.toString().trim();
      var foundExpert = null;
      
      // EK-2'deki işyeri adını kelimelerine ayırıp, her adımda eşleşme arayalım
      var parts = searchKey.split(' ');
      var currentSearchPrefix = ''; // Aranacak öneki tutar

      // İlk kelimeden başlayarak öneki uzatarak eşleşme arıyoruz
      for (var k = 0; k < parts.length; k++) {
        currentSearchPrefix += (k > 0 ? ' ' : '') + parts[k]; // Öneki bir sonraki kelimeyle uzat

        var currentMatches = []; // Bu önekle eşleşen FLIST ünvanları
        for (var j = 0; j < flistEntries.length; j++) {
          var flistUnvan = flistEntries[j];
          if (flistUnvan.startsWith(currentSearchPrefix)) {
            currentMatches.push(flistUnvan);
          }
        }

        // Eğer bu önekle tam olarak BİR TANE eşleşme bulunduysa, bu bizim aradığımız eşleşmedir.
        if (currentMatches.length === 1) {
          foundExpert = flistMap[currentMatches[0]];
          break; // Tekil bir eşleşme bulduk, daha fazla aramaya gerek yok
        } else if (currentMatches.length === 0) {
          // Eğer bu önekle hiç eşleşme bulunamadıysa, daha uzun öneklerle denemek anlamsız.
          // Önceki (daha kısa) öneklerde de tekil eşleşme bulamadıysak, uzmanı boş bırakalım.
          if (k === 0) { // İlk kelimeyle bile eşleşme yoksa
            foundExpert = null;
          }
          break; // Bu dalda eşleşme yok, döngüden çık
        }
        // Eğer currentMatches.length > 1 ise, yani birden fazla eşleşme varsa,
        // bir sonraki kelimeyi ekleyerek öneki uzatmaya devam edeceğiz
        // daha spesifik bir eşleşme bulmak için.
      }
      
      // Notu ekle veya kaldır
      if (foundExpert) {
        targetCell.setNote(foundExpert);
      } else {
        targetCell.clearNote(); // Eşleşme bulunamazsa veya birden fazla eşleşme varsa notu kaldır
      }

    } else {
      targetCell.clearNote(); // İş Yerinin Adı boşsa notu kaldır
    }
  }
}