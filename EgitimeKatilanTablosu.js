/**
 * Eğitim Katılımcı Listesi Oluşturma formunu gösterir.
 */
function showEgitimListesiForm() {
  const htmlOutput = HtmlService.createTemplateFromFile('EgitimListesiHTML');
  
  const ui = htmlOutput.evaluate() 
      .setWidth(950) 
      .setHeight(700); 
  SpreadsheetApp.getUi().showModalDialog(ui, 'Eğitim Katılımcı Listesi Oluştur');
}

/**
 * Sunucu tarafında `EgitimListesiHTML.html` dosyasına erişmek için kullanılır.
 * @param {string} filename - HTML dosyasının adı.
 * @returns {HtmlOutput} HTML çıktısı.
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}


/**
 * FLIST sayfasından benzersiz firma adlarını getirir.
 * @returns {string[]} Benzersiz firma adlarının listesi.
 */
function getFirmsList() {
  Logger.log("DEBUG: getFirmsList çağrıldı.");
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('FLIST');
  if (!sheet) {
    Logger.log("HATA: getFirmsList - FLIST sayfası bulunamadı.");
    throw new Error("FLIST sayfası bulunamadı. Lütfen sayfa adını kontrol edin.");
  }
  const range = sheet.getRange('A:A'); 
  const values = range.getValues();
  
  const firms = new Set();
  for (let i = 1; i < values.length; i++) { // Başlık satırını atla
    if (values[i][0] && String(values[i][0]).trim() !== '') { 
      firms.add(String(values[i][0]).trim());
    }
  }
  const firmArray = Array.from(firms).sort(); 
  Logger.log(`DEBUG: getFirmsList - ${firmArray.length} adet firma bulundu.`);
  return firmArray;
}

/**
 * Seçilen firmaya göre FLIST sayfasından bilgileri getirir.
 * @param {string} firmName - Seçilen firma adı.
 * @returns {Object} Firmanın bilgileri.
 */
function getFirmDetails(firmName) {
  Logger.log(`DEBUG: getFirmDetails çağrıldı. Firma Adı: ${firmName}`);
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('FLIST');
  if (!sheet) {
    Logger.log("HATA: getFirmDetails - FLIST sayfası bulunamadı.");
    throw new Error("FLIST sayfası bulunamadı. Lütfen sayfa adını kontrol edin.");
  }
  const range = sheet.getDataRange();
  const values = range.getValues();
  
  let firmData = {};
  const lowerCaseFirmName = String(firmName || '').trim().toLowerCase(); // Gelen firmayı küçük harfe çevir ve boşlukları temizle

  // Başlık satırını atla (i=1'den başla)
  for (let i = 1; i < values.length; i++) {
    // A sütunundaki firma adını kontrol et
    if (values[i][0] && String(values[i][0]).trim().toLowerCase() === lowerCaseFirmName) { // Karşılaştırmayı da küçük harfe çevir
      firmData.UNVANI = values[i][0] ? String(values[i][0]).trim() : ''; 
      firmData.ISYERIHEKIMI = values[i][18] ? String(values[i][18]).trim() : ''; 
      firmData.IH_BELGE_NO = values[i][20] ? String(values[i][20]).trim() : ''; 
      firmData.ISGUZMANI = values[i][14] ? String(values[i][14]).trim() : ''; 
      firmData.IGU_BELGE_NO = values[i][16] ? String(values[i][16]).trim() : ''; 
      firmData.ADRESI = values[i][7] ? String(values[i][7]).trim() : ''; 
      Logger.log(`DEBUG: getFirmDetails - Firma detayları bulundu: ${JSON.stringify(firmData)}`);
      break; 
    }
  }
  if (Object.keys(firmData).length === 0) {
    Logger.log(`UYARI: getFirmDetails - "${firmName}" için firma detayı bulunamadı.`);
  }
  return firmData;
}


/**
 * EGITIM sayfasından sadece eğitim konularının başlıklarını getirir.
 * @returns {Array<Object>} Başlıklar ve alt konular (sadece başlıklar gönderilecek).
 */
function getEgitimKonulariBasliklari() {
  Logger.log("DEBUG: getEgitimKonulariBasliklari çağrıldı.");
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('EGITIM');
  if (!sheet) {
    Logger.log("HATA: getEgitimKonulariBasliklari - EGITIM sayfası bulunamadı.");
    throw new Error("EGITIM sayfası bulunamadı. Lütfen sayfa adını kontrol edin.");
  }

  const range = sheet.getRange(1, 1, 1, sheet.getLastColumn()); 
  const headers = range.getValues()[0];

  const egitimKonulari = [];
  headers.forEach((header, colIndex) => {
    if (header && String(header).trim() !== '') { 
      egitimKonulari.push({
        baslik: String(header).trim(),
        colIndex: colIndex 
      });
    } else {
        Logger.log(`UYARI: getEgitimKonulariBasliklari - Boş veya geçersiz başlık atlandı. Sütun: ${colIndex}`);
    }
  });
  Logger.log(`DEBUG: getEgitimKonulariBasliklari - ${egitimKonulari.length} adet başlık bulundu.`);
  return egitimKonulari;
}

/**
 * Belirli bir başlığa ait tüm alt konuları EGITIM sayfasından getirir.
 * @param {number} colIndex - Konuların alınacağı sütun indeksi (0 tabanlı).
 * @returns {Array<string>} Alt konuların listesi.
 */
function getEgitimAltKonulariByColumn(colIndex) {
  Logger.log(`DEBUG: getEgitimAltKonulariByColumn çağrıldı. colIndex: ${colIndex}`);
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('EGITIM');
  if (!sheet) {
    Logger.log("HATA: getEgitimAltKonulariByColumn - EGITIM sayfası bulunamadı.");
    throw new Error("EGITIM sayfası bulunamadı. Lütfen sayfa adını kontrol edin.");
  }
  
  if (typeof colIndex !== 'number' || colIndex < 0 || colIndex >= sheet.getLastColumn()) {
      Logger.log(`HATA: getEgitimAltKonulariByColumn - Geçersiz sütun indeksi: ${colIndex}`);
      return []; 
  }

  const startRow = 2; // Başlıklar 1. satırda, konular 2. satırdan başlıyor.
  const numRows = sheet.getLastRow() - startRow + 1; 

  if (numRows <= 0) {
      Logger.log(`UYARI: getEgitimAltKonulariByColumn - Konu bulunamadı, numRows <= 0. Sütun: ${colIndex}`);
      return []; 
  }

  try {
      const range = sheet.getRange(startRow, colIndex + 1, numRows, 1); 
      const values = range.getValues(); 
      Logger.log(`DEBUG: getEgitimAltKonulariByColumn - Sütun ${colIndex} için okunan değerler (ilk 5): ${JSON.stringify(values.slice(0,5))}`); // İlk 5 değeri logla

      const altKonular = [];
      values.forEach((row, rowIndex) => {
        if (row && row[0] !== null && row[0] !== undefined) { 
          const konuMetni = String(row[0]).trim();
          if (konuMetni !== '') {
            altKonular.push(konuMetni);
          } else {
            Logger.log(`UYARI: getEgitimAltKonulariByColumn - Sütun ${colIndex}, Satır ${startRow + rowIndex} boş veya sadece boşluk içeriyor. Atlandı.`);
          }
        } else {
            Logger.log(`UYARI: getEgitimAltKonulariByColumn - Sütun ${colIndex}, Satır ${startRow + rowIndex} null veya undefined değer içeriyor. Atlandı.`);
        }
      });
      Logger.log(`DEBUG: getEgitimAltKonulariByColumn - Sütun ${colIndex} için döndürülen alt konular: ${altKonular.length} adet.`);
      return altKonular;
  } catch (e) {
      Logger.log(`HATA: getEgitimAltKonulariByColumn(${colIndex}) çalışırken istisna: ${e.message}, Stack: ${e.stack}`);
      throw new Error(`Eğitim alt konuları yüklenirken hata: ${e.message}`);
  }
}


/**
 * EK-2 Bilgiler sayfasından filtreye göre personelleri getirir.
 * @param {string} filterValue - Kullanıcının girdiği filtre değeri.
 * @returns {Array<Object>} Filtrelenmiş personellerin listesi.
 */
function getFilteredPersonnelForEgitim(filterValue) {
  Logger.log(`DEBUG: getFilteredPersonnelForEgitim çağrıldı. Sorgu: "${filterValue}"`);
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  // Sayfa adını tam olarak girin, büyük/küçük harf duyarlıdır
  const sheet = ss.getSheetByName('EK-2 BİLGİLER'); 
  
  if (!sheet) { 
    Logger.log("HATA: getFilteredPersonnelForEgitim - Sayfa 'EK-2 BİLGİLER' bulunamadı.");
    throw new Error("Sayfa 'EK-2 BİLGİLER' bulunamadı. Lütfen sayfa adını kontrol edin.");
  }

  const range = sheet.getDataRange();
  const values = range.getValues();
  const dataRows = values.slice(1); // Başlık satırını atla

  const filteredPersonnel = [];
  const lowerCaseFilter = filterValue ? filterValue.toLowerCase() : '';
  const seenPersonnel = new Set(); 

  dataRows.forEach((row, rowIndex) => {
    // İlgili sütunları doğru bir şekilde alıyoruz (0-tabanlı indexler)
    // Sütunlar: B (1), C (2), D (3), E (4), Q (16)
    const firmName = row[1] ? String(row[1]).trim() : ''; 
    const ad = row[2] ? String(row[2]).trim() : ''; 
    const soyad = row[3] ? String(row[3]).trim() : ''; 
    const tc = row[4] ? String(row[4]).trim() : ''; 
    const job = row[16] ? String(row[16]).trim() : ''; 

    // TC kimlik numarasının boş olmaması kontrolü
    if (!tc) {
        Logger.log(`UYARI: Satır ${rowIndex + 2} (Personel): TC Kimlik No boş, atlandı.`);
        return; 
    }

    const searchableString = `${firmName} ${ad} ${soyad} ${tc} ${job}`.toLowerCase(); 

    // Filtreleme koşulu ve mükerrer olmaması
    if ((!lowerCaseFilter || searchableString.includes(lowerCaseFilter)) && !seenPersonnel.has(tc)) {
      filteredPersonnel.push({
        firm: firmName,
        ad: ad,
        soyad: soyad,
        tc: tc,
        is: job
      });
      seenPersonnel.add(tc); 
    }
  });

  // Personelleri Adı Soyadı'na göre Türkçe karakterleri de dikkate alarak sırala
  const sortedPersonnel = filteredPersonnel.sort((a, b) => { 
    const nameA = `${a.ad} ${a.soyad}`;
    const nameB = `${b.ad} ${b.soyad}`;
    return nameA.localeCompare(nameB, 'tr', { sensitivity: 'base' });
  });
  Logger.log(`DEBUG: getFilteredPersonnelForEgitim - ${sortedPersonnel.length} adet personel bulundu.`);
  return sortedPersonnel;
}


/**
 * Boş bir belge oluşturur ve içeriğini dinamik olarak inşa eder.
 * Bu fonksiyon bir kereliğine çalıştırılmalı ve oluşturduğu temiz şablonun ID'si alınmalıdır.
 * Aldığınız ID'yi aşağıdaki createEgitimDocument fonksiyonundaki templateId değişkenine yapıştırın.
 */
function createCleanEgitimTemplate() {
  Logger.log('DEBUG: createCleanEgitimTemplate başlatıldı.');
  const templateName = "Eğitim Katılımcı Belgesi Şablonu";
  const doc = DocumentApp.create(templateName);
  const body = doc.getBody();

  doc.setPageWidth(595.3); 	
  doc.setPageHeight(841.9); 

  doc.setMarginTop(36); 	
  doc.setMarginBottom(36);
  doc.setMarginLeft(36);
  doc.setMarginRight(36);

  const mainTitleParagraph = body.appendParagraph('{{UNVANI}}');
  mainTitleParagraph.setAlignment(DocumentApp.HorizontalAlignment.CENTER).setFontSize(11).setBold(true);

  const subTitleParagraph = body.appendParagraph('VERİLEN EĞİTİME KATILAN LİSTESİ');
  subTitleParagraph.setAlignment(DocumentApp.HorizontalAlignment.CENTER).setFontSize(11).setBold(true);
  body.appendParagraph(''); 

  const infoTableData = [
    ['Eğitimi Veren Kişi', 'Verildiği Tarih'],
    ['{{ISYERIHEKIMI}}', '{{TARIHLER}}'],
    ['Belge No / Tarihi:', 'Eğitim Süresi'],
    ['{{IH_BELGE_NO}}', '{{TOPLAM_SAAT}}'],
    ['{{ISGUZMANI}}', 'Eğitimin Verildiği Yer'],
    ['Belge No / Tarihi:', '{{ADRESI}}'],
    ['{{IGU_BELGE_NO}}', ''] 
  ];
  const infoTable = body.appendTable(infoTableData);

  infoTable.setColumnWidth(0, 250); 
  infoTable.setColumnWidth(1, 250); 
  
  if (infoTable.getNumRows() > 0 && infoTable.getRow(0).getNumCells() > 1) { 
    const cell0 = infoTable.getRow(0).getCell(0);
    const cell1 = infoTable.getRow(0).getCell(1);
    if (cell0.getNumChildren() > 0 && cell0.getChild(0).getType() === DocumentApp.ElementType.PARAGRAPH) {
      cell0.getChild(0).asParagraph().setBold(true).setFontSize(9);
    }
    if (cell1.getNumChildren() > 0 && cell1.getChild(0).getType() === DocumentApp.ElementType.PARAGRAPH) {
      cell1.getChild(0).asParagraph().setBold(true).setFontSize(9);
    }
  }

  for(let r = 0; r < infoTable.getNumRows(); r++) {
      for(let c = 0; c < infoTable.getRow(r).getNumCells(); c++) { 
          const cell = infoTable.getCell(r, c);
          if (cell.getNumChildren() > 0 && cell.getChild(0).getType() === DocumentApp.ElementType.PARAGRAPH) {
              cell.getChild(0).asParagraph().setFontFamily('Arial').setFontSize(9);
          }
          cell.setVerticalAlignment(DocumentApp.VerticalAlignment.MIDDLE);
      }
  }

  body.appendParagraph(''); // Boş satır
  const egitimKonulariTitle = body.appendParagraph('Eğitim Konuları');
  egitimKonulariTitle.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  egitimKonulariTitle.setBold(true).setFontSize(11);
  body.appendParagraph(''); // Boş satır

  // Eğitim konuları yer tutucusu
  const egitimKonulariPlaceholder = body.appendParagraph('{{EGITIM KONULARI}}'); // Sadece metin olmalı
  egitimKonulariPlaceholder.setAlignment(DocumentApp.HorizontalAlignment.LEFT);
  body.appendParagraph(''); // Boş satır

  const katilanlarListesiTitle = body.appendParagraph('Katılanlar Listesi');
  katilanlarListesiTitle.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  katilanlarListesiTitle.setBold(true).setFontSize(11);
  body.appendParagraph(''); // Boş satır

  // Katılımcılar listesi yer tutucusu
  const katilanlarListesiPlaceholder = body.appendParagraph('{{KATILANLAR LİSTESİ}}'); // Sadece metin olmalı
  katilanlarListesiPlaceholder.setAlignment(DocumentApp.HorizontalAlignment.LEFT);
  body.appendParagraph(''); // Boş satır (Belgenin sonu için güvenlik boşluğu)

  doc.saveAndClose();

  const file = DriveApp.getFileById(doc.getId());
  Logger.log(`DEBUG: Yeni şablon oluşturuldu: ${file.getName()} - ID: ${file.getId()} - URL: ${file.getUrl()}`);
  return file.getUrl();
}


function createEgitimDocument(data) {
  const {
    selectedFirm,
    UNVANI,
    ISYERIHEKIMI,
    IH_BELGE_NO,
    ISGUZMANI,
    IGU_BELGE_NO,
    ADRESI,
    formattedDates,
    formattedHours,
    selectedEgitimKonulari,
    selectedPersonnel
  } = data;

  const EGITIM_TARIHI = formattedDates + (formattedHours ? " | " + formattedHours : "");

  // Eğitim konularını formatla
  let egitimKonulariFormatted = "";
  selectedEgitimKonulari.forEach(konuObj => {
    egitimKonulariFormatted += konuObj.baslik + "\n";
    if (Array.isArray(konuObj.konular)) {
      konuObj.konular.forEach(alt => {
        egitimKonulariFormatted += "    - " + alt + "\n";
      });
    }
    egitimKonulariFormatted += "\n";
  });

  // Katılımcılar listesini formatla
  let katilanlarFormatted = "TC Kimlik No | Adı Soyadı | Mesleği\n";
  katilanlarFormatted += "--------------------------------------\n";
  selectedPersonnel.forEach(p => {
    katilanlarFormatted += `${p.tc} | ${p.ad} ${p.soyad} | ${p.is || ''}\n`;
  });

  // Şablon ID'si (az önce oluşturduğumuz şablonun ID’sini buraya gir!)
  const templateId = '1Cq8VG-ph4M6mJa-Mx_9y4EnKQM3MoQAuJr9F4Y00Nuw';  // örn: '1jnytmZfHhHi9bkgBCX5uMRLNFTffmGKqTurMfwKzShE'
  const folderId = '1yrwxxbb7WA2MxQ0UoZ5LN9jJahOaD5Uq'; // Belgenin kaydedileceği klasör ID'si

  const templateFile = DriveApp.getFileById(templateId);
  const newFile = templateFile.makeCopy(`${UNVANI} - Eğitim Katılımcı Listesi`, DriveApp.getFolderById(folderId));
  const doc = DocumentApp.openById(newFile.getId());
  const body = doc.getBody();

  // Yer tutucuları değiştir
  body.replaceText('{{UNVANI}}', UNVANI || '');
  body.replaceText('{{ISYERIHEKIMI}}', ISYERIHEKIMI || '');
  body.replaceText('{{IH_BELGE_NO}}', IH_BELGE_NO || '');
  body.replaceText('{{ISGUZMANI}}', ISGUZMANI || '');
  body.replaceText('{{IGU_BELGE_NO}}', IGU_BELGE_NO || '');
  body.replaceText('{{ADRESI}}', ADRESI || '');
  body.replaceText('{{EGITIM_TARIHI}}', EGITIM_TARIHI || '');
  body.replaceText('{{EGITIM KONULARI}}', egitimKonulariFormatted || '');
  body.replaceText('{{KATILANLAR LİSTESİ}}', katilanlarFormatted || '');

  doc.saveAndClose();

  return newFile.getUrl();
}
