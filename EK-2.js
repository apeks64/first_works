const TEMPLATE_FILE_ID = '1jnytmZfHhHi9bkgBCX5uMRLNFTffmGKqTurMfwKzShE';
const DESTINATION_FOLDER_ID = '1TVjW_R8SSZsQ3qNcyA6X6ghlJNyN0niJ';

// ================================
// FORM SUBMIT TRİGGER FONKSİYONU
// ================================
/**
 * Form yanıtından sonra otomatik olarak tetiklenir (tetikleyici ayarlanmalı).
 * @param {Object} e - Form yanıtı olayı nesnesi.
 */
function onFormSubmit(e) {
  Logger.log("=== onFormSubmit başladı ===");
  Logger.log("Event objesi: " + JSON.stringify(e));
  
  try {
    if (!e || !e.range) {
      Logger.log("HATA: Event objesi veya range bulunamadı!");
      return;
    }
    
    const row = e.range.getRow();
    Logger.log("İşlem yapılacak satır: " + row);
    
    // 1. Sadece yeni eklenen satırı büyük harfe çevir (optimize edilmiş)
    Logger.log("1. convertRowToUpperCase başlıyor...");
    convertRowToUpperCase(row);
    
    // 2. Yeni yanıtları arşivle
    Logger.log("2. archiveNewFormResponses başlıyor...");
    archiveNewFormResponses();
    
    // 3. Belge oluştur
    Logger.log("3. generateDocument başlıyor...");
    generateDocument(row, true);
    
    Logger.log("=== Form yanıtı işlemi başarıyla tamamlandı ===");
  } catch (error) {
    Logger.log("=== HATA OLUŞTU ===");
    Logger.log("Hata mesajı: " + error.message);
    Logger.log("Hata stack: " + error.stack);
  }
}

// ================================
// TÜRKÇE KARAKTER DÖNÜŞÜMLERİ
// ================================
/**
 * Türkçe karakterleri doğru büyük harfe dönüştürür
 */
function convertToTurkishUpperCase(str) {
  if (typeof str !== 'string' || str === '') return str;
  
  const turkishMap = {
    'ı': 'I', 'i': 'İ', 'ç': 'Ç', 'ş': 'Ş', 'ğ': 'Ğ', 'ü': 'Ü', 'ö': 'Ö',
  };
  return str.replace(/[ıiçşğüö]/g, match => turkishMap[match]).toUpperCase();
}

/**
 * Sadece belirtilen satırdaki verileri Türkçe uyumlu büyük harfe çevirir (Optimize edilmiş)
 * @param {number} rowNumber - İşlenecek satır numarası
 * @param {string} sheetName - İşlenecek sayfa adı (opsiyonel)
 */
function convertRowToUpperCase(rowNumber, sheetName = null) {
  try {
    Logger.log(`convertRowToUpperCase başladı - Satır: ${rowNumber}`);
    
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = sheetName ? 
      spreadsheet.getSheetByName(sheetName) : 
      spreadsheet.getActiveSheet();
    
    if (!sheet) {
      Logger.log(`HATA: ${sheetName || 'Aktif sayfa'} bulunamadı.`);
      return;
    }
    
    if (rowNumber <= 1) {
      Logger.log("Başlık satırı işlenmedi.");
      return;
    }
    
    const range = sheet.getRange(rowNumber, 1, 1, sheet.getLastColumn());
    const values = range.getValues();
    let changedCount = 0;
    
    for (let j = 0; j < values[0].length; j++) {
      const originalValue = values[0][j];
      
      // Sadece string, boş olmayan ve email olmayan değerleri işle
      if (typeof originalValue === 'string' && 
          originalValue !== '' && 
          !originalValue.includes('@')) {
        
        const convertedValue = convertToTurkishUpperCase(originalValue);
        
        if (originalValue !== convertedValue) {
          values[0][j] = convertedValue;
          changedCount++;
        }
      }
    }
    
    if (changedCount > 0) {
      range.setValues(values);
      Logger.log(`Satır ${rowNumber}: ${changedCount} hücre büyük harfe çevrildi.`);
    } else {
      Logger.log(`Satır ${rowNumber}: Değişiklik gerektiren hücre bulunamadı.`);
    }
    
  } catch (error) {
    Logger.log(`convertRowToUpperCase hatası (Satır ${rowNumber}): ` + error.message);
    throw error;
  }
}

/**
 * Tüm sayfadaki verileri Türkçe karakterlere uygun büyük harfe çevirir
 * (Manuel kullanım için - otomatik tetikleyicide kullanmayın!)
 */
function convertAllToUpperCase() {
  try {
    Logger.log("convertAllToUpperCase başladı - TÜM VERİLER işleniyor...");
    
    const sheetNames = ['EK-2 BİLGİLER', 'EK-2_DB'];
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    
    sheetNames.forEach(sheetName => {
      const sheet = spreadsheet.getSheetByName(sheetName);
      if (!sheet) {
        Logger.log(`UYARI: ${sheetName} sayfası bulunamadı.`);
        return;
      }
      
      const range = sheet.getDataRange();
      const values = range.getValues();
      let changedCount = 0;
      
      // Başlık satırını atla (i = 1'den başla)
      for (let i = 1; i < values.length; i++) {
        for (let j = 0; j < values[i].length; j++) {
          const originalValue = values[i][j];
          
          // Sadece string, boş olmayan ve email olmayan değerleri işle
          if (typeof originalValue === 'string' && 
              originalValue !== '' && 
              !originalValue.includes('@')) {
            
            const convertedValue = convertToTurkishUpperCase(originalValue);
            
            if (originalValue !== convertedValue) {
              values[i][j] = convertedValue;
              changedCount++;
            }
          }
        }
      }
      
      if (changedCount > 0) {
        range.setValues(values);
        Logger.log(`${sheetName}: ${changedCount} hücre büyük harfe çevrildi.`);
      } else {
        Logger.log(`${sheetName}: Değişiklik gerektiren hücre bulunamadı.`);
      }
    });
    
    Logger.log("convertAllToUpperCase tamamlandı.");
  } catch (error) {
    Logger.log("convertAllToUpperCase hatası: " + error.message);
    throw error;
  }
}

// ================================
// ARŞİVLEME FONKSİYONU
// ================================
function archiveNewFormResponses() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetA = ss.getSheetByName('EK-2 BİLGİLER'); 
  const sheetB = ss.getSheetByName('EK-2_DB'); 

  if (!sheetA || !sheetB) {
    Logger.log("⚠️ Bir veya daha fazla sayfa mevcut değil.");
    return;
  }

  const tcColumnIndex = 5; // E sütunu (TC Kimlik)
  const dateColumnIndex = 1; // A sütunu (Tarih)
  
  const dataB = sheetB.getDataRange().getValues();
  const dataA = sheetA.getDataRange().getValues();
  
  Logger.log("✅ EK-2 BİLGİLER'den alınan satır sayısı: " + dataA.length);

  // Mevcut TC + Tarih kombinasyonlarını sakla
  const existingTCDateCombinations = new Set();
  for (let i = 1; i < dataB.length; i++) {
    const row = dataB[i];
    const tcNumber = row[tcColumnIndex - 1];
    const dateValue = row[dateColumnIndex - 1];
    
    // Tarihi DD/MM/YYYY formatına çevir
    let formattedDate = '';
    if (dateValue instanceof Date) {
      formattedDate = formatDateToDDMMYYYY(dateValue);
    } else if (typeof dateValue === 'string' && dateValue.includes('/')) {
      formattedDate = dateValue;
    }
    
    if (tcNumber && formattedDate) {
      const combination = `${tcNumber}_${formattedDate}`;
      existingTCDateCombinations.add(combination);
    }
  }

  const newUniqueRows = [];
  for (let i = 1; i < dataA.length; i++) {
    const row = dataA[i];
    const tcNumber = row[tcColumnIndex - 1];
    const dateValue = row[dateColumnIndex - 1];
    
    // Tarihi DD/MM/YYYY formatına çevir
    let formattedDate = '';
    if (dateValue instanceof Date) {
      formattedDate = formatDateToDDMMYYYY(dateValue);
    } else if (typeof dateValue === 'string' && dateValue.includes('/')) {
      formattedDate = dateValue;
    }
    
    if (tcNumber && formattedDate) {
      const combination = `${tcNumber}_${formattedDate}`;
      
      // TC + Tarih kombinasyonu yoksa ekle
      if (!existingTCDateCombinations.has(combination)) {
        newUniqueRows.push(row);
        existingTCDateCombinations.add(combination);
        Logger.log(`✅ Yeni kombinasyon eklendi: TC=${tcNumber}, Tarih=${formattedDate}`);
      } else {
        Logger.log(`ℹ️ Mevcut kombinasyon atlandı: TC=${tcNumber}, Tarih=${formattedDate}`);
      }
    }
  }

  Logger.log("✅ EK-2_DB'ye eklenecek yeni satır sayısı: " + newUniqueRows.length);
  
  if (newUniqueRows.length > 0) {
    const startRow = sheetB.getLastRow() + 1;
    sheetB.getRange(startRow, 1, newUniqueRows.length, newUniqueRows[0].length).setValues(newUniqueRows);
    Logger.log("✅ Yeni satırlar başarıyla eklendi.");
  } else {
    Logger.log("ℹ️ Yeni eklenen satır yok (tüm TC+Tarih kombinasyonları mevcut).");
  }
}
// ================================
// BELGE OLUŞTURMA FONKSİYONLARI
// ================================
function generateDocument(rowNumber, silentMode = false) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    if (!sheet) throw new Error('Tablo bulunamadı.');

    if (!rowNumber) rowNumber = sheet.getLastRow();

    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const data = sheet.getRange(rowNumber, 1, 1, sheet.getLastColumn()).getValues()[0];
    const responseData = headers.reduce((obj, header, i) => ({ ...obj, [header]: data[i] }), {});
    
    Logger.log('responseData oluşturuldu, satır: ' + rowNumber);

    const workplaceName = responseData['İş Yerinin Adı'] ? responseData['İş Yerinin Adı'].substring(0, 13) : 'Bilinmiyor';
    const firstName = responseData['Adı'] || 'Bilinmiyor';
    const lastName = responseData['Soyadı'] || 'Bilinmiyor';

    const fileName = `${workplaceName} ${firstName} ${lastName} - Muayene Formu`;
    const destinationFolder = DriveApp.getFolderById(DESTINATION_FOLDER_ID);
    const template = DriveApp.getFileById(TEMPLATE_FILE_ID);
    const newDoc = template.makeCopy(fileName, destinationFolder);
    const doc = DocumentApp.openById(newDoc.getId());
    const body = doc.getBody();

    fillPlaceholders(body, responseData);
    addFLISTData(responseData, body);
    addSignatureImage(responseData, body);

    doc.saveAndClose();
    Logger.log('Belge başarıyla oluşturuldu: ' + fileName);
  } catch (e) {
    Logger.log('generateDocument hatası: ' + e.message);
    if (!silentMode) {
      SpreadsheetApp.getUi().alert('Hata oluştu: ' + e.message);
    }
    throw e;
  }
}

// ================================
// YER İMLERİ VE VERİ İŞLEME
// ================================
function fillPlaceholders(body, data) {
  const yesNoPlaceholders = [
    'Son bir yıl içinde ve sürekli olarak aşağıdaki yakınmalardan herhangi birini  yaşadınız mı? [Balgamlı öksürük]',
    'Son bir yıl içinde ve sürekli olarak aşağıdaki yakınmalardan herhangi birini  yaşadınız mı? [Nefes darlığı]',
    'Son bir yıl içinde ve sürekli olarak aşağıdaki yakınmalardan herhangi birini  yaşadınız mı? [Göğüs ağrısı]',
    'Son bir yıl içinde ve sürekli olarak aşağıdaki yakınmalardan herhangi birini  yaşadınız mı? [Çarpıntı]',
    'Son bir yıl içinde ve sürekli olarak aşağıdaki yakınmalardan herhangi birini  yaşadınız mı? [Sırt ağrısı]',
    'Son bir yıl içinde ve sürekli olarak aşağıdaki yakınmalardan herhangi birini  yaşadınız mı? [İshal veya kabızlık]',
    'Son bir yıl içinde ve sürekli olarak aşağıdaki yakınmalardan herhangi birini  yaşadınız mı? [Eklemlerde ağrı]',
    'Son bir yıl içinde ve sürekli olarak aşağıdaki yakınmalardan herhangi birini  yaşadınız mı? [Bayılma]',
    'Son bir yıl içinde ve sürekli olarak aşağıdaki yakınmalardan herhangi birini  yaşadınız mı? [Allerji]',
    'Aşağıdaki hastalıklardan herhangi birini geçirdiniz mi? [Kalp hastalığı]',
    'Aşağıdaki hastalıklardan herhangi birini geçirdiniz mi? [Şeker hastalığı]',
    'Aşağıdaki hastalıklardan herhangi birini geçirdiniz mi? [Böbrek rahatsızlığı]',
    'Aşağıdaki hastalıklardan herhangi birini geçirdiniz mi? [Sarılık]',
    'Aşağıdaki hastalıklardan herhangi birini geçirdiniz mi? [Mide veya on iki parmak ülseri]',
    'Aşağıdaki hastalıklardan herhangi birini geçirdiniz mi? [İşitme kaybı]',
    'Aşağıdaki hastalıklardan herhangi birini geçirdiniz mi? [Görme bozukluğu]',
    'Aşağıdaki hastalıklardan herhangi birini geçirdiniz mi? [Sinir sistemi hastalığı]',
    'Aşağıdaki hastalıklardan herhangi birini geçirdiniz mi? [Deri hastalığı]',
    'Aşağıdaki hastalıklardan herhangi birini geçirdiniz mi? [Besin zehirlenmesi]',
    'Hastanede yattınız mı?',
    'Ameliyat Oldunuz mu?',
    'İş kazası geçirdiniz mi?',
    'Sigara içiyor musunuz?',
    'Alkol alıyor musunuz?',
  ];

  // Evet/Hayır/Bırakmış kontrolleri
  yesNoPlaceholders.forEach((placeholder) => {
    const yesKey = `{{${placeholder} - EVET}}`;
    const noKey = `{{${placeholder} - HAYIR}}`;
    const quitKey = `{{${placeholder} - BIRAKMIŞ}}`;

    if (data[placeholder] === 'EVET') {
      body.replaceText(escapeRegExp(yesKey), '✓');
      body.replaceText(escapeRegExp(noKey), '');
      body.replaceText(escapeRegExp(quitKey), '');
    } else if (data[placeholder] === 'HAYIR') {
      body.replaceText(escapeRegExp(yesKey), '');
      body.replaceText(escapeRegExp(noKey), '✓');
      body.replaceText(escapeRegExp(quitKey), '');
    } else if (data[placeholder] === 'BIRAKMIŞ') {
      body.replaceText(escapeRegExp(yesKey), '');
      body.replaceText(escapeRegExp(noKey), '');
      body.replaceText(escapeRegExp(quitKey), '✓');
    } else {
      body.replaceText(escapeRegExp(yesKey), '');
      body.replaceText(escapeRegExp(noKey), '');
      body.replaceText(escapeRegExp(quitKey), '');
    }
  });

  // Tüm yer imleri için döngü
  for (const key in data) {
    let placeholder = `{{${key}}}`;
    let value = data[key];

    if (value instanceof Date) {
      value = formatDateToDDMMYYYY(value);
    }
    body.replaceText(escapeRegExp(placeholder), String(value || ''));
  }
}

// ================================
// YARDIMCI FONKSİYONLAR
// ================================
/**
 * Düzenli ifadelerde özel anlam taşıyan karakterleri "kaçış" (escape) karakteri ekleyerek güvenli hale getirir.
 */
function escapeRegExp(str) {
  return str.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

/**
 * Date nesnesini 'dd/MM/yyyy' formatında bir string'e dönüştürür.
 */
function formatDateToDDMMYYYY(date) {
  if (!date) return '';
  const d = new Date(date);
  const day = String(d.getDate()).padStart(2, '0');
  const month = String(d.getMonth() + 1).padStart(2, '0');
  const year = d.getFullYear();
  return `${day}/${month}/${year}`;
}

/**
 * 'dd/MM/yyyy' formatında bir string'i Date nesnesine dönüştürür.
 */
function parseDateFromDDMMYYYY(dateString) {
  if (!dateString) return null;
  const parts = dateString.split('/');
  return new Date(parts[2], parts[1] - 1, parts[0]);
}

/**
 * Belgeye FLIST sayfasındaki verileri ekler.
 */
function addFLISTData(data, body) {
  const flistSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('FLIST');
  if (!flistSheet) {
    Logger.log('FLIST sayfası bulunamadı');
    return;
  }

  const flistData = flistSheet.getDataRange().getValues();
  const headers = flistData[0];
  const rows = flistData.slice(1);
  
  // İş yerinin adını daha esnek kontrol etmek için
  const matchingRow = rows.find(row => 
    row[headers.indexOf('UNVANI')] && 
    data['İş Yerinin Adı'] &&
    row[headers.indexOf('UNVANI')].slice(0, 5).toLowerCase() === 
    data['İş Yerinin Adı'].slice(0, 5).toLowerCase()
  );

  if (matchingRow) {
    Logger.log('FLIST eşleşmesi bulundu');
    headers.forEach((header, index) => {
      const placeholder = `{{${header}}}`;
      const value = matchingRow[index] || 'Veri bulunamadı';
      body.replaceText(escapeRegExp(placeholder), String(value));
    });
  } else {
    Logger.log('FLIST eşleşmesi bulunamadı');
  }
}

/**
 * Belgeye IH_IMZA verisine göre imza görseli ekler.
 */
function addSignatureImage(responseData, body) {
  const imzaUrl = responseData.IH_IMZA;
  if (!imzaUrl) {
    Logger.log('İmza URL\'si bulunamadı');
    return;
  }

  const fileId = getFileIdFromUrl(imzaUrl);
  if (!fileId) {
    Logger.log('İmza dosya ID\'si çıkarılamadı');
    return;
  }

  try {
    const file = DriveApp.getFileById(fileId);
    const imageBlob = file.getBlob();

    const searchResult = body.findText('{{IH_IMZA}}');
    if (!searchResult) {
      Logger.log('İmza placeholder\'ı bulunamadı');
      return;
    }

    const foundElement = searchResult.getElement();
    const parent = foundElement.getParent();

    foundElement.asText().setText('');

    const insertedImage = body.insertImage(body.getChildIndex(parent) + 1, imageBlob);
    insertedImage.setWidth(120);
    insertedImage.setHeight(40);
    
    Logger.log('İmza başarıyla eklendi');
  } catch (error) {
    Logger.log('İmza ekleme hatası: ' + error.message);
  }
}

/**
 * URL'den dosya ID'sini çıkaran yardımcı fonksiyon
 */
function getFileIdFromUrl(url) {
  const match = url.match(/\/d\/([a-zA-Z0-9_-]+)/);
  if (match && match[1]) {
    return match[1];
  }
  return null;
}

// ================================
// UI VE TOPLU İŞLEM FONKSİYONLARI
// ================================
function openFilterAndGenerateDialog() {
  const htmlOutput = HtmlService.createHtmlOutputFromFile('FilterAndGenerateDialog')
      .setWidth(600)
      .setHeight(999);
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Personel Seç ve Belge Oluştur');
}

function getFilteredData(filters) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  if (!sheet) return { error: 'Aktif sayfa bulunamadı.' };
  
  const values = sheet.getDataRange().getValues();
  const headers = values[0];
  const bColIndex = headers.indexOf('İş Yerinin Adı');
  const cColIndex = headers.indexOf('Adı');
  const dColIndex = headers.indexOf('Soyadı');
  
  if (bColIndex === -1 || cColIndex === -1 || dColIndex === -1) {
    return { error: 'Gerekli sütun başlıkları (İş Yerinin Adı, Adı, Soyadı) bulunamadı.' };
  }
  
  const bFilter = filters.bFilter ? filters.bFilter.trim().toLowerCase() : '';
  const cdFilter = filters.cdFilter ? filters.cdFilter.trim().toLowerCase() : '';
  const selectedWorkplaces = filters.selectedWorkplaces || [];

  const filteredRows = values.slice(1).filter(row => {
    const bValue = String(row[bColIndex]).trim();
    const bValueLower = bValue.toLowerCase();
    const cdValue = `${String(row[cColIndex]).trim()} ${String(row[dColIndex]).trim()}`.toLowerCase();

    const bMatch = !bFilter || bValueLower.includes(bFilter);
    const cdMatch = !cdFilter || cdValue.includes(cdFilter);
    const workplaceSelectionMatch = selectedWorkplaces.length === 0 || selectedWorkplaces.includes(bValue);

    return bMatch && cdMatch && workplaceSelectionMatch;
  });
  
  const filteredWorkplaces = [...new Set(filteredRows.map(row => String(row[bColIndex]).trim()))].filter(v => v).sort();
  const filteredCombinedNames = [...new Set(filteredRows.map(row => `${String(row[cColIndex]).trim()} ${String(row[dColIndex]).trim()}`))].filter(v => v.trim()).sort();
  
  return { 
    'İş Yerinin Adı': filteredWorkplaces,
    'Adı Soyadı': filteredCombinedNames 
  };
}

function generateDocumentsFromFilteredRows(selections) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  if (!sheet) {
    return { success: false, message: 'Tablo bulunamadı.' };
  }

  const allData = sheet.getDataRange().getValues();
  const headers = allData[0];
  const bColIndex = headers.indexOf('İş Yerinin Adı');
  const cColIndex = headers.indexOf('Adı');
  const dColIndex = headers.indexOf('Soyadı');

  if (bColIndex === -1 || cColIndex === -1 || dColIndex === -1) {
    return { success: false, message: 'Gerekli sütun başlıkları (İş Yerinin Adı, Adı, Soyadı) bulunamadı.' };
  }

  const rowsToGenerate = allData.map((row, index) => {
    if (index === 0) return -1; // Başlık satırını atla

    const combinedCDValue = `${String(row[cColIndex]).trim()} ${String(row[dColIndex]).trim()}`.trim();

    if (selections.selectedNames.includes(combinedCDValue)) {
      return index + 1;
    }
    return -1;
  }).filter(rowNum => rowNum !== -1);
  
  if (rowsToGenerate.length === 0) {
    return { success: false, message: 'Seçili kriterlere uyan satır bulunamadı.' };
  }

  let createdCount = 0;
  let failedCount = 0;
  
  rowsToGenerate.forEach((rowNumber) => {
    try {
      generateDocument(rowNumber, true);
      createdCount++;
    } catch (e) {
      Logger.log(`Satır ${rowNumber} için belge oluşturulurken hata: ` + e.message);
      failedCount++;
    }
  });
  
  return { 
    success: true, 
    createdCount: createdCount, 
    failedCount: failedCount,
    totalCount: rowsToGenerate.length
  };
}

function createSingleDocument() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const activeRangeList = sheet.getActiveRangeList();
  const uniqueRowNumbers = new Set();
  
  if (!activeRangeList) {
    SpreadsheetApp.getUi().alert('⚠️ Lütfen belge oluşturmak için en az bir satır seçin.');
    return;
  }

  // Seçili tüm aralıklardaki satır numaralarını al
  activeRangeList.getRanges().forEach(range => {
    for (let i = 0; i < range.getNumRows(); i++) {
      uniqueRowNumbers.add(range.getRow() + i);
    }
  });

  if (uniqueRowNumbers.size === 0) {
    SpreadsheetApp.getUi().alert('⚠️ Lütfen belge oluşturmak için en az bir satır seçin.');
    return;
  }

  let createdCount = 0;
  let failedCount = 0;

  uniqueRowNumbers.forEach((rowNumber) => {
    try {
      if (rowNumber > 1) { // Başlık satırını atla
        generateDocument(rowNumber, true);
        createdCount++;
      }
    } catch (e) {
      Logger.log(`Satır ${rowNumber} için belge oluşturulurken hata: ` + e.message);
      failedCount++;
    }
  });

  if (createdCount > 0) {
    SpreadsheetApp.getUi().alert(`✅ ${createdCount} adet belge başarıyla oluşturuldu!`);
  } else if (failedCount > 0) {
    SpreadsheetApp.getUi().alert(`⚠️ Belge oluşturma işlemi tamamlandı. ${failedCount} adet hata oluştu.`);
  } else {
    SpreadsheetApp.getUi().alert('ℹ️ Seçili satırlardan belge oluşturulamadı.');
  }
}